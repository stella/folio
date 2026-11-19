import { expect, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationBatch,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

const apply = (
  reviewer: FolioDocxReviewer,
  operations: readonly FolioDocumentOperation[],
  options: Pick<
    FolioDocumentOperationBatch,
    "existingRevisions" | "atomic"
  > = {},
) =>
  reviewer.applyDocumentOperations(
    { version: 1, mode: "tracked-changes", operations, ...options },
    { undefinedReferences: "refuse" },
  );

const revisedDocument = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await paragraphsDocx([[textRun("Start old end.")], [textRun("Neighbor.")]]),
    { author: "First Reviewer" },
  );
  expect(
    apply(reviewer, [
      {
        id: "first",
        type: "replaceInBlock",
        blockId: "10000000",
        find: "old",
        replace: "new",
      },
    ]).applied,
  ).toHaveLength(1);
  return await reviewer.toBuffer();
};

const POLICY = { existingRevisions: "refuse-cross-author" } as const;
const replace = {
  id: "replace",
  type: "replaceInBlock",
  blockId: "10000000",
  find: "new",
  replace: "fresh",
} as const;
const outside = {
  id: "neighbor",
  type: "replaceInBlock",
  blockId: "10000001",
  find: "Neighbor",
  replace: "Next",
} as const;

test("rejects a held batch when deleting an intervening paragraph exposes a foreign break", async () => {
  const docx = await createDocx({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "10000000",
            content: [textRun("First")],
            pPrMark: {
              kind: "ins",
              info: { id: 1, author: "First Reviewer" },
            },
          },
          {
            type: "paragraph",
            paraId: "10000001",
            content: [textRun("Middle")],
          },
          {
            type: "paragraph",
            paraId: "10000002",
            content: [textRun("Final")],
          },
        ],
      },
    },
  });
  const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
    author: "Second Reviewer",
    session: "legacy",
  });
  const result = apply(
    reviewer,
    [
      { id: "middle", type: "deleteBlock", blockId: "10000001" },
      { id: "final", type: "deleteBlock", blockId: "10000002" },
    ],
    POLICY,
  );
  expect(result.status).toBe("rejected");
  expect(result.applied).toEqual([]);
  expect(result.skipped).toEqual([
    { id: "middle", reason: "atomicBatchRejected" },
    expect.objectContaining({
      id: "final",
      reason: "existingRevisionConflict",
    }),
  ]);
  expect(reviewer.getContent().map(({ text }) => text)).toEqual([
    "First",
    "Middle",
    "Final",
  ]);
});

for (const session of ["legacy", "canonical"] as const) {
  test(`${session} refuses a final deletion that retracts another author's break`, async () => {
    const docx = await createDocx({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "10000000",
              content: [textRun("First")],
              pPrMark: {
                kind: "ins",
                info: { id: 1, author: "First Reviewer" },
              },
            },
            {
              type: "paragraph",
              paraId: "10000001",
              content: [textRun("Final")],
            },
          ],
        },
      },
    });
    const reviewer = await FolioDocxReviewer.fromBuffer(docx, {
      author: "Second Reviewer",
      session,
    });
    const result = apply(
      reviewer,
      [{ id: "delete", type: "deleteBlock", blockId: "10000001" }],
      POLICY,
    );
    expect(result.applied).toEqual([]);
    if (session === "legacy") {
      expect(result.skipped.at(0)?.reason).toBe("existingRevisionConflict");
    }
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First",
      "Final",
    ]);
  });

  test.each([false, true])(
    `${session} protects a survivor's pending formatting when the carry changes it (%s)`,
    async (sameFormatting) => {
      const docx = await createDocx({
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "10000000",
                content: [textRun("First")],
                ...(sameFormatting && {
                  formatting: { alignment: "right" } as const,
                }),
              },
              {
                type: "paragraph",
                paraId: "10000001",
                content: [textRun("Second")],
              },
            ],
          },
        },
      });
      const source = await FolioDocxReviewer.fromBuffer(docx, {
        author: "First Reviewer",
        session,
      });
      expect(
        apply(source, [
          {
            id: "format",
            type: "setBlockParagraphProperties",
            blockId: "10000001",
            properties: { alignment: "right" },
          },
        ]).applied,
      ).toHaveLength(1);
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await source.toBuffer(),
        { author: "Second Reviewer", session },
      );
      const result = apply(
        reviewer,
        [{ id: "merge", type: "mergeBlockWithNext", blockId: "10000000" }],
        POLICY,
      );
      if (sameFormatting) {
        expect(result.applied).toHaveLength(1);
      } else {
        expect(result.skipped.at(0)?.reason).toBe("existingRevisionConflict");
        expect(result.applied).toEqual([]);
      }
    },
  );

  test(`${session} refuses another author's range and keeps independent edits`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await revisedDocument(),
      {
        author: "Second Reviewer",
        session,
      },
    );

    const atomic = apply(reviewer, [replace, outside], {
      ...POLICY,
      atomic: true,
    });
    expect(atomic.status).toBe("rejected");
    const result = apply(reviewer, [replace, outside], POLICY);
    expect(result.skipped).toEqual([
      expect.objectContaining({
        id: replace.id,
        reason: "existingRevisionConflict",
      }),
    ]);
    expect(result.applied.map(({ id }) => id)).toEqual([outside.id]);
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "Start new end.",
      "Next.",
    ]);
  });

  test(`${session} preserves default and explicit allow and admits the same author`, async () => {
    const document = await revisedDocument();
    for (const [author, options] of [
      ["Second Reviewer", {}],
      ["Second Reviewer", { existingRevisions: "allow" }],
      ["First Reviewer", POLICY],
    ] as const) {
      const reviewer = await FolioDocxReviewer.fromBuffer(document.slice(0), {
        author,
        session,
      });
      expect(apply(reviewer, [replace], options).applied).toHaveLength(1);
      expect(reviewer.getContent().at(0)?.text).toBe("Start fresh end.");
    }
  });

  test(`${session} protects formatting, paragraph deletion and the interior of a split`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await revisedDocument(),
      {
        author: "Second Reviewer",
        session,
      },
    );
    const edits = [
      {
        id: "format",
        type: "formatRange",
        range: {
          type: "textRange",
          story: "main",
          blockId: "10000000",
          startOffset: 6,
          endOffset: 9,
          selectedTextHash: hashFolioAIBlockText("new"),
        },
        formatting: { bold: true },
      },
      { id: "delete", type: "deleteBlock", blockId: "10000000" },
      { id: "split", type: "splitBlock", blockId: "10000000", offset: 7 },
    ] as const satisfies readonly FolioDocumentOperation[];
    for (const edit of edits) {
      expect(apply(reviewer, [edit], POLICY).skipped).toEqual([
        expect.objectContaining({
          id: edit.id,
          reason: "existingRevisionConflict",
        }),
      ]);
    }
    expect(
      apply(
        reviewer,
        [
          {
            id: "adjacent",
            type: "splitBlock",
            blockId: "10000000",
            offset: 9,
          },
        ],
        POLICY,
      ).applied,
    ).toHaveLength(1);
  });
}
