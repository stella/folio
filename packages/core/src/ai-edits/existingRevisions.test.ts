import { expect, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationBatch,
} from "../document-operations";
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

for (const session of ["legacy", "canonical"] as const) {
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
