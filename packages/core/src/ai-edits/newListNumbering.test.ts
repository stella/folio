/**
 * `numbering: { kind: "newList", format }` (issue #1092): an operation can start a
 * list in a package with no numbering part, or a second list beside an
 * existing one, and the saved package defines every instance it references.
 */

import { describe, expect, test } from "bun:test";

import { paragraphNumberingReferenceId } from "@stll/docx-core/model";

import {
  parseFolioDocumentOperationBatch,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const reviewerOf = async (markdown: string): Promise<FolioDocxReviewer> => {
  const { docx } = await ensureParaIds(await createDocx(fromMarkdown(markdown)));
  return FolioDocxReviewer.fromBuffer(docx, { author: "Agent" });
};

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads "${text}"`);
  }
  return block.id;
};

const listInstancesByText = async (
  reviewer: FolioDocxReviewer,
): Promise<Record<string, number>> => {
  const reparsed = await parseDocx(await reviewer.toBuffer(), {
    preloadFonts: false,
    detectVariables: false,
  });
  const instances: Record<string, number> = {};
  for (const block of reparsed.package.document.content) {
    if (block.type !== "paragraph") continue;
    const numId = paragraphNumberingReferenceId(block.formatting?.numPr);
    const text = block.content
      .flatMap((item) =>
        item.type === "run"
          ? item.content.flatMap((part) => (part.type === "text" ? [part.text] : []))
          : [],
      )
      .join("");
    if (numId !== undefined) instances[text] = numId;
  }
  return instances;
};

const twoNewLists = (reviewer: FolioDocxReviewer): FolioDocumentOperation[] => [
  {
    id: "left",
    type: "insertAfterBlock",
    blockId: blockId(reviewer, "Left."),
    text: "Left list",
    numbering: { kind: "newList", format: "bullet" },
  },
  {
    id: "right",
    type: "insertAfterBlock",
    blockId: blockId(reviewer, "Right."),
    text: "Right list",
    numbering: { kind: "newList", format: "bullet" },
  },
];

/** Apply, save, check every reference is defined, and read the labels back. */
const savedLabels = async (
  reviewer: FolioDocxReviewer,
  operations: FolioDocumentOperation[],
): Promise<string[]> => {
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.status).toBe("committed");
  expect(result.issues).toEqual([]);

  const saved = await reviewer.toBuffer();
  const reparsed = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
  const defined = new Set((reparsed.package.numbering?.nums ?? []).map(({ numId }) => numId));
  for (const block of reparsed.package.document.content) {
    if (block.type !== "paragraph") continue;
    const numId = paragraphNumberingReferenceId(block.formatting?.numPr);
    if (numId !== undefined) expect(defined.has(numId)).toBe(true);
  }
  return (await FolioDocxReviewer.fromBuffer(saved))
    .getContent()
    .map(({ text, displayLabel }) => `${displayLabel ?? ""} ${text}`.trim());
};

describe("an operation that starts a new list", () => {
  test("defines the list in a package without a numbering part", async () => {
    const reviewer = await reviewerOf("Intro.\n\nTail.");
    const labels = await savedLabels(reviewer, [
      {
        id: "list",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Intro."),
        text: "First\nSecond",
        formattingScope: "allParagraphs",
        numbering: { kind: "newList", format: "numbered" },
      },
    ]);

    expect(labels).toEqual(["Intro.", "1. First", "2. Second", "Tail."]);
  });

  test("starts a second list at one without renumbering the first", async () => {
    const reviewer = await reviewerOf("1. Alpha\n2. Beta\n\nTail.");
    const labels = await savedLabels(reviewer, [
      {
        id: "list",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Tail."),
        text: "Gamma\nDelta",
        formattingScope: "allParagraphs",
        numbering: { kind: "newList", format: "numbered" },
      },
    ]);

    expect(labels).toEqual(["1. Alpha", "2. Beta", "Tail.", "1. Gamma", "2. Delta"]);
  });

  test("starts a bullet list", async () => {
    const reviewer = await reviewerOf("Intro.");
    const labels = await savedLabels(reviewer, [
      {
        id: "list",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Intro."),
        text: "First\nSecond",
        formattingScope: "allParagraphs",
        numbering: { kind: "newList", format: "bullet" },
      },
    ]);

    expect(labels).toEqual(["Intro.", "• First", "• Second"]);
  });

  test("allocates a batch in the same right-to-left order as one-at-a-time replay", async () => {
    const batch = await reviewerOf("Left.\n\nMiddle.\n\nRight.");
    const replay = await reviewerOf("Left.\n\nMiddle.\n\nRight.");
    const operations = twoNewLists(batch);
    batch.applyDocumentOperations({ version: 1, mode: "direct", operations }, {});

    const replayOperations = twoNewLists(replay);
    for (const operation of [...replayOperations].reverse()) {
      replay.applyDocumentOperations({ version: 1, mode: "direct", operations: [operation] }, {});
    }

    expect(await listInstancesByText(batch)).toEqual(await listInstancesByText(replay));
  });

  test("allocates direct and accepted tracked batches alike after save and reopen", async () => {
    const direct = await reviewerOf("Left.\n\nMiddle.\n\nRight.");
    const tracked = await reviewerOf("Left.\n\nMiddle.\n\nRight.");
    direct.applyDocumentOperations(
      { version: 1, mode: "direct", operations: twoNewLists(direct) },
      {},
    );
    tracked.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: twoNewLists(tracked),
      },
      {},
    );
    tracked.acceptAll();

    const reopenedDirect = await FolioDocxReviewer.fromBuffer(await direct.toBuffer());
    const reopenedTracked = await FolioDocxReviewer.fromBuffer(await tracked.toBuffer());
    expect(await listInstancesByText(reopenedTracked)).toEqual(
      await listInstancesByText(reopenedDirect),
    );
  });

  test("numbers existing paragraphs, one list per operation", async () => {
    const reviewer = await reviewerOf("One\n\nProse.\n\nTwo");
    const labels = await savedLabels(reviewer, [
      {
        id: "first",
        type: "setBlockParagraphProperties",
        blockId: blockId(reviewer, "One"),
        properties: { numbering: { kind: "newList", format: "numbered" } },
      },
      {
        id: "second",
        type: "setBlockParagraphProperties",
        blockId: blockId(reviewer, "Two"),
        properties: { numbering: { kind: "newList", format: "numbered", level: 0 } },
      },
    ]);

    expect(labels).toEqual(["1. One", "Prose.", "1. Two"]);
  });

  test("does not allocate for paragraph properties refused by a pending mark deletion", async () => {
    const prepared = await reviewerOf("Left.\n\nRefused.\n\nTail.");
    const leftBlockId = blockId(prepared, "Left.");
    const refusedBlockId = blockId(prepared, "Refused.");
    const deletion = prepared.applyDocumentOperations(
      {
        version: 1,
        mode: "tracked-changes",
        operations: [
          {
            id: "delete-refused-mark",
            type: "deleteBlock",
            blockId: refusedBlockId,
          },
        ],
      },
      {},
    );
    expect(deletion.status).toBe("committed");

    const source = await prepared.toBuffer();
    const batch = await FolioDocxReviewer.fromBuffer(source, { author: "Agent" });
    const sequential = await FolioDocxReviewer.fromBuffer(source, { author: "Agent" });
    const operations: FolioDocumentOperation[] = [
      {
        id: "refused-list",
        type: "setBlockParagraphProperties",
        blockId: refusedBlockId,
        properties: { numbering: { kind: "newList", format: "numbered" } },
      },
      {
        id: "persisted-list",
        type: "insertAfterBlock",
        blockId: leftBlockId,
        text: "Persisted list",
        numbering: { kind: "newList", format: "numbered" },
      },
    ];

    const batchResult = batch.applyDocumentOperations(
      { version: 1, mode: "tracked-changes", operations },
      {},
    );
    const refusedResult = sequential.applyDocumentOperations(
      { version: 1, mode: "tracked-changes", operations: [operations[0]!] },
      {},
    );
    sequential.applyDocumentOperations(
      { version: 1, mode: "tracked-changes", operations: [operations[1]!] },
      {},
    );

    expect(batchResult.skipped).toEqual([
      expect.objectContaining({ id: "refused-list", reason: "pendingParagraphMarkDeletion" }),
    ]);
    expect(refusedResult.skipped).toEqual(batchResult.skipped);
    expect((await listInstancesByText(batch))["Refused."]).toBeUndefined();
    expect(await listInstancesByText(batch)).toEqual(await listInstancesByText(sequential));
  });

  test.each([
    [{ kind: "newList", format: "roman" }, "$.operations[0].properties.numbering.format"],
    [
      { kind: "newList", format: "numbered", level: 9 },
      "$.operations[0].properties.numbering.level",
    ],
    [{ kind: "newList", format: "bullet", numId: 3 }, "$.operations[0].properties.numbering.numId"],
    [
      { kind: "newList", format: "bullet", start: "old" },
      "$.operations[0].properties.numbering.start",
    ],
  ])("rejects %j", (numbering, path) => {
    expect(() =>
      parseFolioDocumentOperationBatch({
        version: 1,
        operations: [
          { id: "1", type: "setBlockParagraphProperties", blockId: "a", properties: { numbering } },
        ],
      }),
    ).toThrow(expect.objectContaining({ _tag: "InvalidFolioDocumentOperationBatchError", path }));
  });
});
