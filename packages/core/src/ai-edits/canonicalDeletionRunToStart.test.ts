/**
 * A batch's operations each run against the paragraphs as the batch has
 * left them: a run deleted back to its container's start collapses right to
 * left onto its first paragraph; a property edit runs before the deletion
 * that takes its paragraph's break; and an edit of a paragraph an earlier
 * operation joined away edits the paragraph that now holds it.
 */
import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const idOf = (index: number) =>
  (0x1000_0000 + index).toString(16).toUpperCase();
const deletion = (index: number): FolioDocumentOperation => ({
  id: `delete-${String(index)}`,
  type: "deleteBlock",
  blockId: idOf(index),
});

const open = async (texts: readonly string[]) =>
  FolioDocxReviewer.fromBuffer(
    await paragraphsDocx(texts.map((text) => [textRun(text)])),
    { session: "canonical" },
  );

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );

test("a direct batch deleting every paragraph keeps the first, emptied, in any order", async () => {
  for (const run of [
    [0, 1, 2],
    [2, 1, 0],
    [1, 2, 0],
  ]) {
    const reviewer = await open(["P0.", "P1.", "P2."]);
    expect(apply(reviewer, "direct", run.map(deletion)).skipped).toEqual([]);
    expect(reviewer.getContent().map(({ id, text }) => ({ id, text }))).toEqual(
      [{ id: idOf(0), text: "" }],
    );
  }
});

// Tracked, a run of two is refused: deleting its first paragraph after the
// last paragraph's deletion took that paragraph's properties.
test("a batch deleting every paragraph of three settles on the first one's properties, tracked as directly", async () => {
  for (const mode of ["direct", "tracked-changes"] as const) {
    for (const run of [
      [0, 1, 2],
      [2, 1, 0],
      [1, 2, 0],
    ]) {
      const texts = run.map((_, index) => `P${String(index)}.`);
      const reviewer = await open(texts);
      expect(
        apply(reviewer, "direct", [
          {
            id: "restyle",
            type: "setBlockParagraphProperties",
            blockId: idOf(0),
            properties: { alignment: "center" },
          },
        ]).skipped,
      ).toEqual([]);
      expect(apply(reviewer, mode, run.map(deletion)).skipped).toEqual([]);
      expect(reviewer.acceptAll().skipped).toEqual([]);
      expect(
        reviewer.getContent().map(({ text, directAlignment }) => ({
          text,
          directAlignment,
        })),
      ).toEqual([{ text: "", directAlignment: "center" }]);
    }
  }
});

test("a tracked property edit runs before the deletion that retires its paragraph's break", async () => {
  for (const removed of [1, 2]) {
    const texts = Array.from(
      { length: removed + 2 },
      (_, index) => `P${String(index)}.`,
    );
    const reviewer = await open(texts);
    const deletions = texts.slice(2).map((_, index) => deletion(index + 2));
    const result = apply(reviewer, "tracked-changes", [
      {
        id: "restyle",
        type: "setBlockParagraphProperties",
        blockId: idOf(1),
        properties: { alignment: "center" },
      },
      ...deletions,
    ]);
    expect(result.skipped).toEqual([]);
    expect(reviewer.acceptAll().skipped).toEqual([]);
    expect(
      reviewer.getContent().map(({ text, directAlignment }) => ({
        text,
        directAlignment,
      })),
    ).toEqual([
      { text: "P0.", directAlignment: undefined },
      { text: "P1.", directAlignment: "center" },
    ]);
  }
});

test("an edit of a split's first half runs in the paragraph its retraction leaves", async () => {
  for (const edit of ["merge", "properties"] as const) {
    const reviewer = await open(["P0.", "P1.", "Closing remarks."]);
    expect(
      apply(reviewer, "tracked-changes", [
        {
          id: "split",
          type: "splitBlock",
          blockId: idOf(2),
          offset: "Closing ".length,
        },
      ]).skipped,
    ).toEqual([]);
    const second = reviewer.getContent().at(-1);
    if (second === undefined) throw new Error("The split has no second half");
    const operations: FolioDocumentOperation[] = [
      edit === "merge"
        ? {
            id: "edit",
            type: "mergeBlockWithNext",
            blockId: idOf(1),
            separator: " ",
          }
        : {
            id: "edit",
            type: "setBlockParagraphProperties",
            blockId: idOf(2),
            properties: { alignment: "center" },
          },
      { id: "retract", type: "deleteBlock", blockId: second.id },
    ];
    expect(apply(reviewer, "tracked-changes", operations).skipped).toEqual([]);
    expect(reviewer.acceptAll().skipped).toEqual([]);
    expect(
      reviewer.getContent().map(({ text, directAlignment }) => ({
        text,
        directAlignment,
      })),
    ).toEqual(
      edit === "merge"
        ? [
            { text: "P0.", directAlignment: undefined },
            { text: "P1. Closing ", directAlignment: undefined },
          ]
        : [
            { text: "P0.", directAlignment: undefined },
            { text: "P1.", directAlignment: undefined },
            { text: "Closing ", directAlignment: "center" },
          ],
    );
  }
});
