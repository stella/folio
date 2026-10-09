import assert from "node:assert/strict";
import { test } from "node:test";

import { modelPendingTextEffect } from "../support/pending-text-oracle.ts";

type Operation = Parameters<typeof modelPendingTextEffect>[0];
const blockId = "412DDC7F";
const nextId = "483CD5EB";
const range = {
  type: "textRange",
  story: "main",
  blockId,
  startOffset: 2,
  endOffset: 4,
  selectedTextHash: "unused-by-text-effect-model",
} as const;
const unchanged = [
  [blockId, "a 😀 b"],
  [nextId, "tail"],
] as const;
const tableEffect = "unsupportedTable";

// Every operation kind requires a fixture and an independent expected effect.
const cases = {
  replaceInBlock: {
    operation: { id: "replace", type: "replaceInBlock", blockId, find: "😀", replace: "é" },
    expected: [
      [blockId, "a é b"],
      [nextId, "tail"],
    ],
  },
  replaceRange: {
    operation: { id: "range", type: "replaceRange", range, replace: "é" },
    expected: [
      [blockId, "a é b"],
      [nextId, "tail"],
    ],
  },
  replaceBlock: {
    operation: { id: "block", type: "replaceBlock", blockId, text: "new" },
    expected: [
      [blockId, "new"],
      [nextId, "tail"],
    ],
  },
  deleteBlock: {
    operation: { id: "delete", type: "deleteBlock", blockId },
    expected: [[nextId, "tail"]],
  },
  splitBlock: {
    operation: { id: "split", type: "splitBlock", blockId, offset: 4, separator: " " },
    expected: [
      [blockId, "a 😀"],
      ["model:split:split", "b"],
      [nextId, "tail"],
    ],
  },
  mergeBlockWithNext: {
    operation: { id: "merge", type: "mergeBlockWithNext", blockId, separator: " " },
    expected: [[blockId, "a 😀 b tail"]],
  },
  insertBeforeBlock: {
    operation: { id: "before", type: "insertBeforeBlock", blockId, text: "first\n\nsecond" },
    expected: [["model:before:0", "first"], ["model:before:1", "second"], ...unchanged],
  },
  insertAfterBlock: {
    operation: {
      id: "after",
      type: "insertAfterBlock",
      blockId,
      text: "first\nsecond",
      lineBreakMode: "inline",
    },
    expected: [
      [blockId, "a 😀 b"],
      ["model:after:0", "first\nsecond"],
      [nextId, "tail"],
    ],
  },
  formatRange: {
    operation: { id: "format", type: "formatRange", range, formatting: { bold: true } },
    expected: unchanged,
  },
  commentOnRange: {
    operation: { id: "comment-range", type: "commentOnRange", range, comment: { text: "Note" } },
    expected: unchanged,
  },
  commentOnBlock: {
    operation: { id: "comment-block", type: "commentOnBlock", blockId, comment: { text: "Note" } },
    expected: unchanged,
  },
  setBlockParagraphProperties: {
    operation: {
      id: "properties",
      type: "setBlockParagraphProperties",
      blockId,
      properties: { styleId: "Heading2" },
    },
    expected: unchanged,
  },
  insertTable: {
    operation: { id: "table", type: "insertTable", blockId, rows: [["cell"]] },
    expected: tableEffect,
  },
  deleteTable: {
    operation: { id: "table", type: "deleteTable", blockId },
    expected: tableEffect,
  },
  insertSignatureTable: {
    operation: { id: "table", type: "insertSignatureTable", blockId, parties: [{ name: "Party" }] },
    expected: tableEffect,
  },
  insertTableRow: {
    operation: { id: "table", type: "insertTableRow", blockId, cellTexts: ["cell"] },
    expected: tableEffect,
  },
  deleteTableRow: {
    operation: { id: "table", type: "deleteTableRow", blockId },
    expected: tableEffect,
  },
  insertTableColumn: {
    operation: { id: "table", type: "insertTableColumn", blockId, cellTexts: ["cell"] },
    expected: tableEffect,
  },
  deleteTableColumn: {
    operation: { id: "table", type: "deleteTableColumn", blockId },
    expected: tableEffect,
  },
  mergeTableCells: {
    operation: { id: "table", type: "mergeTableCells", blockId, endBlockId: nextId },
    expected: tableEffect,
  },
  splitTableCell: {
    operation: { id: "table", type: "splitTableCell", blockId },
    expected: tableEffect,
  },
} as const satisfies Record<
  Operation["type"],
  { operation: Operation; expected: readonly (readonly [string, string])[] | "unsupportedTable" }
>;

for (const [kind, { operation, expected }] of Object.entries(cases)) {
  test(`pending text oracle models ${kind} or refuses table geometry explicitly`, () => {
    assert.equal(operation.type, kind);
    const model = new Map<string, string>(unchanged);
    if (expected === tableEffect) {
      assert.throws(
        () => modelPendingTextEffect(operation, model),
        /cannot model table operation/u,
      );
      assert.deepEqual([...model], unchanged);
      return;
    }
    modelPendingTextEffect(operation, model);
    assert.deepEqual([...model], expected);
  });
}

test("pending text effects compose over live paragraph order", () => {
  const model = new Map<string, string>(unchanged);
  modelPendingTextEffect(cases.splitBlock.operation, model);
  modelPendingTextEffect(cases.mergeBlockWithNext.operation, model);
  assert.deepEqual([...model], unchanged);
  modelPendingTextEffect(cases.insertBeforeBlock.operation, model);
  modelPendingTextEffect(cases.deleteBlock.operation, model);
  assert.equal(model.has(blockId), false);
  assert.deepEqual(
    [...model],
    [
      ["model:before:0", "first"],
      ["model:before:1", "second"],
      [nextId, "tail"],
    ],
  );
});
