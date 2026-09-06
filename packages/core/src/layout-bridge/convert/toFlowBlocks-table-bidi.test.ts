import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import type { TableBlock } from "../../layout-engine/types";
import { schema } from "../../prosemirror/schema";
import { toFlowBlocks } from "./toFlowBlocks";

const cell = (text: string): PMNode =>
  schema.node("tableCell", { borders: null }, [
    schema.node("paragraph", null, [schema.text(text)]),
  ]);

const tableDoc = (originalFormatting: unknown): PMNode =>
  schema.node("doc", null, [
    schema.node("table", { columnWidths: [2000, 3000], _originalFormatting: originalFormatting }, [
      schema.node("tableRow", null, [cell("A"), cell("B")]),
    ]),
  ]);

const firstTable = (doc: PMNode): TableBlock => {
  const block = toFlowBlocks(doc).find((b) => b.kind === "table");
  if (!block) {
    throw new Error("Expected a table block");
  }
  return block;
};

describe("toFlowBlocks table bidiVisual", () => {
  test("carries w:bidiVisual onto the table block (eigenpal/docx-editor#940)", () => {
    expect(firstTable(tableDoc({ bidi: true })).bidi).toBe(true);
  });

  test("leaves bidi unset for an ordinary LTR table", () => {
    expect(firstTable(tableDoc(null)).bidi).toBeUndefined();
    expect(firstTable(tableDoc({})).bidi).toBeUndefined();
  });
});

describe("toFlowBlocks table indent compatibility", () => {
  test("carries the document's text-edge indent policy onto every table block", () => {
    const block = toFlowBlocks(tableDoc(null), {
      tableIndentCompatibility: { type: "legacy" },
    }).find((b) => b.kind === "table");

    expect(block?.indentCompatibility).toEqual({ type: "legacy" });
  });

  test("leaves the policy unset when the document uses border-edge indents", () => {
    expect(firstTable(tableDoc(null)).indentCompatibility).toBeUndefined();
  });
});
