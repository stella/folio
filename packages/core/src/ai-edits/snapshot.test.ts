/**
 * The snapshot's cost, and the table coordinates that pay for it.
 *
 * `createFolioAIEditSnapshot` walks the document once. Every fact it needs
 * about a block's surroundings — the row that hides it, the cell it sits in —
 * is on the path the walk already took, so it never resolves a position.
 * `doc.resolve` re-descends from the root and finds each level's child by
 * scanning that level's fragment from index 0, so one resolve per block is
 * O(blocks^2) on a flat document. The guard below is the invariant rather than
 * a stopwatch: the snapshot may not resolve at all.
 */

import { describe, expect, test } from "bun:test";
import { type Node as PMNode, Schema } from "prosemirror-model";

import { createFolioAIEditSnapshot } from "./snapshot";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: { paraId: { default: null } },
    },
    text: { group: "inline" },
    table: { content: "tableRow+", group: "block", tableRole: "table" },
    tableRow: {
      content: "tableCell+",
      tableRole: "row",
      attrs: { hidden: { default: false } },
    },
    tableCell: { content: "block+", tableRole: "cell" },
  },
});

const paragraph = (text: string): PMNode =>
  schema.node("paragraph", { paraId: null }, text.length === 0 ? [] : [schema.text(text)]);

const cell = (content: PMNode[]): PMNode => schema.node("tableCell", null, content);

const row = (cells: PMNode[], hidden = false): PMNode => schema.node("tableRow", { hidden }, cells);

const table = (rows: PMNode[]): PMNode => schema.node("table", null, rows);

/**
 * Make the document refuse to resolve a position, so a snapshot that reaches
 * for one fails here instead of quietly reintroducing the quadratic term.
 */
const refusingToResolve = (doc: PMNode): PMNode => {
  Object.defineProperty(doc, "resolve", {
    configurable: true,
    value: () => {
      throw new Error("The snapshot resolved a position; walk the document instead.");
    },
  });
  return doc;
};

describe("createFolioAIEditSnapshot", () => {
  test("locates a block in a table inside a table", () => {
    const nested = table([row([cell([paragraph("nested first"), paragraph("nested second")])])]);
    const doc = schema.node("doc", null, [
      paragraph("opening"),
      table([
        row([cell([paragraph("r0c0")]), cell([paragraph("r0c1")])]),
        row([cell([paragraph("r1c0")]), cell([paragraph("before nested"), nested])]),
      ]),
      paragraph("closing"),
    ]);

    const { blocks } = createFolioAIEditSnapshot(doc);

    expect(blocks.map(({ text }) => text)).toEqual([
      "opening",
      "r0c0",
      "r0c1",
      "r1c0",
      "before nested",
      "nested first",
      "nested second",
      "closing",
    ]);
    expect(blocks.at(0)?.table).toBeUndefined();
    expect(blocks.at(-1)?.table).toBeUndefined();
    expect(blocks.at(2)?.table).toEqual({
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex: 0,
      cellIndex: 1,
      paragraphIndex: 0,
    });
    // The nested table is the second table in document order, and its own
    // coordinates are relative to itself, not to the table containing it.
    // The nested table's own index, with the outer table it belongs to.
    expect(blocks.at(5)?.table).toEqual({
      outerTableIndex: 0,
      tableIndex: 1,
      rowIndex: 0,
      cellIndex: 0,
      paragraphIndex: 0,
    });
    expect(blocks.at(6)?.table).toEqual({
      outerTableIndex: 0,
      tableIndex: 1,
      rowIndex: 0,
      cellIndex: 0,
      paragraphIndex: 1,
    });
  });

  test("omits a hidden row's text, including a nested table inside it", () => {
    const doc = schema.node("doc", null, [
      table([
        row([cell([paragraph("visible")])]),
        row(
          [cell([paragraph("hidden"), table([row([cell([paragraph("hidden nested")])])])])],
          true,
        ),
      ]),
    ]);

    expect(createFolioAIEditSnapshot(doc).blocks.map(({ text }) => text)).toEqual(["visible"]);
  });

  test("resolves no positions, so its cost stays linear in block count", () => {
    const rows = Array.from({ length: 40 }, (_unused, index) =>
      row([
        cell([paragraph(`row ${String(index)} left`)]),
        cell([paragraph(`row ${String(index)} right`)]),
      ]),
    );
    const paragraphs = Array.from({ length: 200 }, (_unused, index) =>
      paragraph(`body paragraph ${String(index)}`),
    );
    const doc = refusingToResolve(
      schema.node("doc", null, [...paragraphs, table(rows), ...paragraphs]),
    );

    expect(createFolioAIEditSnapshot(doc).blocks).toHaveLength(480);
  });
});
