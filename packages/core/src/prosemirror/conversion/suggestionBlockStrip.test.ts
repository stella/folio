/**
 * Class-guard tests for the BLOCK/TABLE suggestion serialization strip.
 *
 * Suggested whole-node inserts (paragraph, table, row, column cell) must be
 * dropped from serialized DOCX entirely; suggested deletes must serialize as
 * though they never happened. Real (user) tracked structural changes adjacent
 * to suggested ones must survive unchanged.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import type { BlockContent, Document, Table } from "../../types/document";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";

const DATE = "2026-07-17T00:00:00.000Z";

const para = (text: string, attrs: Record<string, unknown> = {}): PMNode =>
  // SAFETY: editor schema always defines paragraph.
  schema.nodes["paragraph"]!.create(attrs, text.length > 0 ? [schema.text(text)] : []);

const cell = (text: string, attrs: Record<string, unknown> = {}): PMNode =>
  // SAFETY: editor schema always defines tableCell.
  schema.nodes["tableCell"]!.create(attrs, [para(text)]);

const row = (cells: PMNode[], attrs: Record<string, unknown> = {}): PMNode =>
  // SAFETY: editor schema always defines tableRow.
  schema.nodes["tableRow"]!.create(attrs, cells);

const table = (rows: PMNode[], attrs: Record<string, unknown> = {}): PMNode =>
  // SAFETY: editor schema always defines table.
  schema.nodes["table"]!.create(attrs, rows);

const doc = (blocks: PMNode[]): PMNode =>
  // SAFETY: editor schema always defines doc.
  schema.nodes["doc"]!.create({}, blocks);

const suggested = (revisionId: number, suggestionId: string) => ({
  revisionId,
  author: "AI",
  date: DATE,
  provenance: "suggested" as const,
  suggestionId,
});

const structuralInsert = (revisionId: number, suggestionId: string) => ({
  revisionId,
  author: "AI",
  date: DATE,
  suggestionId,
});

const blocks = (model: Document): BlockContent[] => model.package.document.content;

const firstTable = (model: Document): Table => {
  const block = blocks(model).find((b) => b.type === "table");
  if (block?.type !== "table") {
    throw new Error("expected a table block");
  }
  return block;
};

const plainText = (model: Document): string => JSON.stringify(model.package.document.content);

describe("block/table suggestion strip", () => {
  test("a suggested-inserted paragraph is dropped entirely", () => {
    const model = fromProseDoc(
      doc([
        para("keep me"),
        para("proposed paragraph", { _suggestedInsert: structuralInsert(900, "s1") }),
      ]),
    );
    expect(blocks(model).length).toBe(1);
    expect(plainText(model)).not.toContain("proposed paragraph");
    expect(plainText(model)).toContain("keep me");
  });

  test("a suggested-inserted table is dropped entirely", () => {
    const model = fromProseDoc(
      doc([
        para("before"),
        table([row([cell("proposed cell")])], { _suggestedInsert: structuralInsert(901, "s2") }),
      ]),
    );
    expect(blocks(model).some((b) => b.type === "table")).toBe(false);
    expect(plainText(model)).not.toContain("proposed cell");
  });

  test("a suggested-inserted row is dropped; real inserted rows survive", () => {
    const model = fromProseDoc(
      doc([
        table([
          row([cell("original")]),
          row([cell("real added")], { trIns: structuralInsert(10, "ignored") }),
          row([cell("suggested added")], { trIns: suggested(11, "s3") }),
        ]),
      ]),
    );
    const t = firstTable(model);
    expect(t.rows.length).toBe(2);
    expect(plainText(model)).not.toContain("suggested added");
    expect(plainText(model)).toContain("real added");
    // The real inserted row keeps its structural change; the suggested one is gone.
    const inserted = t.rows.filter((r) => r.structuralChange?.type === "tableRowInsertion");
    expect(inserted.length).toBe(1);
  });

  test("a suggested-deleted row serializes as a plain row (deletion never happened)", () => {
    const model = fromProseDoc(
      doc([table([row([cell("stays")], { trDel: suggested(12, "s4") }), row([cell("other")])])]),
    );
    const t = firstTable(model);
    expect(t.rows.length).toBe(2);
    expect(plainText(model)).toContain("stays");
    expect(t.rows.every((r) => r.structuralChange === undefined)).toBe(true);
  });

  test("a suggested-inserted column cell is dropped; real cell survives", () => {
    const model = fromProseDoc(
      doc([
        table([
          row([
            cell("keep"),
            cell("suggested col", {
              cellMarker: { kind: "ins", info: suggested(13, "s5") },
            }),
          ]),
        ]),
      ]),
    );
    const t = firstTable(model);
    expect(t.rows[0]?.cells.length).toBe(1);
    expect(plainText(model)).not.toContain("suggested col");
    expect(plainText(model)).toContain("keep");
  });

  test("a suggested-deleted column cell keeps the cell as plain content", () => {
    const model = fromProseDoc(
      doc([
        table([
          row([cell("a"), cell("b", { cellMarker: { kind: "del", info: suggested(14, "s6") } })]),
        ]),
      ]),
    );
    const t = firstTable(model);
    expect(t.rows[0]?.cells.length).toBe(2);
    expect(t.rows[0]?.cells.every((c) => c.structuralChange === undefined)).toBe(true);
    expect(plainText(model)).toContain("b");
  });

  test("nested: suggested inline edit inside a suggested-inserted paragraph → whole block gone", () => {
    const insMark = schema.marks["insertion"]!.create({
      revisionId: 20,
      author: "AI",
      date: DATE,
      provenance: "suggested",
      suggestionId: "s7",
    });
    const model = fromProseDoc(
      doc([
        para("anchor"),
        // SAFETY: schema always defines paragraph/text/insertion.
        schema.nodes["paragraph"]!.create({ _suggestedInsert: structuralInsert(20, "s7") }, [
          schema.text("nested proposed", [insMark]),
        ]),
      ]),
    );
    expect(blocks(model).length).toBe(1);
    expect(plainText(model)).not.toContain("nested proposed");
  });
});
