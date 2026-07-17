/**
 * Accept/reject + getSuggestions for BLOCK/TABLE suggestions.
 *
 * Accepting converts a suggestion into a real tracked change (or, for a whole
 * inserted table, applies it directly); rejecting inverse-applies it. Reuses
 * the manual-doc builders so table geometry stays explicit.
 */

import { describe, expect, test } from "bun:test";
import type { Command } from "prosemirror-state";
import { EditorState } from "prosemirror-state";
import type { Node as PMNode } from "prosemirror-model";

import { fromProseDoc } from "../conversion/fromProseDoc";
import { schema } from "../schema";
import type { Document, Table } from "../../types/document";
import { acceptSuggestion, getSuggestions, rejectSuggestion } from "./comments";

const DATE = "2026-07-17T00:00:00.000Z";

const para = (text: string, attrs: Record<string, unknown> = {}): PMNode =>
  schema.nodes["paragraph"]!.create(attrs, text.length > 0 ? [schema.text(text)] : []);

const cell = (text: string, attrs: Record<string, unknown> = {}): PMNode =>
  schema.nodes["tableCell"]!.create(attrs, [para(text)]);

const row = (cells: PMNode[], attrs: Record<string, unknown> = {}): PMNode =>
  schema.nodes["tableRow"]!.create(attrs, cells);

const table = (rows: PMNode[], attrs: Record<string, unknown> = {}): PMNode =>
  schema.nodes["table"]!.create(attrs, rows);

const makeState = (blocks: PMNode[]) =>
  EditorState.create({ schema, doc: schema.nodes["doc"]!.create({}, blocks) });

const run = (state: EditorState, command: Command): { state: EditorState; ok: boolean } => {
  let next = state;
  const ok = command(state, (tr) => {
    next = state.apply(tr);
  });
  return { state: next, ok };
};

const suggestedMarker = (revisionId: number, suggestionId: string) => ({
  revisionId,
  author: "AI",
  date: DATE,
  provenance: "suggested" as const,
  suggestionId,
});

const firstTable = (model: Document): Table => {
  const block = model.package.document.content.find((b) => b.type === "table");
  if (block?.type !== "table") {
    throw new Error("expected a table block");
  }
  return block;
};

const json = (model: Document): string => JSON.stringify(model.package.document.content);

describe("block/table suggestion accept/reject + getSuggestions", () => {
  test("getSuggestions reports block/table kinds and appliedAs", () => {
    const state = makeState([
      para("proposed para", { _suggestedInsert: { ...suggestedMarker(20, "sBlock") } }),
      table([row([cell("x")], { trIns: suggestedMarker(11, "sRow") })]),
      table([row([cell("y")])], { _suggestedInsert: { ...suggestedMarker(30, "sTable") } }),
    ]);
    const byId = new Map(getSuggestions(state).map((s) => [s.suggestionId, s]));

    expect(byId.get("sBlock")?.kinds).toContain("insertBlock");
    expect(byId.get("sBlock")?.appliedAs).toBe("tracked");
    expect(byId.get("sRow")?.kinds).toContain("insertRow");
    expect(byId.get("sRow")?.appliedAs).toBe("tracked");
    expect(byId.get("sTable")?.kinds).toContain("insertTable");
    expect(byId.get("sTable")?.appliedAs).toBe("direct");
  });

  test("accept a suggested row insertion → real tracked row (author reassigned)", () => {
    const state = makeState([
      table([row([cell("orig")]), row([cell("added")], { trIns: suggestedMarker(11, "sRow") })]),
    ]);
    const { ok, state: next } = run(
      state,
      acceptSuggestion("sRow", { author: "Alice", date: DATE }),
    );
    expect(ok).toBe(true);
    expect(getSuggestions(next)).toEqual([]);

    // The row now serializes as a real tracked insertion authored by Alice.
    const t = firstTable(fromProseDoc(next.doc));
    const inserted = t.rows.find((r) => r.structuralChange?.type === "tableRowInsertion");
    expect(inserted?.structuralChange?.info.author).toBe("Alice");
    expect(json(fromProseDoc(next.doc))).toContain("added");
  });

  test("reject a suggested row insertion → row removed", () => {
    const state = makeState([
      table([row([cell("orig")]), row([cell("added")], { trIns: suggestedMarker(11, "sRow") })]),
    ]);
    const { ok, state: next } = run(state, rejectSuggestion("sRow"));
    expect(ok).toBe(true);
    const t = firstTable(fromProseDoc(next.doc));
    expect(t.rows.length).toBe(1);
    expect(json(fromProseDoc(next.doc))).not.toContain("added");
  });

  test("accept a suggested paragraph insertion → inserted-paragraph tracked change", () => {
    const insMark = schema.marks["insertion"]!.create({
      revisionId: 20,
      author: "AI",
      date: DATE,
      provenance: "suggested",
      suggestionId: "sBlock",
    });
    const state = makeState([
      para("anchor"),
      schema.nodes["paragraph"]!.create(
        { _suggestedInsert: { ...suggestedMarker(20, "sBlock") } },
        [schema.text("proposed", [insMark])],
      ),
    ]);
    const { ok, state: next } = run(
      state,
      acceptSuggestion("sBlock", { author: "Bob", date: DATE }),
    );
    expect(ok).toBe(true);
    expect(getSuggestions(next)).toEqual([]);

    // The accepted paragraph serializes with a real inserted paragraph mark and
    // a real (user) run insertion — no suggestion leaks.
    const model = fromProseDoc(next.doc);
    const serialized = json(model);
    expect(serialized).toContain("proposed");
    expect(serialized).toContain("pPrMark");
    expect(serialized).toContain('"author":"Bob"');
    expect(serialized).not.toContain("suggested");
  });

  test("reject a suggested paragraph insertion → paragraph removed", () => {
    const insMark = schema.marks["insertion"]!.create({
      revisionId: 20,
      author: "AI",
      date: DATE,
      provenance: "suggested",
      suggestionId: "sBlock",
    });
    const state = makeState([
      para("anchor"),
      schema.nodes["paragraph"]!.create(
        { _suggestedInsert: { ...suggestedMarker(20, "sBlock") } },
        [schema.text("proposed", [insMark])],
      ),
    ]);
    const { ok, state: next } = run(state, rejectSuggestion("sBlock"));
    expect(ok).toBe(true);
    expect(next.doc.childCount).toBe(1);
    expect(json(fromProseDoc(next.doc))).not.toContain("proposed");
  });

  test("accept a suggested table insertion applies directly (table kept, marker cleared)", () => {
    const state = makeState([
      para("before"),
      table([row([cell("sig")])], { _suggestedInsert: { ...suggestedMarker(30, "sTable") } }),
    ]);
    const { ok, state: next } = run(
      state,
      acceptSuggestion("sTable", { author: "Carol", date: DATE }),
    );
    expect(ok).toBe(true);
    expect(getSuggestions(next)).toEqual([]);
    const model = fromProseDoc(next.doc);
    expect(model.package.document.content.some((b) => b.type === "table")).toBe(true);
    expect(json(model)).toContain("sig");
  });
});

// A suggested inline insertion mark (as an insertAfterBlock / replaceInBlock
// suggestion produces on its runs).
const insMark = (revisionId: number, suggestionId: string) =>
  schema.marks["insertion"]!.create({
    revisionId,
    author: "AI",
    date: DATE,
    provenance: "suggested",
    suggestionId,
  });

describe("heterogeneous (mixed) suggestion groups", () => {
  // One suggestionId spanning a whole-node paragraph insert AND a separate
  // inline edit — the reject regression: the old code deleted the node insert
  // and returned, leaving the inline mark unresolved.
  const mixedNodeAndInline = () =>
    makeState([
      para("keep ", {}), // will hold an inline suggested insertion below
      schema.nodes["paragraph"]!.create({ _suggestedInsert: { ...suggestedMarker(20, "mix") } }, [
        schema.text("proposed", [insMark(20, "mix")]),
      ]),
    ]);

  const mixedWithInlineEdit = () =>
    makeState([
      schema.nodes["paragraph"]!.create({}, [
        schema.text("keep "),
        schema.text("added", [insMark(21, "mix")]),
      ]),
      schema.nodes["paragraph"]!.create({ _suggestedInsert: { ...suggestedMarker(20, "mix") } }, [
        schema.text("proposed", [insMark(20, "mix")]),
      ]),
    ]);

  test("reject resolves BOTH the node insert and the separate inline edit", () => {
    const state = mixedWithInlineEdit();
    const { ok, state: next } = run(state, rejectSuggestion("mix"));
    expect(ok).toBe(true);
    // Inline-inserted "added" removed AND the whole inserted paragraph removed.
    expect(next.doc.textContent).toBe("keep ");
    expect(getSuggestions(next)).toEqual([]);
    expect(json(fromProseDoc(next.doc))).not.toContain("proposed");
    expect(json(fromProseDoc(next.doc))).not.toContain("added");
  });

  test("accept resolves BOTH parts of the group in one transaction", () => {
    const state = mixedWithInlineEdit();
    const { ok, state: next } = run(state, acceptSuggestion("mix", { author: "Bob", date: DATE }));
    expect(ok).toBe(true);
    expect(getSuggestions(next)).toEqual([]);
    const serialized = json(fromProseDoc(next.doc));
    // Both the inline edit and the inserted paragraph became real tracked changes.
    expect(serialized).toContain("added");
    expect(serialized).toContain("proposed");
    expect(serialized).toContain("pPrMark");
    expect(serialized).not.toContain("suggested");
  });

  test("node-only heterogeneous group still fully rejects", () => {
    const state = mixedNodeAndInline();
    const { ok, state: next } = run(state, rejectSuggestion("mix"));
    expect(ok).toBe(true);
    expect(next.doc.textContent).toBe("keep ");
    expect(getSuggestions(next)).toEqual([]);
  });

  test("a group mixing an inserted table with other edits reports appliedAs 'mixed'", () => {
    const state = makeState([
      schema.nodes["paragraph"]!.create({}, [schema.text("edit", [insMark(21, "mix2")])]),
      table([row([cell("sig")])], { _suggestedInsert: { ...suggestedMarker(30, "mix2") } }),
    ]);
    const suggestion = getSuggestions(state).find((s) => s.suggestionId === "mix2");
    expect(suggestion?.kinds).toContain("insertTable");
    expect(suggestion?.kinds).toContain("insertion");
    expect(suggestion?.appliedAs).toBe("mixed");

    // Accept resolves the whole group: table applied directly, inline → tracked.
    const { ok, state: next } = run(
      state,
      acceptSuggestion("mix2", { author: "Carol", date: DATE }),
    );
    expect(ok).toBe(true);
    expect(getSuggestions(next)).toEqual([]);
    const model = fromProseDoc(next.doc);
    expect(model.package.document.content.some((b) => b.type === "table")).toBe(true);
    expect(json(model)).toContain("edit");
  });
});

describe("acceptSuggestion date validation", () => {
  test("a malformed date is normalized so no invalid w:date is written", () => {
    const state = makeState([
      schema.nodes["paragraph"]!.create({}, [schema.text("x", [insMark(5, "s")])]),
    ]);
    const { ok, state: next } = run(
      state,
      acceptSuggestion("s", { author: "A", date: "not-a-date" }),
    );
    expect(ok).toBe(true);
    // The stamped date parses as a valid ISO instant, not the raw junk string.
    let stampedDate: unknown;
    next.doc.descendants((node) => {
      const mark = node.marks.find((m) => m.type.name === "insertion");
      if (mark) {
        stampedDate = mark.attrs["date"];
      }
      return undefined;
    });
    expect(typeof stampedDate).toBe("string");
    expect(Number.isNaN(new Date(stampedDate as string).getTime())).toBe(false);
    expect(stampedDate).not.toBe("not-a-date");
  });

  test("a valid date is preserved (canonicalized to ISO)", () => {
    const state = makeState([
      schema.nodes["paragraph"]!.create({}, [schema.text("x", [insMark(6, "s")])]),
    ]);
    const { state: next } = run(
      state,
      acceptSuggestion("s", { author: "A", date: "2026-07-17T00:00:00.000Z" }),
    );
    let stampedDate: unknown;
    next.doc.descendants((node) => {
      const mark = node.marks.find((m) => m.type.name === "insertion");
      if (mark) {
        stampedDate = mark.attrs["date"];
      }
      return undefined;
    });
    expect(stampedDate).toBe("2026-07-17T00:00:00.000Z");
  });
});
