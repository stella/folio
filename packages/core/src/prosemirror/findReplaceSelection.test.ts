import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";

import { createDefaultFindOptions } from "../utils/findReplace";
import { findInProseMirrorDocument, resolveFindMatchRange } from "./findReplaceSelection";

const schema = new Schema({
  nodes: {
    doc: { content: "(paragraph | table | textBox | blockSdt)+" },
    paragraph: { group: "block", content: "inline*" },
    table: { group: "block", content: "tableRow+" },
    tableRow: { content: "(tableCell | tableHeader)+" },
    tableCell: { content: "(paragraph | table)+" },
    tableHeader: { content: "(paragraph | table)+" },
    textBox: { group: "block", content: "(paragraph | table)+" },
    blockSdt: { group: "block", content: "(paragraph | table | blockSdt)+" },
    tab: { group: "inline", inline: true },
    hardBreak: { group: "inline", inline: true },
    text: { group: "inline" },
  },
});

describe("Folio find match selection", () => {
  test("maps paragraph-relative find offsets to ProseMirror positions", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Preferred stock")]),
      schema.node("paragraph", null, [schema.text("Common stock")]),
    ]);

    expect(
      resolveFindMatchRange(doc, {
        paragraphIndex: 1,
        startOffset: 7,
        endOffset: 12,
      }),
    ).toEqual({ from: 25, to: 30 });
  });

  test("skips text box paragraphs to match document search traversal", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Preferred stock")]),
      schema.node("textBox", null, [
        schema.node("paragraph", null, [schema.text("stock in text box")]),
      ]),
      schema.node("paragraph", null, [schema.text("Common stock")]),
    ]);

    expect(
      resolveFindMatchRange(doc, {
        paragraphIndex: 1,
        startOffset: 7,
        endOffset: 12,
      }),
    ).toEqual({ from: 46, to: 51 });
  });

  test("maps offsets after inline search tokens", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("pre"),
        schema.node("tab"),
        schema.text("stock"),
        schema.node("hardBreak"),
        schema.text("tail"),
      ]),
    ]);

    expect(
      resolveFindMatchRange(doc, {
        paragraphIndex: 0,
        startOffset: 4,
        endOffset: 9,
      }),
    ).toEqual({ from: 5, to: 10 });
  });

  test("finds and resolves live ranges across body and table paragraphs", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Outside stock")]),
      schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", null, [
            schema.node("paragraph", null, [schema.text("Inside stock")]),
          ]),
        ]),
      ]),
    ]);

    const matches = findInProseMirrorDocument(doc, "stock", createDefaultFindOptions());

    expect(matches).toHaveLength(2);
    expect(matches.map(({ paragraphIndex, text }) => ({ paragraphIndex, text }))).toEqual([
      { paragraphIndex: 0, text: "stock" },
      { paragraphIndex: 1, text: "stock" },
    ]);
    expect(matches.every(({ from, to }) => doc.textBetween(from, to) === "stock")).toBe(true);
  });

  test("finds paragraphs nested in block content controls", () => {
    const doc = schema.node("doc", null, [
      schema.node("blockSdt", null, [
        schema.node("paragraph", null, [schema.text("Controlled stock")]),
      ]),
    ]);

    const match = findInProseMirrorDocument(doc, "stock", createDefaultFindOptions()).at(0);

    expect(match).toMatchObject({ paragraphIndex: 0, text: "stock" });
    expect(match && resolveFindMatchRange(doc, match)).toEqual(
      match ? { from: match.from, to: match.to } : null,
    );
  });

  test("honors case and whole-word options", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Stock stock stockholder")]),
    ]);

    expect(
      findInProseMirrorDocument(doc, "Stock", {
        matchCase: true,
        matchWholeWord: true,
      }),
    ).toHaveLength(1);
    expect(
      findInProseMirrorDocument(doc, "stock", {
        matchCase: false,
        matchWholeWord: true,
      }),
    ).toHaveLength(2);
  });
});
