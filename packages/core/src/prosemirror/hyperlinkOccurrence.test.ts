import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { expectHyperlinkMarkAttrs } from "./attrs";
import { canonicalizeHyperlinkOccurrenceIndexes } from "./hyperlinkOccurrence";
import { schema } from "./schema";

const hyperlinkIndexes = (state: EditorState): number[] => {
  const indexes: number[] = [];
  state.doc.descendants((node) => {
    const hyperlink = node.marks.find(({ type }) => type.name === "hyperlink");
    if (!hyperlink) return;
    const index = expectHyperlinkMarkAttrs(hyperlink)._docxHyperlinkIndex;
    if (index !== undefined && indexes.at(-1) !== index) indexes.push(index);
  });
  return indexes;
};

describe("hyperlink occurrence canonicalization", () => {
  test("closes view-local gaps without merging adjacent same-target occurrences", () => {
    const href = "https://example.invalid/same";
    const first = schema.marks.hyperlink.create({ href, _docxHyperlinkIndex: 7 });
    const second = schema.marks.hyperlink.create({ href, _docxHyperlinkIndex: 12 });
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("one", [first]), schema.text("two", [second])]),
      ]),
    });

    const canonical = state.apply(canonicalizeHyperlinkOccurrenceIndexes(state.tr));
    expect(hyperlinkIndexes(canonical)).toEqual([0, 1]);
  });

  test("gives a surviving later occurrence the first view-local identity", () => {
    const hyperlink = schema.marks.hyperlink.create({
      href: "https://example.invalid/later",
      _docxHyperlinkIndex: 12,
    });
    const state = EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("later", [hyperlink])]),
      ]),
    });

    const canonical = state.apply(canonicalizeHyperlinkOccurrenceIndexes(state.tr));
    expect(hyperlinkIndexes(canonical)).toEqual([0]);
  });
});
