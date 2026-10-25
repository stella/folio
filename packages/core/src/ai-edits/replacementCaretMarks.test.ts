/**
 * Text inserted at the caret takes the caret's stored formatting, whichever
 * path builds it (plain text, text with a clean-text control such as a tab,
 * text carrying the link of the span it stands in for); text inserted
 * elsewhere takes the formatting around it. Stored marks only exist at an
 * empty selection, so the caret is where they apply.
 */

import { describe, expect, test } from "bun:test";
import type { Mark } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { schema } from "../prosemirror/schema";
import { insertCleanText } from "./apply";

const TEXT = "the quick brown fox";
const FIND = "quick";
const FROM = 1 + TEXT.indexOf(FIND);
const TO = FROM + FIND.length;

const markOf = (name: string, attrs?: Record<string, unknown>): Mark => {
  const type = schema.marks[name];
  if (!type) throw new Error(`expected a ${name} mark`);
  return type.create(attrs);
};

/** `TEXT` with `findMarks` on `FIND`, the caret after `caretAfter`, bold stored there. */
const stateWith = (findMarks: readonly Mark[], caretAfter: string) => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", {}, [
      schema.text(TEXT.slice(0, FROM - 1)),
      schema.text(FIND, findMarks),
      schema.text(TEXT.slice(TO - 1)),
    ]),
  ]);
  const caret = 1 + TEXT.indexOf(caretAfter) + caretAfter.length;
  const state = EditorState.create({ schema, doc });
  const stored = state.apply(
    state.tr
      .setSelection(TextSelection.create(doc, caret))
      .setStoredMarks([markOf("bold")]),
  );
  expect(stored.storedMarks?.map(({ type }) => type.name)).toEqual(["bold"]);
  return stored;
};

/** The mark names on the inserted text, which must be uniform. */
const insertedMarks = (
  findMarks: readonly Mark[],
  caretAfter: string,
  text: string,
): string[] => {
  // Inserted after `FIND`, standing in for it, so it carries its link.
  const { transaction, start, end } = insertCleanText({
    tr: stateWith(findMarks, caretAfter).tr,
    from: TO,
    to: TO,
    text,
    standsInFor: { from: FROM, to: TO },
  });
  const names = new Set<string>();
  transaction.doc.nodesBetween(start, end, (node) => {
    if (node.isInline)
      names.add(
        node.marks
          .map(({ type }) => type.name)
          .toSorted()
          .join(","),
      );
  });
  expect(names.size).toBe(1);
  return [...names][0]?.split(",").filter(Boolean) ?? [];
};

const LINK = markOf("hyperlink", { href: "https://example.com" });

describe("a replacement's caret formatting", () => {
  for (const { name, findMarks, text } of [
    { name: "plain text", findMarks: [], text: "smart" },
    { name: "text with a tab", findMarks: [], text: "sm\tart" },
    { name: "text over a link", findMarks: [LINK], text: "smart" },
  ]) {
    test(`${name} at the caret takes the caret's formatting`, () => {
      const marks = insertedMarks(findMarks, FIND, text);
      expect(marks).toContain("bold");
      for (const mark of findMarks) expect(marks).toContain(mark.type.name);
    });

    test(`${name} away from the caret takes the text's own`, () => {
      const marks = insertedMarks(findMarks, "fox", text);
      expect(marks).not.toContain("bold");
      for (const mark of findMarks) expect(marks).toContain(mark.type.name);
    });
  }
});
