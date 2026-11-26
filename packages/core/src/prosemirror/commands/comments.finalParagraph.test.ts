import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { rejectChange } from "../../__tests__/reference-editor/comments";
import { schema } from "../schema";

const insertedParagraphMark = {
  kind: "ins",
  info: { id: 8, author: "Reviewer", date: "2026-10-21" },
} as const;

const paragraph = (text: string) =>
  schema.node("paragraph", null, text ? [schema.text(text)] : []);
const table = (text: string) =>
  schema.node("table", null, [
    schema.node("tableRow", null, [
      schema.node("tableCell", null, [paragraph(text)]),
    ]),
  ]);

const rejectFinalParagraph = (doc: PMNode): PMNode => {
  let state = EditorState.create({ schema, doc });
  const last = doc.lastChild;
  if (last === null) throw new Error("Expected a final block");
  const from = doc.content.size - last.nodeSize;
  const dispatch = (tr: Transaction): void => {
    state = state.apply(tr);
  };
  expect(rejectChange(from, doc.content.size)(state, dispatch)).toBe(true);
  return state.doc;
};

test("rejecting an empty inserted final paragraph after a table keeps it as the terminator", () => {
  const doc = schema.node("doc", null, [
    paragraph("before"),
    table("cell"),
    schema.node("paragraph", { pPrMark: insertedParagraphMark }),
  ]);

  const resolved = rejectFinalParagraph(doc);

  expect(resolved.childCount).toBe(3);
  expect(resolved.lastChild?.type.name).toBe("paragraph");
});

test("rejecting an empty inserted final paragraph after a paragraph removes it", () => {
  const doc = schema.node("doc", null, [
    paragraph("before"),
    schema.node("paragraph", { pPrMark: insertedParagraphMark }),
  ]);

  const resolved = rejectFinalParagraph(doc);

  expect(resolved.childCount).toBe(1);
  expect(resolved.lastChild?.textContent).toBe("before");
});

test("rejecting an added mark keeps the sole paragraph in a block control", () => {
  const original = schema.node("paragraph", { paraId: "ABCDEF01" });
  const control = schema.node("blockSdt", null, [
    schema.node("paragraph", {
      ...original.attrs,
      pPrMark: insertedParagraphMark,
    }),
  ]);
  const doc = schema.node("doc", null, [paragraph("before"), control]);

  const resolved = rejectFinalParagraph(doc);

  expect(resolved.lastChild?.type.name).toBe("blockSdt");
  expect(resolved.lastChild?.childCount).toBe(1);
  expect(resolved.lastChild?.firstChild?.eq(original)).toBe(true);
});
