/**
 * Typing over a selection that reaches across a comment's end, or over a
 * point comment, directly: the comment keeps the words before its boundary
 * and the new words follow it, as accepting the same typing as a suggestion
 * leaves them. Undo restores the document.
 */
import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";

import {
  createCanonicalEditorHarness,
  resolveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import { assertExactModel } from "../../../../test/exactModel";
import { documentShape, shapeArrayBuffer } from "../__tests__/documentShapes";
import { parseShapeDocument } from "../__tests__/editorHarness";
import type { Document, ParagraphContent } from "../types/document";

/** The positions of `needle` in the first textblock holding it. */
const rangeOf = (doc: PMNode, needle: string) => {
  let found: { from: number; to: number } | undefined;
  doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (!node.isTextblock) return true;
    const at = node.textContent.indexOf(needle);
    if (at < 0) return false;
    let offset = 0;
    node.forEach((child, childOffset) => {
      if (!child.isText) return;
      const length = child.text?.length ?? 0;
      const start = position + 1 + childOffset;
      if (found === undefined && at >= offset && at < offset + length)
        found = { from: start + at - offset, to: -1 };
      const end = at + needle.length;
      if (found !== undefined && end > offset && end <= offset + length)
        found.to = start + end - offset;
      offset += length;
    });
    return false;
  });
  if (found === undefined || found.to < 0)
    throw new TypeError(`The fixture holds no "${needle}"`);
  return found;
};

/** Text and comment markers in order, the shape a reader sees. */
const markers = (document: Document) => {
  const out: string[] = [];
  const visit = (items: readonly ParagraphContent[]) => {
    for (const item of items) {
      if (item.type === "run")
        for (const piece of item.content)
          if (piece.type === "text") out.push(piece.text);
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd" ||
        item.type === "commentReference"
      )
        out.push(`<${item.type}:${String(item.id)}>`);
      if (item.type === "insertion") visit(item.content);
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") visit(block.content);
  return out.join("");
};

const type = async (
  shape: "comments" | "invisible-markers",
  needle: string,
  mode: "editing" | "suggesting",
) => {
  const source = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer(documentShape(shape))),
  );
  const driver = createCanonicalEditorHarness(source, mode);
  try {
    const { from, to } = rangeOf(driver.state.doc, needle);
    driver.dispatch(
      driver.state.tr.setSelection(
        TextSelection.create(driver.state.doc, from, to),
      ),
    );
    const before = driver.snapshot();
    driver.typeText("X");
    expect(driver.refusals).toEqual([]);
    const typed = driver.snapshot();
    // The caret follows the new word.
    const { from: caret } = driver.state.selection;
    expect(driver.state.doc.textBetween(caret - 1, caret)).toBe("X");
    expect(driver.history.undo()).toBe(true);
    assertExactModel(driver.snapshot(), before);
    return typed;
  } finally {
    driver.dispose();
  }
};

test.each([
  [
    "comments",
    "words sit",
    "<commentRangeStart:0>Commented <commentRangeEnd:0><commentReference:0>X in this paragraph.",
  ],
  [
    "comments",
    "ted words sit in",
    "<commentRangeStart:0>Commen<commentRangeEnd:0><commentReference:0>X this paragraph.",
  ],
  [
    "invisible-markers",
    "sits between",
    "A point comment <commentRangeStart:3><commentRangeEnd:3><commentReference:3>X these words.",
  ],
  [
    "invisible-markers",
    "words stand",
    "<commentRangeStart:0>Marked <commentRangeEnd:0><commentReference:0>X between invisible markers here.",
  ],
] as const)(
  "%s: typing over %p puts the new word past the comment, as the accepted suggestion does",
  async (shape, needle, expected) => {
    const direct = await type(shape, needle, "editing");
    expect(markers(direct)).toContain(expected);
    const accepted = resolveCanonicalHarnessDocument(
      await type(shape, needle, "suggesting"),
      "accept",
    ).model;
    expect(markers(direct)).toBe(markers(accepted));
  },
);
