import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { schema } from "../prosemirror/schema";
import { resolveWholeStory } from "./wholeStoryRevisionResolution";

const paragraph = (text: string) => schema.node("paragraph", null, text ? [schema.text(text)] : []);
const cell = (text: string, attrs?: Record<string, unknown>) =>
  schema.node("tableCell", attrs, [paragraph(text)]);
const row = (cells: PMNode[], attrs?: Record<string, unknown>) =>
  schema.node("tableRow", attrs, cells);
const table = (rows: PMNode[]) => schema.node("table", null, rows);

const paragraphPosition = (doc: PMNode, text: string): number => {
  let found: number | null = null;
  doc.descendants((node, position) => {
    if (node.type.name === "paragraph" && node.textContent === text) {
      found = position;
      return false;
    }
    return true;
  });
  if (found === null) throw new Error(`Missing paragraph: ${text}`);
  return found;
};

test("row marker clearance leaves table positions unchanged", () => {
  const doc = schema.node("doc", null, [
    table([
      row([cell("first")]),
      row([cell("second")], { trIns: { revisionId: 1, author: "Reviewer" } }),
    ]),
  ]);
  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null });
  const source = paragraphPosition(doc, "second");
  const final = paragraphPosition(result.resolved, "second");

  expect(result.mapping.map(source, 1)).toBe(final);
  expect(result.structuralMap.map(source, 1)).toBe(source);
});

test("cell deletion preserves positions in surviving cells", () => {
  const doc = schema.node("doc", null, [
    table([
      row([
        cell("left"),
        cell("deleted", { cellMarker: { kind: "del", info: { revisionId: 2 } } }),
        cell("right"),
      ]),
    ]),
  ]);
  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null });
  const source = paragraphPosition(doc, "right");
  const final = paragraphPosition(result.resolved, "right");

  expect(result.mapping.map(source, 1)).toBe(final);
  expect(final).toBeLessThan(source);
});

test("nested table resolution preserves a following paragraph position", () => {
  const inner = table([
    row([cell("removed")], { trDel: { revisionId: 3, author: "Reviewer" } }),
    row([cell("kept")]),
  ]);
  const outerCell = schema.node("tableCell", null, [
    paragraph("before"),
    inner,
    paragraph("after"),
  ]);
  const doc = schema.node("doc", null, [table([row([outerCell])])]);
  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null });
  const source = paragraphPosition(doc, "after");
  const final = paragraphPosition(result.resolved, "after");

  expect(result.mapping.map(source, 1)).toBe(final);
  expect(final).toBeLessThan(source);
});
