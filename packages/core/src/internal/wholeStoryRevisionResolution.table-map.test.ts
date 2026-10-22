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
const deletedParagraphMark = {
  kind: "del",
  info: { id: 7, author: "Reviewer", date: "2026-10-19" },
} as const;

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
  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null, numbering: null });
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
        cell("deleted", {
          cellMarker: { kind: "del", info: { revisionId: 2, author: "Reviewer" } },
        }),
        cell("right"),
      ]),
    ]),
  ]);
  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null, numbering: null });
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
  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null, numbering: null });
  const source = paragraphPosition(doc, "after");
  const final = paragraphPosition(result.resolved, "after");

  expect(result.mapping.map(source, 1)).toBe(final);
  expect(final).toBeLessThan(source);
});

test("accepting an empty paragraph-mark deletion removes the paragraph before a table", () => {
  const following = table([row([cell("cell")])]);
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { pPrMark: deletedParagraphMark }),
    following,
  ]);

  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null, numbering: null });

  expect(result.failed).toBe(false);
  expect(result.resolved.content.content).toEqual([following]);
});

test("accepting a non-empty paragraph-mark deletion runs its content into the table", () => {
  const cellParagraph = paragraph("cell");
  const following = table([row([schema.node("tableCell", null, [cellParagraph])])]);
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { pPrMark: deletedParagraphMark }, [schema.text("kept")]),
    following,
  ]);

  const result = resolveWholeStory({ doc, mode: "accept", styleResolver: null, numbering: null });

  expect(result.failed).toBe(false);
  expect(result.resolved.childCount).toBe(1);
  const survivor = result.resolved.firstChild?.firstChild?.firstChild?.firstChild;
  expect(survivor?.textContent).toBe("keptcell");
  expect(survivor?.attrs["paraId"]).toBe(cellParagraph.attrs["paraId"]);
  expect(survivor?.attrs["pPrMark"]).toBeNull();
});

const insertedParagraphMark = {
  kind: "ins",
  info: { id: 8, author: "Reviewer", date: "2026-10-21" },
} as const;

test("rejecting an empty inserted final paragraph after a table keeps it as the terminator", () => {
  const terminator = schema.node("paragraph", { pPrMark: insertedParagraphMark });
  const doc = schema.node("doc", null, [
    paragraph("before"),
    table([row([cell("cell")])]),
    terminator,
  ]);

  const result = resolveWholeStory({ doc, mode: "reject", styleResolver: null, numbering: null });

  expect(result.failed).toBe(false);
  expect(result.resolved.childCount).toBe(3);
  expect(result.resolved.lastChild?.type.name).toBe("paragraph");
});

test("rejecting an inserted boundary after a table keeps its section on the joined paragraph", () => {
  const sectionProperties = { sectionStart: "nextPage", pageWidth: 10_000 };
  const doc = schema.node("doc", null, [
    table([row([cell("cell")])]),
    schema.node(
      "paragraph",
      { pPrMark: insertedParagraphMark, _sectionProperties: sectionProperties },
      [schema.text("before")],
    ),
    paragraph("after"),
  ]);

  const result = resolveWholeStory({ doc, mode: "reject", styleResolver: null, numbering: null });

  expect(result.failed).toBe(false);
  expect(result.resolved.lastChild?.textContent).toBe("beforeafter");
  expect(result.resolved.lastChild?.attrs["_sectionProperties"]).toEqual(sectionProperties);
});

test("rejecting an empty inserted final paragraph in a cell leaves the paragraph before it last", () => {
  const terminator = schema.node("paragraph", { pPrMark: insertedParagraphMark });
  const doc = schema.node("doc", null, [
    table([row([schema.node("tableCell", null, [paragraph("text"), terminator])])]),
    paragraph("after"),
  ]);

  const result = resolveWholeStory({ doc, mode: "reject", styleResolver: null, numbering: null });

  expect(result.failed).toBe(false);
  const resolvedCell = result.resolved.firstChild?.firstChild?.firstChild;
  expect(resolvedCell?.childCount).toBe(1);
  expect(resolvedCell?.lastChild?.textContent).toBe("text");
});

test("rejecting the only paragraph of a cell keeps it as the cell terminator", () => {
  const terminator = schema.node("paragraph", { pPrMark: insertedParagraphMark });
  const doc = schema.node("doc", null, [
    table([row([schema.node("tableCell", null, [terminator])])]),
    paragraph("after"),
  ]);

  const result = resolveWholeStory({ doc, mode: "reject", styleResolver: null, numbering: null });

  expect(result.failed).toBe(false);
  const resolvedCell = result.resolved.firstChild?.firstChild?.firstChild;
  expect(resolvedCell?.childCount).toBe(1);
  expect(resolvedCell?.firstChild?.type.name).toBe("paragraph");
});
