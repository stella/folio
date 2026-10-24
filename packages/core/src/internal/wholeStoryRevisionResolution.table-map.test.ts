import { panic } from "better-result";
import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { Step, Transform } from "prosemirror-transform";

import {
  acceptAIEditRevision,
  rejectAIEditRevision,
} from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import { runParagraphIntoTable } from "../prosemirror/tableRunIn";
import { resolveWholeStory } from "./wholeStoryRevisionResolution";

const paragraph = (text: string) =>
  schema.node("paragraph", null, text ? [schema.text(text)] : []);
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

const anchorPosition = (doc: PMNode, commentId: number): number => {
  let found: number | null = null;
  doc.descendants((node, position) => {
    if (
      node.type.name === "rangeAnchor" &&
      node.attrs["start"]?.id === commentId
    ) {
      found = position;
      return false;
    }
    return true;
  });
  if (found === null) panic(`Missing comment anchor: ${commentId}`);
  return found;
};

type ResolveRevisionOptions = {
  state: EditorState;
  revisionId: number;
  mode: "accept" | "reject";
};

const resolveRevision = ({
  state,
  revisionId,
  mode,
}: ResolveRevisionOptions): EditorState => {
  let next: EditorState | undefined;
  const command =
    mode === "accept"
      ? acceptAIEditRevision(revisionId)
      : rejectAIEditRevision(revisionId);
  const accepted = command(state, (transaction) => {
    next = state.apply(transaction);
  });
  if (!accepted || next === undefined)
    panic(`Could not accept revision ${revisionId}`);
  return next;
};

test.each(["accept", "reject"] as const)(
  "%s resolves comment hosts across retiring predecessor chains",
  (mode) => {
    for (const text of ["", "kept"]) {
      for (const precedingHost of [false, true]) {
        const comment = schema.node("rangeAnchor", {
          start: { type: "commentRangeStart", id: 37 },
          end: { type: "commentRangeEnd", id: 37 },
        });
        const mark = (id: number) => ({
          pPrMark: {
            kind: mode === "accept" ? "del" : "ins",
            info: { id, author: "Reviewer" },
          },
        });
        const doc = schema.node("doc", null, [
          ...(precedingHost ? [paragraph("lead")] : []),
          schema.node("paragraph", mark(1), text ? [schema.text(text)] : []),
          schema.node("paragraph", mark(2), [comment]),
          table([row([cell("cell")])]),
          paragraph("tail"),
        ]);
        const wholeStory = resolveWholeStory({
          doc,
          mode,
          styleResolver: null,
          numbering: null,
        });
        expect(wholeStory.failed).toBe(false);
        for (const order of [
          [1, 2],
          [2, 1],
        ]) {
          let state = EditorState.create({ schema, doc });
          for (const revisionId of order) {
            state = resolveRevision({ state, revisionId, mode });
          }
          expect(state.doc.eq(wholeStory.resolved)).toBe(true);
        }
        const receivesOutside = precedingHost && text.length === 0;
        const host = receivesOutside
          ? wholeStory.resolved.firstChild
          : wholeStory.resolved.child(precedingHost ? 1 : 0).firstChild
              ?.firstChild?.firstChild;
        expect(host?.textContent).toBe(
          receivesOutside ? "lead" : `${text}cell`,
        );
        expect(
          host?.children.some(
            (child) =>
              child.type.name === "rangeAnchor" &&
              child.attrs["start"]?.id === 37,
          ),
        ).toBe(true);
      }
    }
  },
);

test("row marker clearance leaves table positions unchanged", () => {
  const doc = schema.node("doc", null, [
    table([
      row([cell("first")]),
      row([cell("second")], { trIns: { revisionId: 1, author: "Reviewer" } }),
    ]),
  ]);
  const result = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: null,
    numbering: null,
  });
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
          cellMarker: {
            kind: "del",
            info: { revisionId: 2, author: "Reviewer" },
          },
        }),
        cell("right"),
      ]),
    ]),
  ]);
  const result = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: null,
    numbering: null,
  });
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
  const result = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: null,
    numbering: null,
  });
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

  const result = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: null,
    numbering: null,
  });

  expect(result.failed).toBe(false);
  expect(result.resolved.content.content).toEqual([following]);
});

test("an unhosted point comment runs into the first cell and preserves transaction mapping and inversion", () => {
  const comment = schema.node("rangeAnchor", {
    start: { type: "commentRangeStart", id: 19 },
    end: { type: "commentRangeEnd", id: 19 },
  });
  const cellParagraph = schema.node("paragraph", { paraId: "10000002" }, [
    schema.text("cell"),
  ]);
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { pPrMark: deletedParagraphMark }, [comment]),
    table([row([schema.node("tableCell", null, [cellParagraph])])]),
    paragraph("tail"),
  ]);
  const transform = new Transform(doc);
  const runIn = runParagraphIntoTable({
    tr: transform,
    paragraphPos: 0,
    styleResolver: null,
  });
  if (runIn === null) panic("The unhosted point comment must run in.");
  const step = transform.steps.at(0);
  if (step === undefined) panic("The run-in must produce a step.");
  expect(
    Step.fromJSON(schema, step.toJSON()).apply(doc).doc?.eq(transform.doc),
  ).toBe(true);
  expect(step.invert(doc).apply(transform.doc).doc?.eq(doc)).toBe(true);
  const sourceCell = paragraphPosition(doc, "cell");
  const finalCell = paragraphPosition(transform.doc, "cell");
  expect(runIn.position).toBe(finalCell);
  expect(transform.mapping.map(sourceCell + 1, 1)).toBe(finalCell + 2);

  const result = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: null,
    numbering: null,
  });
  expect(result.failed).toBe(false);
  expect(result.resolved.childCount).toBe(2);
  const receiver =
    result.resolved.firstChild?.firstChild?.firstChild?.firstChild;
  expect(receiver?.attrs["paraId"]).toBe("10000002");
  expect(receiver?.textContent).toBe("cell");
  expect(receiver?.firstChild?.attrs).toEqual(comment.attrs);
  expect(result.mapping.map(anchorPosition(doc, 19), 1)).toBe(
    anchorPosition(result.resolved, 19),
  );
});

test("whole-story acceptance matches accepting adjacent paragraph marks in turn", () => {
  const comment = schema.node("rangeAnchor", {
    start: { type: "commentRangeStart", id: 29 },
    end: { type: "commentRangeEnd", id: 29 },
  });
  const cellParagraph = schema.node("paragraph", { paraId: "10000003" }, [
    schema.text("cell"),
  ]);
  const insertionMark = schema.marks["insertion"];
  if (insertionMark === undefined) panic("The schema has an insertion mark.");
  const doc = schema.node("doc", null, [
    schema.node(
      "paragraph",
      {
        pPrMark: {
          kind: "del",
          info: { id: 2, author: "Reviewer", date: "2026-10-19" },
        },
      },
      [comment],
    ),
    schema.node(
      "paragraph",
      {
        pPrMark: {
          kind: "del",
          info: { id: 1, author: "Reviewer", date: "2026-10-19" },
        },
      },
      [
        schema.text("x", [
          insertionMark.create({
            revisionId: 3,
            author: "Reviewer",
            date: "2026-10-19",
          }),
        ]),
      ],
    ),
    table([row([schema.node("tableCell", null, [cellParagraph])])]),
    paragraph("tail"),
  ]);
  const wholeStory = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: null,
    numbering: null,
  });
  let incremental = EditorState.create({ schema, doc });
  for (const revisionId of [2, 3, 1])
    incremental = resolveRevision({
      state: incremental,
      revisionId,
      mode: "accept",
    });
  expect(wholeStory.failed).toBe(false);
  expect(incremental.doc.eq(wholeStory.resolved)).toBe(true);
  const receiver =
    wholeStory.resolved.firstChild?.firstChild?.firstChild?.firstChild;
  expect(receiver?.attrs["paraId"]).toBe("10000003");
  expect(receiver?.textContent).toBe("xcell");
  expect(receiver?.firstChild?.attrs).toEqual(comment.attrs);
});

test("accepting a non-empty paragraph-mark deletion runs its content into the table", () => {
  const cellParagraph = paragraph("cell");
  const following = table([
    row([schema.node("tableCell", null, [cellParagraph])]),
  ]);
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { pPrMark: deletedParagraphMark }, [
      schema.text("kept"),
    ]),
    following,
  ]);

  const result = resolveWholeStory({
    doc,
    mode: "accept",
    styleResolver: null,
    numbering: null,
  });

  expect(result.failed).toBe(false);
  expect(result.resolved.childCount).toBe(1);
  const survivor =
    result.resolved.firstChild?.firstChild?.firstChild?.firstChild;
  expect(survivor?.textContent).toBe("keptcell");
  expect(survivor?.attrs["paraId"]).toBe(cellParagraph.attrs["paraId"]);
  expect(survivor?.attrs["pPrMark"]).toBeNull();
});

const insertedParagraphMark = {
  kind: "ins",
  info: { id: 8, author: "Reviewer", date: "2026-10-21" },
} as const;

test("rejecting an empty inserted final paragraph after a table keeps it as the terminator", () => {
  const terminator = schema.node("paragraph", {
    pPrMark: insertedParagraphMark,
  });
  const doc = schema.node("doc", null, [
    paragraph("before"),
    table([row([cell("cell")])]),
    terminator,
  ]);

  const result = resolveWholeStory({
    doc,
    mode: "reject",
    styleResolver: null,
    numbering: null,
  });

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

  const result = resolveWholeStory({
    doc,
    mode: "reject",
    styleResolver: null,
    numbering: null,
  });

  expect(result.failed).toBe(false);
  expect(result.resolved.lastChild?.textContent).toBe("beforeafter");
  expect(result.resolved.lastChild?.attrs["_sectionProperties"]).toEqual(
    sectionProperties,
  );
});

test("rejecting an empty inserted final paragraph in a cell leaves the paragraph before it last", () => {
  const terminator = schema.node("paragraph", {
    pPrMark: insertedParagraphMark,
  });
  const doc = schema.node("doc", null, [
    table([
      row([schema.node("tableCell", null, [paragraph("text"), terminator])]),
    ]),
    paragraph("after"),
  ]);

  const result = resolveWholeStory({
    doc,
    mode: "reject",
    styleResolver: null,
    numbering: null,
  });

  expect(result.failed).toBe(false);
  const resolvedCell = result.resolved.firstChild?.firstChild?.firstChild;
  expect(resolvedCell?.childCount).toBe(1);
  expect(resolvedCell?.lastChild?.textContent).toBe("text");
});

test("rejecting the only paragraph of a cell keeps it as the cell terminator", () => {
  const terminator = schema.node("paragraph", {
    pPrMark: insertedParagraphMark,
  });
  const doc = schema.node("doc", null, [
    table([row([schema.node("tableCell", null, [terminator])])]),
    paragraph("after"),
  ]);

  const result = resolveWholeStory({
    doc,
    mode: "reject",
    styleResolver: null,
    numbering: null,
  });

  expect(result.failed).toBe(false);
  const resolvedCell = result.resolved.firstChild?.firstChild?.firstChild;
  expect(resolvedCell?.childCount).toBe(1);
  expect(resolvedCell?.firstChild?.type.name).toBe("paragraph");
});
