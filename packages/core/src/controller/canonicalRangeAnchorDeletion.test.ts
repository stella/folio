import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import { editorParagraphGroups, OP_STORIES } from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  CANONICAL_PARAGRAPH_SHAPE_FACTORIES,
  canonicalInlineShapeDocument,
} from "../../typecheck/canonical-inline-shapes.typecheck";
import {
  createCanonicalSession,
  deletionRange,
  measuredDeletionRange,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import type { Document, Paragraph, Run } from "../types/document";

setDefaultTimeout(propertyTestTimeout(60000));

const run = (text: string) =>
  ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;
const visibleText = (document: Document) =>
  editorParagraphGroups(document, OP_STORIES.MAIN)
    .map(({ text }) => text)
    .join("");

test("grapheme deletion retains source range anchors in both directions through review and history", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(
        "bookmarkStart",
        "moveFromRangeStart",
        "moveToRangeStart",
      ),
      fc.constantFrom("e", "a", "n"),
      fc.constantFrom("\u0301", "\u0308", "\u0303"),
      (kind, base, combining) => {
        const fixture = CANONICAL_PARAGRAPH_SHAPE_FACTORIES[kind]().at(0);
        if (!fixture) panic("Source range fixture is absent");
        const [start, end] = fixture.content;
        if (!start || !end)
          panic("Source range fixture lost its paired markers");
        const document = canonicalInlineShapeDocument({
          story: "main",
          content: [run(base), start, run(combining), end],
        });
        for (const direction of ["backward", "forward"] as const) {
          const session = createCanonicalSession(document).unwrap();
          session.setMode({ type: "suggesting", author: "Reviewer" });
          const baseline = session.document;
          const caret = session.projection
            .positionAt({
              story: OP_STORIES.MAIN,
              blockId: "12345678",
              offset:
                direction === "forward"
                  ? 1
                  : 1 + base.length + combining.length,
              zeroWidthBefore: 0,
            })
            .unwrap();
          let state = EditorState.create({
            doc: session.projection.doc,
            selection: TextSelection.create(session.projection.doc, caret),
          });
          const range = deletionRange(state, direction).unwrap();
          const commit = session
            .prepareReplace(state, {
              ...range,
              text: "",
              semantic:
                direction === "backward" ? "deleteBackward" : "deleteForward",
            })
            .unwrap();
          state = publishCanonicalProjection({
            session,
            state,
            commit,
          }).unwrap().state;
          expect(state.doc.eq(session.projection.doc)).toBe(true);
          expect(visibleText(session.document)).toBe("LR");
          const suggested = session.document;
          const paragraph = suggested.package.document.content.at(0);
          if (paragraph?.type !== "paragraph")
            panic("Grapheme deletion lost its paragraph");
          const markers = (source: Paragraph) =>
            source.content.filter(
              (item) => item.type === start.type || item.type === end.type,
            );
          expect(markers(paragraph)).toEqual(fixture.content);
          const revisionIds = paragraph.content.flatMap((item) =>
            item.type === "deletion" ? [item.info.id] : [],
          );
          expect(revisionIds.length).toBeGreaterThan(0);
          for (const resolution of ["accept", "reject"] as const) {
            const resolved = createCanonicalSession(suggested).unwrap();
            const before = EditorState.create({ doc: resolved.projection.doc });
            const resolutionCommit = resolved
              .prepareResolve(before, { revisionIds, resolution })
              .unwrap();
            publishCanonicalProjection({
              session: resolved,
              state: before,
              commit: resolutionCommit,
            }).unwrap();
            expect(visibleText(resolved.document)).toBe(
              resolution === "accept" ? "LR" : `L${base}${combining}R`,
            );
            const resolvedParagraph =
              resolved.document.package.document.content.at(0);
            if (resolvedParagraph?.type !== "paragraph")
              panic("Review lost its paragraph");
            expect(markers(resolvedParagraph)).toEqual(fixture.content);
          }
          const undo = session.prepareUndo(state).unwrap();
          state = publishCanonicalProjection({
            session,
            state,
            commit: undo,
          }).unwrap().state;
          expect(session.document).toStrictEqual(baseline);
          const redo = session.prepareRedo(state).unwrap();
          state = publishCanonicalProjection({
            session,
            state,
            commit: redo,
          }).unwrap().state;
          expect(session.document).toStrictEqual(suggested);
          expect(state.doc.eq(session.projection.doc)).toBe(true);
        }
      },
    ),
    {
      seed: 197,
      numRuns: 30,
      id: "grapheme deletion retains source range anchors in both directions through review and history",
    },
  );
});

test.each([
  ["bookmarkStart", "editing"],
  ["bookmarkStart", "suggesting"],
  ["moveToRangeStart", "editing"],
  ["moveToRangeStart", "suggesting"],
] as const)(
  "a measured word deletion ending past a %s's markers deletes the word and keeps them, in %s",
  (kind, mode) => {
    const fixture = CANONICAL_PARAGRAPH_SHAPE_FACTORIES[kind]().at(0);
    if (!fixture) panic("Source range fixture is absent");
    const [start, end] = fixture.content;
    if (!start || !end) panic("Source range fixture lost its paired markers");
    const document = canonicalInlineShapeDocument({
      story: "main",
      content: [run("heading "), start, run("text"), end],
    });
    const session = createCanonicalSession(document).unwrap();
    if (mode === "suggesting")
      session.setMode({ type: "suggesting", author: "Reviewer" });
    const doc = session.projection.doc;
    const anchors: { pos: number; size: number }[] = [];
    doc.descendants((node, pos) => {
      if (node.isInline && !node.isText)
        anchors.push({ pos, size: node.nodeSize });
    });
    const first = anchors.at(0);
    const last = anchors.at(-1);
    if (!first || !last) panic("The markers are not in the projection");
    // The browser measures the word from before the opening marker to after
    // the closing one, the caret at the end of the line as End puts it.
    const measured = { from: first.pos, to: last.pos + last.size };
    const state = EditorState.create({
      doc,
      selection: TextSelection.create(doc, measured.to),
    });
    const range = measuredDeletionRange(state, measured).unwrap();
    if (range === null) panic("The measured word covers no character");
    expect(doc.textBetween(range.from, range.to)).toBe("text");
    expect(range.from).toBeGreaterThan(measured.from);
    expect(range.to).toBeLessThan(measured.to);
    const commit = session
      .prepareReplace(state, { ...range, text: "", semantic: "deleteBackward" })
      .unwrap();
    publishCanonicalProjection({ session, state, commit }).unwrap();
    expect(visibleText(session.document)).toBe(
      visibleText(document).replace("text", ""),
    );
    const paragraph = session.document.package.document.content.at(0);
    if (paragraph?.type !== "paragraph")
      panic("The deletion lost its paragraph");
    expect(
      paragraph.content.filter(
        (item) => item.type === start.type || item.type === end.type,
      ),
    ).toEqual(fixture.content);
  },
);
