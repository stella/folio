import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import {
  CanonicalSessionError,
  createCanonicalSession,
  publishCanonicalProjection,
  type CanonicalCommit,
  type CanonicalInputSemantic,
  type CanonicalSession,
} from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(120_000));

const seed = (runLength = 50): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          textId: "87654321",
          content: [
            {
              type: "run",
              formatting: { bold: true },
              content: [{ type: "text", text: "a".repeat(runLength) }],
            },
            {
              type: "run",
              formatting: { italic: true },
              content: [{ type: "text", text: "b".repeat(runLength) }],
            },
          ],
        },
      ],
    },
  },
});

const accept = (session: CanonicalSession, state: EditorState, commit: CanonicalCommit) =>
  publishCanonicalProjection({ session, state, commit }).unwrap().state;

const semanticCases = {
  typing: "typing",
  deleteBackward: "deleteBackward",
  deleteForward: "deleteForward",
  paste: "paste",
  composition: "composition",
  replacement: "replacement",
  structure: "structure",
} as const satisfies { [Semantic in CanonicalInputSemantic]: Semantic };
const semantics = Object.values(semanticCases);

const actionArbitrary = fc.record({
  semantic: fc.constantFrom(...semantics),
  delay: fc.constantFrom(0, 1, 499, 500, 501, -1),
  selection: fc.constantFrom("keep", "move", "roundtrip", "range"),
  offset: fc.nat(100),
  text: fc.constantFrom("X", "中", "😀"),
});

describe("canonical semantic undo partition", () => {
  test("generated intent sequences restore every group exactly without projection adjacency", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(actionArbitrary, { minLength: 20, maxLength: 40 }),
        async (generated) => {
          // Exercise the same trace from empty, short and multi-run paragraphs.
          for (const runLength of [0, 1, 50]) {
            const session = createCanonicalSession(seed(runLength)).unwrap();
            let state = EditorState.create({ schema, doc: session.projection.doc });
            const initial = session.document;
            const groups: {
              before: Document;
              after: Document;
              preSelection: ReturnType<typeof state.selection.toJSON>;
              postSelection: ReturnType<typeof state.selection.toJSON>;
            }[] = [];
            let previous:
              | { semantic: CanonicalInputSemantic; eligible: boolean; time: number; caret: number }
              | undefined;
            let time = 1_000;
            // Every declared intent is exercised in each trace, followed by arbitrary transitions.
            const actions = [
              ...(
                [
                  "deleteBackward",
                  "deleteForward",
                  "typing",
                  "deleteBackward",
                  "deleteForward",
                ] as const
              ).map((semantic) => ({
                semantic,
                delay: 1,
                selection: "roundtrip",
                offset: 0,
                text: "X",
              })),
              ...semantics.map((semantic) => ({
                semantic,
                delay: 1,
                selection: "keep",
                offset: 0,
                text: "X",
              })),
              ...generated,
            ];
            const exercised = new Set<CanonicalInputSemantic>();
            for (const action of actions) {
              exercised.add(action.semantic);
              time += action.delay;
              const paragraph =
                session.projection.paragraph("12345678") ?? panic("Missing test paragraph");
              const gaps = [paragraph.start];
              let position = paragraph.start;
              for (const unit of paragraph.text) {
                position += unit.length;
                gaps.push(position);
              }
              const selected =
                action.selection !== "keep" &&
                (action.selection !== "roundtrip" || gaps.length > 1);
              let caret = state.selection.head;
              if (action.selection === "move" || action.selection === "range") {
                caret = gaps.at(action.offset % gaps.length) ?? panic("Missing generated caret");
                const head =
                  action.selection === "range"
                    ? (gaps.at((action.offset + 1) % gaps.length) ??
                      panic("Missing generated head"))
                    : caret;
                state = state.apply(
                  state.tr.setSelection(TextSelection.create(state.doc, caret, head)),
                );
              }
              if (action.selection === "roundtrip" && gaps.length > 1) {
                const previousCaret = state.selection.head;
                const away =
                  gaps.find((gap) => gap !== previousCaret) ?? panic("Missing roundtrip gap");
                state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, away)));
                state = state.apply(
                  state.tr.setSelection(TextSelection.create(state.doc, previousCaret)),
                );
              }
              if (selected) session.breakUndoGroup();
              const preSelection = state.selection.toJSON();
              const before = session.document;
              let from = state.selection.from;
              let to = state.selection.to;
              const text =
                action.semantic === "deleteBackward" || action.semantic === "deleteForward"
                  ? ""
                  : action.text;
              if (text.length === 0 && gaps.length === 1) {
                const version = session.version;
                const projection = session.projection;
                const canUndo = session.canUndo;
                const canRedo = session.canRedo;
                const refused = session.prepareReplace(state, {
                  from,
                  to,
                  text,
                  semantic: action.semantic,
                  time,
                });
                if (refused.isOk()) panic("Empty deletion unexpectedly staged a commit.");
                expect(refused.error.message).toBe("The input makes no text change.");
                expect(session.document).toStrictEqual(before);
                expect(session.projection).toBe(projection);
                expect(session.version).toBe(version);
                expect(session.canUndo).toBe(canUndo);
                expect(session.canRedo).toBe(canRedo);
                expect(state.selection.toJSON()).toEqual(preSelection);
                if (selected) previous = undefined;
                continue;
              }
              if (text.length === 0 && state.selection.empty) {
                const index = gaps.indexOf(caret);
                if (action.semantic === "deleteBackward") {
                  if (index === 0) {
                    caret = gaps.at(1) ?? panic("Missing backward caret");
                    state = state.apply(
                      state.tr.setSelection(TextSelection.create(state.doc, caret)),
                    );
                  }
                  to = caret;
                  from =
                    gaps.at(Math.max(0, gaps.indexOf(caret) - 1)) ?? panic("Missing backward gap");
                } else {
                  if (index === gaps.length - 1) {
                    caret = gaps.at(-2) ?? panic("Missing forward caret");
                    state = state.apply(
                      state.tr.setSelection(TextSelection.create(state.doc, caret)),
                    );
                  }
                  from = caret;
                  to = gaps.at(gaps.indexOf(caret) + 1) ?? panic("Missing forward gap");
                }
              }
              const actualPreSelection = state.selection.toJSON();
              const eligible =
                state.selection.empty &&
                (action.semantic === "typing" ||
                  action.semantic === "deleteBackward" ||
                  action.semantic === "deleteForward");
              const joins =
                previous !== undefined &&
                eligible &&
                previous.eligible &&
                previous.semantic === action.semantic &&
                !selected &&
                previous.caret === state.selection.head &&
                time >= previous.time &&
                time - previous.time <= 500;
              state = accept(
                session,
                state,
                session
                  .prepareReplace(state, { from, to, text, semantic: action.semantic, time })
                  .unwrap(),
              );
              const after = session.document;
              const postSelection = state.selection.toJSON();
              if (joins) {
                const group = groups.at(-1) ?? panic("Missing expected group");
                group.after = after;
                group.postSelection = postSelection;
              } else {
                groups.push({ before, after, preSelection: actualPreSelection, postSelection });
              }
              previous = { semantic: action.semantic, eligible, time, caret: state.selection.head };
              // Range direction is part of the selection oracle, not a min/max-only address.
              if (action.selection === "range") expect(actualPreSelection).toEqual(preSelection);
            }
            expect([...exercised].sort()).toEqual([...semantics].sort());
            for (const group of groups.toReversed()) {
              state = accept(session, state, session.prepareUndo(state).unwrap());
              expect(session.document).toStrictEqual(group.before);
              expect(state.selection.toJSON()).toEqual(group.preSelection);
            }
            expect(session.canUndo).toBe(false);
            expect(session.document).toStrictEqual(initial);
            for (const group of groups) {
              state = accept(session, state, session.prepareRedo(state).unwrap());
              expect(session.document).toStrictEqual(group.after);
              expect(state.selection.toJSON()).toEqual(group.postSelection);
            }
            expect(session.canRedo).toBe(false);
          }
        },
      ),
      { numRuns: 75 },
    );
  });

  test("history replay and a caret roundtrip close an otherwise adjacent typing run", () => {
    const session = createCanonicalSession(seed()).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const type = (text: string, time: number) => {
      state = accept(
        session,
        state,
        session
          .prepareReplace(state, {
            from: state.selection.head,
            to: state.selection.head,
            text,
            semantic: "typing",
            time,
          })
          .unwrap(),
      );
    };
    type("X", 0);
    const first = session.document;
    const previousCaret = state.selection.head;
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, previousCaret + 1)));
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, previousCaret)));
    session.breakUndoGroup();
    type("Y", 1);
    state = accept(session, state, session.prepareUndo(state).unwrap());
    expect(session.document).toStrictEqual(first);
    state = accept(session, state, session.prepareRedo(state).unwrap());
    const replayed = session.document;
    type("Z", 2);
    state = accept(session, state, session.prepareUndo(state).unwrap());
    expect(session.document).toStrictEqual(replayed);
  });

  test("composition replacement is isolated and replays exact identities and reversed selection", () => {
    const session = createCanonicalSession(seed()).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    state = accept(
      session,
      state,
      session
        .prepareReplace(state, {
          from: 1,
          to: 1,
          text: "X",
          semantic: "typing",
          time: 0,
        })
        .unwrap(),
    );
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 54, 48)));
    const baseline = session.document;
    const selected = state.selection.toJSON();
    expect(session.beginComposition().isOk()).toBe(true);
    expect(session.prepareUndo(state).isErr()).toBe(true);
    session.endComposition();
    state = accept(
      session,
      state,
      session
        .prepareReplace(state, {
          from: 48,
          to: 54,
          text: "中文😀",
          semantic: "composition",
          time: 1,
        })
        .unwrap(),
    );
    const committed = session.document;
    const caret = state.selection.toJSON();
    state = accept(
      session,
      state,
      session
        .prepareReplace(state, {
          from: state.selection.head,
          to: state.selection.head,
          text: "Y",
          semantic: "typing",
          time: 2,
        })
        .unwrap(),
    );
    state = accept(session, state, session.prepareUndo(state).unwrap());
    expect(session.document).toStrictEqual(committed);
    state = accept(session, state, session.prepareUndo(state).unwrap());
    expect(session.document).toStrictEqual(baseline);
    expect(state.selection.toJSON()).toEqual(selected);
    state = accept(session, state, session.prepareRedo(state).unwrap());
    expect(session.document).toStrictEqual(committed);
    expect(state.selection.toJSON()).toEqual(caret);
  });

  test("pending composition blocks snapshots, history and stale staged commits", () => {
    const session = createCanonicalSession(seed()).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const baseline = session.document;
    const prepared = session.prepareReplace(state, { from: 1, to: 1, text: "X" }).unwrap();
    expect(session.beginComposition().isOk()).toBe(true);
    expect(session.beginComposition().isErr()).toBe(true);
    expect(() => session.document).toThrow(CanonicalSessionError);
    expect(session.prepareReplace(state, { from: 1, to: 1, text: "X" }).isErr()).toBe(true);
    expect(session.prepareUndo(state).isErr()).toBe(true);
    expect(session.prepareRedo(state).isErr()).toBe(true);
    expect(prepared.publish().isErr()).toBe(true);
    session.endComposition();
    expect(prepared.publish().isErr()).toBe(true);
    expect(session.document).toStrictEqual(baseline);
    state = accept(
      session,
      state,
      session
        .prepareReplace(state, { from: 1, to: 1, text: "中", semantic: "composition" })
        .unwrap(),
    );
    state = accept(session, state, session.prepareUndo(state).unwrap());
    expect(session.document).toStrictEqual(baseline);
    expect(state.selection.from).toBe(1);
  });
});
