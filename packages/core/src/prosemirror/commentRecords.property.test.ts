import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { redo, undo } from "prosemirror-history";
import { type EditorState, TextSelection } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createHarnessState } from "../__tests__/editorHarness";
import { withHostCommentThreads } from "../docx/hostCommentThreads";
import type { Comment, Paragraph, ParagraphContent } from "../types/content";
import type { Document } from "../types/document";
import { acceptAllChanges, rejectAllChanges } from "./commands/comments";
import {
  addCommentThread,
  appendCommentRecords,
  commentIdsNamedIn,
  commentRecordsOf,
  createCommentRecordsPlugin,
  deleteCommentThread,
  setCommentResolved,
} from "./commentRecords";
import { fromProseDoc } from "./conversion/fromProseDoc";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const commentBody = (text: string): Paragraph[] => [
  { type: "paragraph", content: [run(text)] },
];

const comment = (id: number, text: string, parentId?: number): Comment => ({
  id,
  author: "Reviewer",
  date: "2024-01-01T00:00:00Z",
  content: commentBody(text),
  ...(parentId === undefined ? {} : { parentId }),
});

/** Three paragraphs: the second carries thread 1 (with reply 2). */
const commentedDocument = (): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", content: [run("Opening words")] },
        {
          type: "paragraph",
          content: [
            { type: "commentRangeStart", id: 1 },
            run("Please remit funds."),
            { type: "commentRangeEnd", id: 1 },
            { type: "commentReference", id: 1 },
          ],
        },
        { type: "paragraph", content: [run("Closing words")] },
      ],
      comments: [comment(1, "Check the account"), comment(2, "A reply", 1)],
    },
  },
});

const stateFor = (document: Document, mode: "editing" | "suggesting") =>
  createHarnessState(document, mode, [
    createCommentRecordsPlugin(document.package.document.comments ?? []),
  ]);

/** What a save writes for comments: their records and their anchors. */
const saved = (state: EditorState, base: Document) => {
  const records = commentRecordsOf(state) ?? [];
  const document = withHostCommentThreads(
    fromProseDoc(state.doc, base, { stylesheetSource: { type: "package" } }),
    records,
    { anchoredInEditor: commentIdsNamedIn(state.doc) },
  );
  return {
    comments: JSON.stringify(document.package.document.comments ?? []),
    anchored: [...commentIdsNamedIn(state.doc)].toSorted((a, b) => a - b),
  };
};

type Op =
  | { type: "add"; paragraph: number }
  | { type: "reply"; target: number }
  | { type: "resolve"; target: number; done: boolean }
  | { type: "delete"; target: number }
  | { type: "type"; paragraph: number; text: string };

const opArbitrary: fc.Arbitrary<Op> = fc.oneof(
  fc.record({
    type: fc.constant("add" as const),
    paragraph: fc.nat({ max: 2 }),
  }),
  fc.record({ type: fc.constant("reply" as const), target: fc.nat() }),
  fc.record({
    type: fc.constant("resolve" as const),
    target: fc.nat(),
    done: fc.boolean(),
  }),
  fc.record({ type: fc.constant("delete" as const), target: fc.nat() }),
  fc.record({
    type: fc.constant("type" as const),
    paragraph: fc.nat({ max: 2 }),
    text: fc.constantFrom("x", "new ", "words"),
  }),
);

/** The range of the `index`th paragraph's text. */
const paragraphRange = (state: EditorState, index: number) => {
  let found: { from: number; to: number } | null = null;
  let seen = 0;
  state.doc.descendants((node, position) => {
    if (found || node.type.name !== "paragraph") return !found;
    if (seen === index) {
      found = { from: position + 1, to: position + node.nodeSize - 1 };
    }
    seen += 1;
    return false;
  });
  return found ?? { from: 1, to: 1 };
};

let nextId = 100;

/** Apply `op`, returning the next state, or null when it does nothing. */
const applyOp = (state: EditorState, op: Op): EditorState | null => {
  const records = commentRecordsOf(state) ?? [];
  const pick = (target: number, rootsOnly: boolean) => {
    const pool = rootsOnly
      ? records.filter(({ parentId }) => parentId === undefined)
      : records;
    return pool.length === 0 ? undefined : pool[target % pool.length];
  };
  const tr = (() => {
    switch (op.type) {
      case "add": {
        nextId += 1;
        return addCommentThread(
          state,
          paragraphRange(state, op.paragraph),
          comment(nextId, `Note ${String(nextId)}`),
        );
      }
      case "reply": {
        const parent = pick(op.target, false);
        if (!parent) return null;
        nextId += 1;
        return appendCommentRecords(state, [
          comment(nextId, `Reply ${String(nextId)}`, parent.id),
        ]);
      }
      case "resolve": {
        const target = pick(op.target, true);
        return target ? setCommentResolved(state, target.id, op.done) : null;
      }
      case "delete": {
        const target = pick(op.target, true);
        return target ? deleteCommentThread(state, target.id) : null;
      }
      case "type": {
        const range = paragraphRange(state, op.paragraph);
        return state.tr
          .setSelection(TextSelection.create(state.doc, range.to))
          .insertText(op.text);
      }
      default: {
        const unreachable: never = op;
        return unreachable;
      }
    }
  })();
  if (!tr) return null;
  return state.apply(tr);
};

const historyStep = (
  state: EditorState,
  command: typeof undo,
): EditorState | null => {
  let next: EditorState | null = null;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
};

describe("comment records", () => {
  test(
    "undo after any comment edit restores the saved comments exactly",
    () => {
      assertProperty(
        fc.property(
          fc.constantFrom("editing" as const, "suggesting" as const),
          fc.array(opArbitrary, { minLength: 1, maxLength: 8 }),
          (mode, ops) => {
            const base = commentedDocument();
            let state = stateFor(base, mode);
            for (const op of ops) {
              const next = applyOp(state, op);
              if (!next) continue;
              if (op.type === "type") {
                state = next;
                continue;
              }
              const before = saved(state, base);
              const after = saved(next, base);
              const undone = historyStep(next, undo);
              expect(undone).not.toBeNull();
              if (!undone) return;
              expect(saved(undone, base)).toEqual(before);
              const redone = historyStep(undone, redo);
              expect(redone).not.toBeNull();
              if (!redone) return;
              expect(saved(redone, base)).toEqual(after);
              state = redone;
            }
          },
        ),
        {
          seed: 4_412_208,
          numRuns: 120,
          id: "undo after any comment edit restores the saved comments exactly",
        },
      );
    },
    propertyTestTimeout(60_000),
  );

  test("deleting a thread removes its replies and anchors, and undo brings all back", () => {
    const base = commentedDocument();
    const state = stateFor(base, "editing");
    const deleted = deleteCommentThread(state, 1);
    expect(deleted).not.toBeNull();
    if (!deleted) return;
    const after = state.apply(deleted);
    expect(commentRecordsOf(after)).toEqual([]);
    expect([...commentIdsNamedIn(after.doc)]).toEqual([]);
    const undone = historyStep(after, undo);
    expect(undone).not.toBeNull();
    if (!undone) return;
    expect(commentRecordsOf(undone)).toEqual(
      base.package.document.comments ?? [],
    );
    expect([...commentIdsNamedIn(undone.doc)]).toEqual([1]);
    expect(saved(undone, base)).toEqual(saved(state, base));
  });

  test("rejecting the suggestion that inserted a commented passage retires the comment", () => {
    const base = commentedDocument();
    let state = stateFor(base, "suggesting");
    const range = paragraphRange(state, 0);
    state = state.apply(
      state.tr
        .setSelection(TextSelection.create(state.doc, range.to))
        .insertText(" and more"),
    );
    const inserted = { from: range.to, to: range.to + " and more".length };
    const added = addCommentThread(
      state,
      inserted,
      comment(70, "On the insertion"),
    );
    expect(added).not.toBeNull();
    if (!added) return;
    state = state.apply(added);
    const before = commentRecordsOf(state);
    let rejected: EditorState | null = null;
    rejectAllChanges()(state, (tr) => {
      rejected = state.apply(tr);
    });
    expect(rejected).not.toBeNull();
    if (!rejected) return;
    const after: EditorState = rejected;
    expect(commentRecordsOf(after)?.map(({ id }) => id)).toEqual([1, 2]);
    expect(saved(after, base).comments).not.toContain("On the insertion");
    const undone = historyStep(after, undo);
    expect(undone && commentRecordsOf(undone)).toEqual(before);
  });

  test("accepting the deletion of everything a comment covers retires its thread", () => {
    const base = commentedDocument();
    let state = stateFor(base, "suggesting");
    const range = paragraphRange(state, 1);
    state = state.apply(
      state.tr
        .setSelection(TextSelection.create(state.doc, range.from, range.to))
        .deleteSelection(),
    );
    // Rejecting the deletion keeps the thread.
    rejectAllChanges()(state, (tr) => {
      expect(commentRecordsOf(state.apply(tr))?.map(({ id }) => id)).toEqual([
        1, 2,
      ]);
    });
    let accepted: EditorState | null = null;
    acceptAllChanges()(state, (tr) => {
      accepted = state.apply(tr);
    });
    expect(accepted).not.toBeNull();
    if (!accepted) return;
    expect(commentRecordsOf(accepted)).toEqual([]);
    const undone = historyStep(accepted, undo);
    expect(undone && commentRecordsOf(undone)?.map(({ id }) => id)).toEqual([
      1, 2,
    ]);
  });
});
