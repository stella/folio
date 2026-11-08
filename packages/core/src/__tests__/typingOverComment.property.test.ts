/**
 * Typing over exactly the text a comment covers keeps the typed text inside
 * the comment: in editing the comment covers the new text, and
 * in suggesting accepting leaves it on the new text and rejecting on the old.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { EditorState } from "prosemirror-state";
import { AllSelection, TextSelection } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Comment, ParagraphContent } from "../types/content";
import type { Document } from "../types/document";
import {
  createHarnessState,
  type EditorMode,
  HeadlessEditorView,
  resolveAllChanges,
} from "./editorHarness";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const comment = (id: number): Comment => ({
  id,
  author: "Reviewer",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [run(`Note ${String(id)}`)] }],
});

type Layout = {
  before: string;
  target: string;
  after: string;
  /** A second comment opening with the first and closing past its reference. */
  overlapping: boolean;
};

const documentFor = ({
  before,
  target,
  after,
  overlapping,
}: Layout): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            ...(before ? [run(before)] : []),
            { type: "commentRangeStart", id: 1 },
            ...(overlapping
              ? [{ type: "commentRangeStart", id: 2 } as const]
              : []),
            run(target),
            { type: "commentRangeEnd", id: 1 },
            { type: "commentReference", id: 1 },
            ...(overlapping
              ? [
                  { type: "commentRangeEnd", id: 2 } as const,
                  { type: "commentReference", id: 2 } as const,
                ]
              : []),
            ...(after ? [run(after)] : []),
          ],
        },
      ],
      comments: overlapping ? [comment(1), comment(2)] : [comment(1)],
    },
  },
});

/** The text each comment's mark covers, by comment id. */
const commentedText = (state: EditorState): Map<number, string> => {
  const texts = new Map<number, string>();
  state.doc.descendants((node) => {
    if (!node.isText) return true;
    for (const mark of node.marks) {
      const id: unknown = mark.attrs["commentId"];
      if (mark.type.name !== "comment" || typeof id !== "number") continue;
      texts.set(id, (texts.get(id) ?? "") + (node.text ?? ""));
    }
    return false;
  });
  return texts;
};

/** Each comment's anchored text in the saved document. */
const savedCommentedText = (state: EditorState, base: Document) => {
  const saved = fromProseDoc(state.doc, base, {
    stylesheetSource: { type: "package" },
  });
  const texts = new Map<number, string>();
  const open = new Set<number>();
  for (const block of saved.package.document.content) {
    if (block.type !== "paragraph") continue;
    for (const item of block.content) {
      if (item.type === "commentRangeStart") open.add(item.id);
      if (item.type === "commentRangeEnd") open.delete(item.id);
      if (item.type !== "run") continue;
      for (const piece of item.content) {
        if (piece.type !== "text") continue;
        for (const id of open)
          texts.set(id, (texts.get(id) ?? "") + piece.text);
      }
    }
  }
  return texts;
};

/** Select exactly the commented text and type `typed` a key at a time. */
const typeOverComment = (layout: Layout, mode: EditorMode, typed: string) => {
  const base = documentFor(layout);
  const state = createHarnessState(base, mode);
  let range: { from: number; to: number } | null = null;
  state.doc.descendants((node, position) => {
    if (!range && node.isText && node.text === layout.target)
      range = { from: position, to: position + node.nodeSize };
    return !range;
  });
  if (!range) throw new Error("The fixture lost its commented text");
  const { from, to }: { from: number; to: number } = range;
  const view = new HeadlessEditorView(
    state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, from, to)),
    ),
  );
  view.typeText(typed);
  return { base, state: view.state };
};

const words = fc.constantFrom("Alpha", "beta gamma", "x", "Delta.");
const layouts = fc.record({
  before: fc.constantFrom("", "Lead ", "Opening words. "),
  target: fc.constantFrom("Please remit funds.", "the account", "Q"),
  after: fc.constantFrom("", " tail", " and more."),
  overlapping: fc.boolean(),
}) satisfies fc.Arbitrary<Layout>;

describe("typing over the text a comment covers", () => {
  test(
    "editing: the comment covers the typed text, also after a save",
    () => {
      assertProperty(
        fc.property(layouts, words, (layout, typed) => {
          const { base, state } = typeOverComment(layout, "editing", typed);
          const expected = layout.overlapping
            ? [
                [1, typed],
                [2, typed],
              ]
            : [[1, typed]];
          expect([...commentedText(state)].toSorted()).toEqual(expected);
          expect([...savedCommentedText(state, base)].toSorted()).toEqual(
            expected,
          );
        }),
        {
          seed: 5_108_221,
          numRuns: 60,
          id: "editing: the comment covers the typed text, also after a save",
        },
      );
    },
    propertyTestTimeout(60_000),
  );

  test(
    "suggesting: accepting keeps the comment on the typed text, rejecting on the original",
    () => {
      assertProperty(
        fc.property(layouts, words, (layout, typed) => {
          const { base, state } = typeOverComment(layout, "suggesting", typed);
          const ids = layout.overlapping ? [1, 2] : [1];
          const accepted = resolveAllChanges(state, "accept");
          expect([...commentedText(accepted)].toSorted()).toEqual(
            ids.map((id) => [id, typed]),
          );
          expect([...savedCommentedText(accepted, base)].toSorted()).toEqual(
            ids.map((id) => [id, typed]),
          );
          const rejected = resolveAllChanges(state, "reject");
          expect([...commentedText(rejected)].toSorted()).toEqual(
            ids.map((id) => [id, layout.target]),
          );
        }),
        {
          seed: 5_108_222,
          numRuns: 60,
          id: "suggesting: accepting keeps the comment on the typed text, rejecting on the original",
        },
      );
    },
    propertyTestTimeout(60_000),
  );

  test("typing over a select-all of a document the comment covers keeps it on the typed text", () => {
    const layout: Layout = {
      before: "",
      target: "Please remit funds.",
      after: "",
      overlapping: false,
    };
    for (const mode of ["editing", "suggesting"] as const) {
      const base = documentFor(layout);
      const state = createHarnessState(base, mode);
      const view = new HeadlessEditorView(
        state.apply(state.tr.setSelection(new AllSelection(state.doc))),
      );
      view.typeText("New");
      const result =
        mode === "editing"
          ? view.state
          : resolveAllChanges(view.state, "accept");
      expect([mode, commentedText(result).get(1)]).toEqual([mode, "New"]);
      expect(savedCommentedText(result, base).get(1)).toBe("New");
    }
  });

  test("typing over a comment's text and content it does not cover leaves the comment off the typed text", () => {
    const base: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                { type: "commentRangeStart", id: 1 },
                run("Covered"),
                { type: "commentRangeEnd", id: 1 },
                { type: "commentReference", id: 1 },
                { type: "run", content: [{ type: "tab" }] },
                run("tail"),
              ],
            },
          ],
          comments: [comment(1)],
        },
      },
    };
    for (const mode of ["editing", "suggesting"] as const) {
      const state = createHarnessState(base, mode);
      let from = -1;
      let tab = -1;
      state.doc.descendants((node, position) => {
        if (node.isText && node.text === "Covered") from = position;
        if (node.type.name === "tab") tab = position;
        return true;
      });
      expect(from).toBeGreaterThan(0);
      expect(tab).toBeGreaterThan(from);
      const view = new HeadlessEditorView(
        state.apply(
          state.tr.setSelection(TextSelection.create(state.doc, from, tab + 1)),
        ),
      );
      view.typeText("New");
      const result =
        mode === "editing"
          ? view.state
          : resolveAllChanges(view.state, "accept");
      expect(commentedText(result).get(1)).toBeUndefined();
    }
  });

  test("text a composition commits at the caret joins the comment the typing is in", () => {
    const layout: Layout = {
      before: "Lead ",
      target: "Please remit funds.",
      after: " tail",
      overlapping: false,
    };
    const { base, state } = typeOverComment(layout, "editing", "Al");
    const view = new HeadlessEditorView(state);
    // The browser commits composed text into the document directly, without
    // the input handlers.
    view.dispatch(view.state.tr.insertText("pha", view.state.selection.from));
    view.typeText("!");
    expect(commentedText(view.state).get(1)).toBe("Alpha!");
    expect(savedCommentedText(view.state, base).get(1)).toBe("Alpha!");
  });

  test("moving the caret ends typing inside the comment", () => {
    const layout: Layout = {
      before: "Lead ",
      target: "Please remit funds.",
      after: " tail",
      overlapping: false,
    };
    const { state } = typeOverComment(layout, "editing", "New");
    const view = new HeadlessEditorView(state);
    const caret = view.state.selection.from;
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, caret + 2),
      ),
    );
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, caret)),
    );
    view.typeText("!");
    expect(commentedText(view.state).get(1)).toBe("New");
  });
});
