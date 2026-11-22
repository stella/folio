/**
 * One range per comment id, pasting in a ProseMirror session: a pasted copy
 * of commented text takes a fresh comment id and a copy of the thread, or
 * leaves the comment out when the session holds no thread for it; a paste of
 * a comment the document does not name goes in as it is. Each case runs in
 * editing and in suggesting, and saves one range per comment.
 */
import { describe, expect, test } from "bun:test";
import { undo } from "prosemirror-history";
import type { Slice } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";

import {
  createHarnessState,
  EDITOR_MODES,
  type EditorMode,
  findTextblock,
  HeadlessEditorView,
} from "../__tests__/editorHarness";
import type { Comment, ParagraphContent } from "../types/content";
import type { Document } from "../types/document";
import {
  commentIdsNamedIn,
  commentRecordsOf,
  createCommentRecordsPlugin,
} from "./commentRecords";
import { fromProseDoc } from "./conversion/fromProseDoc";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const note = (text: string) => [
  { type: "paragraph" as const, content: [run(text)] },
];

const thread: Comment[] = [
  {
    id: 1,
    author: "Reviewer",
    date: "2024-01-01T00:00:00Z",
    anchorKind: "range",
    content: note("Check this."),
  },
  {
    id: 2,
    parentId: 1,
    author: "Author",
    date: "2024-01-02T00:00:00Z",
    content: note("Done."),
  },
];

/** "Hello" under comment `id`, then " world". */
const documentFor = (id: number, comments: Comment[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            { type: "commentRangeStart", id },
            run("Hello"),
            { type: "commentRangeEnd", id },
            { type: "commentReference", id },
            run(" world"),
          ],
        },
      ],
      comments,
    },
  },
});

const open = (document: Document, mode: EditorMode) =>
  new HeadlessEditorView(
    createHarnessState(document, mode, [
      createCommentRecordsPlugin(document.package.document.comments ?? []),
    ]),
  );

/** The slice of `state` covering `text` in its first matching paragraph. */
const sliceOf = (state: EditorState, text: string): Slice => {
  const block = findTextblock(state.doc, text);
  if (block === null) throw new Error(`no paragraph holds "${text}"`);
  const from = block.pos + 1 + block.node.textContent.indexOf(text);
  return state.doc.slice(from, from + text.length);
};

/** Paste `slice` at the end of the paragraph reading `text`. */
const pasteAtEnd = (view: HeadlessEditorView, text: string, slice: Slice) => {
  const block = findTextblock(view.state.doc, text);
  if (block === null) throw new Error(`no paragraph holds "${text}"`);
  const end = block.pos + block.node.nodeSize - 1;
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, end)),
  );
  view.paste(slice);
};

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

/** How many range starts and ends each comment saves with. */
const savedRanges = (state: EditorState, base: Document) => {
  const saved = fromProseDoc(state.doc, base, {
    stylesheetSource: { type: "package" },
  });
  const counts = new Map<number, { starts: number; ends: number }>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const type: unknown = Reflect.get(value, "type");
    const id: unknown = Reflect.get(value, "id");
    if (
      (type === "commentRangeStart" || type === "commentRangeEnd") &&
      typeof id === "number"
    ) {
      const entry = counts.get(id) ?? { starts: 0, ends: 0 };
      if (type === "commentRangeStart") entry.starts += 1;
      else entry.ends += 1;
      counts.set(id, entry);
      return;
    }
    for (const key of ["content", "children"]) visit(Reflect.get(value, key));
  };
  visit(saved.package.document.content);
  return counts;
};

describe("pasting commented text", () => {
  for (const mode of EDITOR_MODES) {
    test(`${mode}: a copy of a standing comment takes a fresh id and a copy of its thread`, () => {
      const document = documentFor(1, structuredClone(thread));
      const view = open(document, mode);
      const before = view.state;
      pasteAtEnd(view, "Hello world", sliceOf(view.state, "Hello"));
      const records = commentRecordsOf(view.state) ?? [];
      expect(records.map(({ id, parentId }) => [id, parentId])).toEqual([
        [1, undefined],
        [2, 1],
        [3, undefined],
        [4, 3],
      ]);
      expect(records[2]?.content).toEqual(note("Check this."));
      expect(records[3]?.content).toEqual(note("Done."));
      expect(commentedText(view.state)).toEqual(
        new Map([
          [1, "Hello"],
          [3, "Hello"],
        ]),
      );
      expect(savedRanges(view.state, document)).toEqual(
        new Map([
          [1, { starts: 1, ends: 1 }],
          [3, { starts: 1, ends: 1 }],
        ]),
      );
      // Undo takes the copy and its thread out together.
      view.execute(undo);
      expect(view.state.doc.eq(before.doc)).toBe(true);
      expect(commentRecordsOf(view.state)).toEqual(thread);
    });

    test(`${mode}: without a thread for the comment, the paste leaves it out`, () => {
      // The document names comment 1, but the session holds no thread for it.
      const document = documentFor(1, []);
      const view = open(document, mode);
      pasteAtEnd(view, "Hello world", sliceOf(view.state, "Hello"));
      expect(commentRecordsOf(view.state)).toEqual([]);
      expect(commentedText(view.state)).toEqual(new Map([[1, "Hello"]]));
      expect([...commentIdsNamedIn(view.state.doc)]).toEqual([1]);
    });

    test(`${mode}: a comment the document does not name goes in as it is`, () => {
      const source = open(documentFor(7, []), mode);
      const slice: Slice = sliceOf(source.state, "Hello");
      const document = documentFor(1, structuredClone(thread));
      const view = open(document, mode);
      pasteAtEnd(view, "Hello world", slice);
      expect(commentRecordsOf(view.state)).toEqual(thread);
      expect(commentedText(view.state)).toEqual(
        new Map([
          [1, "Hello"],
          [7, "Hello"],
        ]),
      );
    });
  }
});
