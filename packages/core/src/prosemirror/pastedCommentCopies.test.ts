/**
 * One range per comment id, pasting in a ProseMirror session: a pasted copy
 * of commented text takes a fresh comment id and a copy of the thread, or
 * leaves the comment out when the session holds no thread for it; a paste of
 * a comment the document does not name goes in as it is. Each case runs in
 * editing and in suggesting, and saves one range per comment.
 */
import { describe, expect, test } from "bun:test";
import { redo, undo } from "prosemirror-history";
import { Fragment, Slice } from "prosemirror-model";
import {
  type EditorState,
  NodeSelection,
  TextSelection,
} from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import {
  createHarnessState,
  EDITOR_MODES,
  type EditorMode,
  findTextblock,
  HeadlessEditorView,
} from "../__tests__/editorHarness";
import type { Comment, ParagraphContent } from "../types/content";
import type { Document, TableCell } from "../types/document";
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
    for (const key of ["content", "children", "rows", "cells"])
      visit(Reflect.get(value, key));
  };
  visit(saved.package.document.content);
  return counts;
};

/** Select `length` characters (all of `text`) where `text` stands. */
const select = (
  view: HeadlessEditorView,
  text: string,
  length = text.length,
) => {
  const block = findTextblock(view.state.doc, text);
  if (block === null) throw new Error(`no paragraph holds "${text}"`);
  // A reference takes a position that the paragraph's text does not.
  const offset = block.node.textContent.indexOf(text);
  let from = block.pos + 1 + offset;
  block.node.forEach((node, position) => {
    if (node.isAtom && !node.isText && position < offset) from += node.nodeSize;
  });
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.create(view.state.doc, from, from + length),
    ),
  );
};

/** Drag `text` to the end of the document's last paragraph. */
const dragToEnd = (
  view: HeadlessEditorView,
  text: string,
  { move }: { move: boolean },
) => {
  select(view, text);
  view.drag(view.state.doc.content.size - 1, { move });
};

/** "Hello", the reference of point comment 1, then " world". */
const pointDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            run("Hello"),
            { type: "commentReference", id: 1 },
            run(" world"),
          ],
        },
      ],
      comments: [
        {
          id: 1,
          author: "Reviewer",
          anchorKind: "point",
          content: note("Here."),
        },
      ],
    },
  },
});
/** "lo", the reference and " wo" in {@link pointDocument}. */
const POINT_RANGE = [4, 10] as const;

/** Comment 1 covers `original` and a copy of it, with its thread, `copied`. */
const expectCopied = (
  state: EditorState,
  base: Document,
  { original = "Hello", copied = "Hello" } = {},
) => {
  const records = commentRecordsOf(state) ?? [];
  const copy = records[2]?.id ?? Number.NaN;
  const copiedReply = records[3]?.id ?? Number.NaN;
  expect(new Set([1, 2, copy, copiedReply]).size).toBe(4);
  expect(records.map(({ id, parentId }) => [id, parentId])).toEqual([
    [1, undefined],
    [2, 1],
    [copy, undefined],
    [copiedReply, copy],
  ]);
  expect(commentedText(state)).toEqual(
    new Map([
      [1, original],
      [copy, copied],
    ]),
  );
  expect(savedRanges(state, base)).toEqual(
    new Map([
      [1, { starts: 1, ends: 1 }],
      [copy, { starts: 1, ends: 1 }],
    ]),
  );
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

    test(`${mode}: plain text pasted inside a comment joins its range`, () => {
      const document = documentFor(1, structuredClone(thread));
      const view = open(document, mode);
      select(view, "lo", 0);
      view.pasteText("XX");
      expect(commentedText(view.state)).toEqual(new Map([[1, "HelXXlo"]]));
      expect(commentRecordsOf(view.state)).toEqual(thread);
      expect(savedRanges(view.state, document)).toEqual(
        new Map([[1, { starts: 1, ends: 1 }]]),
      );
    });

    test(`${mode}: a slice holding only a range's reference leaves it out`, () => {
      const view = open(documentFor(1, structuredClone(thread)), mode);
      const { doc } = view.state;
      let at: number | undefined;
      doc.descendants((node, pos) => {
        if (node.type.name === "commentReference") at = pos;
      });
      if (at === undefined) throw new Error("no reference");
      pasteAtEnd(view, "Hello world", doc.slice(at, at + 1));
      expect(commentRecordsOf(view.state)).toEqual(thread);
      expect([...commentIdsNamedIn(view.state.doc)]).toEqual([1]);
    });

    test(`${mode}: a pasted copy of a point comment leaves it out`, () => {
      const document = pointDocument();
      const view = open(document, mode);
      const { doc } = view.state;
      pasteAtEnd(view, "Hello world", doc.slice(...POINT_RANGE));
      expect(view.state.doc.textContent).toBe("Hello worldlo wo");
      expect(commentRecordsOf(view.state)).toEqual(
        document.package.document.comments ?? [],
      );
      expect([...commentIdsNamedIn(view.state.doc)]).toEqual([1]);
    });

    test(`${mode}: a copy of a proposed comment belongs to no suggestion`, () => {
      const proposed = thread.map((comment) => ({
        ...comment,
        suggestionId: "source",
      }));
      const view = open(documentFor(1, proposed), mode);
      pasteAtEnd(view, "Hello world", sliceOf(view.state, "Hello"));
      const copies = commentRecordsOf(view.state)?.slice(2) ?? [];
      expect(copies.map(({ id }) => id)).toEqual([3, 4]);
      expect(copies.map((copy) => Object.hasOwn(copy, "suggestionId"))).toEqual(
        [false, false],
      );
      const owners = new Map<unknown, unknown>();
      view.state.doc.descendants((node) => {
        for (const { type, attrs } of node.marks)
          if (type.name === "comment")
            owners.set(attrs["commentId"], attrs["suggestionId"]);
      });
      expect(owners).toEqual(
        new Map<unknown, unknown>([
          [1, "source"],
          [3, null],
        ]),
      );
    });
  }
});

/** {@link documentFor}, then a 2×2 table reading "a" to "d". */
const tableDocument = (comments: Comment[]): Document => {
  const document = documentFor(1, comments);
  const cell = (text: string): TableCell => ({
    type: "tableCell",
    content: [{ type: "paragraph", content: [run(text)] }],
  });
  document.package.document.content.push(
    {
      type: "table",
      columnWidths: [1000, 1000],
      rows: [
        { type: "tableRow", cells: [cell("a"), cell("b")] },
        { type: "tableRow", cells: [cell("c"), cell("d")] },
      ],
    },
    { type: "paragraph", content: [run("After")] },
  );
  return document;
};

/** The positions of the table's cells. */
const cellPositions = (state: EditorState): number[] => {
  const positions: number[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name === "tableCell") positions.push(pos);
  });
  return positions;
};

/** Select the table's cells from the `first` to the `last`. */
const selectCells = (view: HeadlessEditorView, first: number, last: number) => {
  const cells = cellPositions(view.state);
  view.dispatch(
    view.state.tr.setSelection(
      CellSelection.create(
        view.state.doc,
        cells[first] ?? Number.NaN,
        cells[last] ?? Number.NaN,
      ),
    ),
  );
};

/** The saved comment ids with other than one range start and one end. */
const brokenRanges = (state: EditorState, base: Document) =>
  [...savedRanges(state, base)].filter(
    ([, { starts, ends }]) => starts !== 1 || ends !== 1,
  );

/** The ids of the records' threads, each of a comment and its reply. */
const threadRoots = (records: readonly Comment[]): number[] => {
  const roots = records.flatMap(({ id, parentId }) =>
    parentId === undefined ? [id] : [],
  );
  expect(records.map(({ parentId }) => parentId)).toEqual(
    roots.flatMap((id) => [undefined, id]),
  );
  expect(new Set(records.map(({ id }) => id)).size).toBe(records.length);
  return roots;
};

describe("pasting commented text into several cells", () => {
  for (const mode of EDITOR_MODES) {
    test(`${mode}: each cell's copy takes its own id and copy of the thread`, () => {
      const document = tableDocument(structuredClone(thread));
      const view = open(document, mode);
      const before = view.state;
      const slice = sliceOf(view.state, "Hello");
      selectCells(view, 0, 3);
      view.paste(slice);
      const after = view.state;
      const records = commentRecordsOf(after) ?? [];
      const roots = threadRoots(records);
      expect(roots).toHaveLength(5);
      expect(commentedText(after)).toEqual(
        new Map(roots.map((id) => [id, "Hello"])),
      );
      expect(savedRanges(after, document).size).toBe(5);
      expect(brokenRanges(after, document)).toEqual([]);
      view.execute(undo);
      expect(view.state.doc.eq(before.doc)).toBe(true);
      expect(commentRecordsOf(view.state)).toEqual(thread);
      view.execute(redo);
      expect(view.state.doc.eq(after.doc)).toBe(true);
      expect(commentRecordsOf(view.state)).toEqual(records);
    });

    test(`${mode}: a comment the document does not name stands once`, () => {
      const source = open(documentFor(7, []), mode);
      const slice = sliceOf(source.state, "Hello");
      const document = tableDocument([
        ...structuredClone(thread),
        {
          id: 7,
          author: "Reviewer",
          anchorKind: "range",
          content: note("Cut"),
        },
      ]);
      const view = open(document, mode);
      selectCells(view, 0, 1);
      view.paste(slice);
      const records = commentRecordsOf(view.state) ?? [];
      const copy = records[3]?.id ?? Number.NaN;
      expect(records.map(({ id }) => id)).toEqual([1, 2, 7, copy]);
      expect(commentedText(view.state)).toEqual(
        new Map([
          [1, "Hello"],
          [7, "Hello"],
          [copy, "Hello"],
        ]),
      );
      expect(brokenRanges(view.state, document)).toEqual([]);
    });

    test(`${mode}: cells repeated to fill a selection take a copy each`, () => {
      const document = tableDocument(structuredClone(thread));
      const view = open(document, mode);
      // A block of one row: the comment runs from its first cell into its second.
      const { schema } = view.state;
      const commented = schema.marks["comment"]?.create({ commentId: 1 });
      const cell = (text: string) =>
        schema.node("tableCell", null, [
          schema.node("paragraph", null, [
            schema.text(text, commented ? [commented] : []),
          ]),
        ]);
      const block = schema.node("table", { columnWidths: [1000, 1000] }, [
        schema.node("tableRow", null, [cell("x"), cell("y")]),
      ]);
      const before = view.state;
      selectCells(view, 0, 3);
      view.paste(new Slice(Fragment.from(block), 0, 0));
      const roots = threadRoots(commentRecordsOf(view.state) ?? []);
      expect(roots).toHaveLength(3);
      expect(commentedText(view.state)).toEqual(
        new Map(roots.map((id) => [id, id === 1 ? "Hello" : "xy"])),
      );
      expect(brokenRanges(view.state, document)).toEqual([]);
      view.execute(undo);
      expect(view.state.doc.eq(before.doc)).toBe(true);
      expect(commentRecordsOf(view.state)).toEqual(thread);
    });

    test(`${mode}: without a thread, the copies after the first leave the comment out`, () => {
      const source = open(documentFor(7, []), mode);
      const slice = sliceOf(source.state, "Hello");
      const view = open(tableDocument(structuredClone(thread)), mode);
      selectCells(view, 0, 3);
      view.paste(slice);
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

describe("pasting commented text over its source", () => {
  test("editing: the comment stays on the pasted text", () => {
    const proposed = thread.map((comment) => ({
      ...comment,
      suggestionId: "source",
    }));
    const document = documentFor(1, structuredClone(proposed));
    const view = open(document, "editing");
    const before = view.state;
    const slice = sliceOf(view.state, "Hello");
    select(view, "Hello");
    view.paste(slice);
    const after = view.state;
    expect(after.doc.textContent).toBe("Hello world");
    expect(commentRecordsOf(after)).toEqual(proposed);
    expect(commentedText(after)).toEqual(new Map([[1, "Hello"]]));
    expect(brokenRanges(after, document)).toEqual([]);
    const owners = new Set<unknown>();
    after.doc.descendants((node) => {
      for (const { type, attrs } of node.marks)
        if (type.name === "comment") owners.add(attrs["suggestionId"]);
    });
    expect([...owners]).toEqual(["source"]);
    view.execute(undo);
    expect(view.state.doc.eq(before.doc)).toBe(true);
    expect(commentRecordsOf(view.state)).toEqual(proposed);
    view.execute(redo);
    expect(view.state.doc.eq(after.doc)).toBe(true);
    // The same clipboard pastes the comment again, as a copy.
    pasteAtEnd(view, "Hello world", slice);
    expect(commentedText(view.state)).toEqual(
      new Map([
        [1, "Hello"],
        [3, "Hello"],
      ]),
    );
    expect(commentRecordsOf(view.state)?.map(({ id }) => id)).toEqual([
      1, 2, 3, 4,
    ]);
  });

  test("suggesting: the pasted text takes a copy beside the struck source", () => {
    const document = documentFor(1, structuredClone(thread));
    const view = open(document, "suggesting");
    const before = view.state;
    const slice = sliceOf(view.state, "Hello");
    select(view, "Hello");
    view.paste(slice);
    expectCopied(view.state, document);
    view.execute(undo);
    expect(view.state.doc.eq(before.doc)).toBe(true);
    expect(commentRecordsOf(view.state)).toEqual(thread);
  });

  test("suggesting: pasting part of it over part of it leaves one range per comment", () => {
    const document = documentFor(1, structuredClone(thread));
    const view = open(document, "suggesting");
    const slice = sliceOf(view.state, "Hel");
    select(view, "Hello", 3);
    view.paste(slice);
    const records = commentRecordsOf(view.state) ?? [];
    expect(records.map(({ id }) => id)).toEqual([1, 2, 3, 4]);
    expect(commentedText(view.state)).toEqual(
      new Map([
        [1, "Hello"],
        [3, "Hel"],
      ]),
    );
    expect(brokenRanges(view.state, document)).toEqual([]);
  });

  test("suggesting: the pasted text takes a copy where the replaced node stays struck", () => {
    const document = documentFor(1, structuredClone(thread));
    const [paragraph] = document.package.document.content;
    if (paragraph?.type !== "paragraph") throw new Error("no paragraph");
    paragraph.content[1] = { type: "run", content: [{ type: "tab" }] };
    const view = open(document, "suggesting");
    const before = view.state;
    let tab = Number.NaN;
    view.state.doc.descendants((node, pos) => {
      if (node.type.name === "tab") tab = pos;
    });
    const selection = NodeSelection.create(view.state.doc, tab);
    const slice = selection.content();
    view.dispatch(view.state.tr.setSelection(selection));
    view.paste(slice);
    const roots = threadRoots(commentRecordsOf(view.state) ?? []);
    expect(roots).toHaveLength(2);
    expect([...commentIdsNamedIn(view.state.doc)].toSorted()).toEqual(
      roots.toSorted(),
    );
    expect(brokenRanges(view.state, document)).toEqual([]);
    view.execute(undo);
    expect(view.state.doc.eq(before.doc)).toBe(true);
    expect(commentRecordsOf(view.state)).toEqual(thread);
  });

  test("editing: a comment without a thread stays on the pasted text", () => {
    const document = documentFor(1, []);
    const view = open(document, "editing");
    const slice = sliceOf(view.state, "Hello");
    select(view, "Hello");
    view.paste(slice);
    expect(commentedText(view.state)).toEqual(new Map([[1, "Hello"]]));
    expect(commentRecordsOf(view.state)).toEqual([]);
  });

  test("editing: a point comment stays on the pasted text", () => {
    const document = pointDocument();
    const view = open(document, "editing");
    const before = view.state;
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, ...POINT_RANGE),
      ),
    );
    view.paste(view.state.doc.slice(...POINT_RANGE));
    expect(view.state.doc.textContent).toBe("Hello world");
    expect(view.state.doc.eq(before.doc)).toBe(true);
    expect([...commentIdsNamedIn(view.state.doc)]).toEqual([1]);
    expect(commentRecordsOf(view.state)).toEqual(
      document.package.document.comments ?? [],
    );
  });

  test("editing: into its own cell among several, one copy keeps the comment", () => {
    const document = tableDocument([]);
    const [commented] = documentFor(1, []).package.document.content;
    const table = document.package.document.content[1];
    if (commented?.type !== "paragraph" || table?.type !== "table")
      throw new Error("unexpected document");
    const [first] = table.rows[0]?.cells ?? [];
    if (first === undefined) throw new Error("no cell");
    first.content = [commented];
    document.package.document.content.shift();
    document.package.document.comments = structuredClone(thread);
    const view = open(document, "editing");
    const before = view.state;
    const slice = sliceOf(view.state, "Hello");
    selectCells(view, 0, 1);
    view.paste(slice);
    expect(commentRecordsOf(view.state)?.map(({ id }) => id)).toEqual([
      1, 2, 3, 4,
    ]);
    expect(commentedText(view.state)).toEqual(
      new Map([
        [1, "Hello"],
        [3, "Hello"],
      ]),
    );
    expect(brokenRanges(view.state, document)).toEqual([]);
    view.execute(undo);
    expect(view.state.doc.eq(before.doc)).toBe(true);
    expect(commentRecordsOf(view.state)).toEqual(thread);
  });
});

describe("dragging commented text", () => {
  for (const mode of EDITOR_MODES) {
    test(`${mode}: a dropped copy takes a fresh id and a copy of its thread`, () => {
      const document = documentFor(1, structuredClone(thread));
      const view = open(document, mode);
      const before = view.state;
      dragToEnd(view, "Hello", { move: false });
      expect(view.state.doc.textContent).toBe("Hello worldHello");
      expectCopied(view.state, document);
      view.execute(undo);
      expect(view.state.doc.eq(before.doc)).toBe(true);
      expect(commentRecordsOf(view.state)).toEqual(thread);
    });

    test(`${mode}: a dropped copy without a thread leaves the comment out`, () => {
      const view = open(documentFor(1, []), mode);
      dragToEnd(view, "Hello", { move: false });
      expect(commentRecordsOf(view.state)).toEqual([]);
      expect(commentedText(view.state)).toEqual(new Map([[1, "Hello"]]));
      expect([...commentIdsNamedIn(view.state.doc)]).toEqual([1]);
    });

    test(`${mode}: moving part of a commented text leaves one range per comment`, () => {
      const document = documentFor(1, structuredClone(thread));
      const view = open(document, mode);
      const before = view.state;
      dragToEnd(view, "Hel", { move: true });
      expectCopied(view.state, document, {
        original: mode === "editing" ? "lo" : "Hello",
        copied: "Hel",
      });
      view.execute(undo);
      expect(view.state.doc.eq(before.doc)).toBe(true);
      expect(commentRecordsOf(view.state)).toEqual(thread);
    });

    test(`${mode}: an edit after a drop is no part of it`, () => {
      const view = open(documentFor(1, structuredClone(thread)), mode);
      dragToEnd(view, " wor", { move: true });
      const { doc } = view.state;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(doc, doc.content.size - 1),
        ),
      );
      view.dispatch(view.state.tr.insertText("!"));
      expect(commentRecordsOf(view.state)).toEqual(thread);
      expect(commentedText(view.state)).toEqual(new Map([[1, "Hello"]]));
    });
  }

  test("a drag given up leaves nothing for the next paste", () => {
    const view = open(documentFor(1, structuredClone(thread)), "editing");
    select(view, "lo", 0);
    view.dragging = { slice: sliceOf(view.state, "Hello"), move: false };
    for (const plugin of view.state.plugins)
      plugin.props.transformPasted?.call(
        plugin,
        view.dragging.slice,
        view as never,
        false,
      );
    view.dragging = null;
    view.pasteText("XX");
    expect(commentRecordsOf(view.state)).toEqual(thread);
    expect(commentedText(view.state)).toEqual(new Map([[1, "HelXXlo"]]));
  });

  test("a move keeps the comment on the moved text", () => {
    const document = documentFor(1, structuredClone(thread));
    const view = open(document, "editing");
    const before = view.state;
    dragToEnd(view, "Hello", { move: true });
    expect(view.state.doc.textContent).toBe(" worldHello");
    expect(commentRecordsOf(view.state)).toEqual(thread);
    expect(commentedText(view.state)).toEqual(new Map([[1, "Hello"]]));
    expect(savedRanges(view.state, document)).toEqual(
      new Map([[1, { starts: 1, ends: 1 }]]),
    );
    view.execute(undo);
    expect(view.state.doc.eq(before.doc)).toBe(true);
    expect(commentRecordsOf(view.state)).toEqual(thread);
  });

  test("a move without a thread keeps the comment on the moved text", () => {
    const view = open(documentFor(1, []), "editing");
    dragToEnd(view, "Hello", { move: true });
    expect(commentedText(view.state)).toEqual(new Map([[1, "Hello"]]));
  });

  test("a move keeps a point comment it holds", () => {
    const document = pointDocument();
    const view = open(document, "editing");
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, ...POINT_RANGE),
      ),
    );
    view.drag(view.state.doc.content.size - 1, { move: true });
    expect(view.state.doc.textContent).toBe("Helrldlo wo");
    expect(commentRecordsOf(view.state)).toEqual(
      document.package.document.comments ?? [],
    );
    expect([...commentIdsNamedIn(view.state.doc)]).toEqual([1]);
  });

  test("a suggested move copies the comment beside the text it removes", () => {
    const document = documentFor(1, structuredClone(thread));
    const view = open(document, "suggesting");
    const before = view.state;
    dragToEnd(view, "Hello", { move: true });
    expect(view.state.doc.textContent).toBe("Hello worldHello");
    expectCopied(view.state, document);
    view.execute(undo);
    expect(view.state.doc.eq(before.doc)).toBe(true);
    expect(commentRecordsOf(view.state)).toEqual(thread);
  });
});
