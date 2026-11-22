/**
 * One range per comment id: a pasted slice repeating the anchors of a
 * comment that stands in the document gets a fresh comment id and a copy of
 * the thread when it carries the comment, and leaves those anchors out when
 * it does not. Anchors of a comment the document does not hold go in as they
 * are.
 */
import { describe, expect, test } from "bun:test";

import type {
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { commentDocumentIssue } from "../comments";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type InsertContentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionStamp,
} from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const start = (id: number) => ({ type: "commentRangeStart", id }) as const;
const end = (id: number) => ({ type: "commentRangeEnd", id }) as const;
const reference = (id: number) => ({ type: "commentReference", id }) as const;
const body = (paraId: string, text: string): Paragraph[] => [
  { type: "paragraph", paraId, content: [run(text)] },
];

const root: Comment = {
  id: 1,
  author: "A",
  anchorKind: "range",
  content: body("00000010", "Check this."),
};
const reply: Comment = {
  id: 2,
  parentId: 1,
  author: "B",
  content: body("00000011", "Done."),
};

/** "ab" under comment 1 (with a reply), its reference, then "cd": offsets 0..2, 2..3 and 3..5. */
const documentOf = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [start(1), run("ab"), end(1), reference(1), run("cd")],
        },
      ],
      comments: [structuredClone(root), structuredClone(reply)],
    },
  },
});

/** The thread as a paste carries it: the same comments, new paragraph ids. */
const carried: readonly Comment[] = [
  { ...structuredClone(root), content: body("00000020", "Check this.") },
  { ...structuredClone(reply), content: body("00000021", "Done.") },
];

const MODES = {
  direct: undefined,
  tracked: { id: 90, author: "R", date: "2026-01-01T00:00:00Z" },
  suggested: {
    id: 90,
    author: "R",
    date: "2026-01-01T00:00:00Z",
    suggestionId: "s1",
  },
} as const satisfies Record<string, RevisionStamp | undefined>;
const MODE_NAMES = Object.keys(MODES) as (keyof typeof MODES)[];

/**
 * A tracked change holds no bare comment anchor, so the tracked paste copies
 * a hyperlink holding the range, as a link carrying a comment pastes.
 */
const sliceFor = (
  mode: keyof typeof MODES,
  content: ParagraphContent[],
): ParagraphContent[] =>
  mode === "direct"
    ? content
    : [{ type: "hyperlink", anchor: "here", children: content }];

const paste = (
  mode: keyof typeof MODES,
  content: ParagraphContent[],
  extra: Partial<InsertContentOp> = {},
): InsertContentOp => {
  const revision = MODES[mode];
  return {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
    slice: { content: sliceFor(mode, content), openStart: 0, openEnd: 0 },
    ...(revision === undefined
      ? {}
      : { revision, newIds: { revision: [91, 92, 93] } }),
    ...extra,
  };
};

/** Each comment id's anchors in document order. */
const anchorsById = (document: Document): Map<number, string[]> => {
  const out = new Map<number, string[]>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const type: unknown = Reflect.get(value, "type");
    const id: unknown = Reflect.get(value, "id");
    if (
      (type === "commentRangeStart" ||
        type === "commentRangeEnd" ||
        type === "commentReference") &&
      typeof id === "number"
    ) {
      out.set(id, [...(out.get(id) ?? []), type]);
      return;
    }
    for (const key of ["content", "children"]) visit(Reflect.get(value, key));
  };
  visit(document.package.document.content);
  return out;
};

const textOf = (document: Document): string => {
  let text = "";
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    if (Reflect.get(value, "type") === "text")
      text += String(Reflect.get(value, "text"));
    for (const key of ["content", "children"]) visit(Reflect.get(value, key));
  };
  visit(document.package.document.content);
  return text;
};

const RANGE = ["commentRangeStart", "commentRangeEnd", "commentReference"];

describe("a pasted copy of a standing comment's range", () => {
  for (const mode of MODE_NAMES) {
    test(`${mode}: takes a fresh id and a copy of the thread`, () => {
      const before = documentOf();
      const op = paste(mode, [start(1), run("ab"), end(1), reference(1)], {
        comments: carried,
        newCommentIds: [5, 6],
      });
      const applied = applyDocumentOp(before, op);
      if (applied.isErr()) throw applied.error;
      const after = applied.value.document;
      expect(textOf(after)).toBe("abcdab");
      const anchors = anchorsById(after);
      expect(anchors.get(1)).toEqual(RANGE);
      expect(anchors.get(5)).toEqual(RANGE);
      expect(commentDocumentIssue(after)).toBeUndefined();
      const comments = after.package.document.comments ?? [];
      expect(comments.map(({ id, parentId }) => [id, parentId])).toEqual([
        [1, undefined],
        [2, 1],
        [5, undefined],
        [6, 5],
      ]);
      const copies = comments.slice(2);
      expect(copies.map(({ suggestionId }) => suggestionId)).toEqual(
        mode === "suggested" ? ["s1", "s1"] : [undefined, undefined],
      );
      expect(copies.map(({ content }) => content[0]?.paraId)).toEqual([
        "00000020",
        "00000021",
      ]);
      // Undo takes the copy and its thread out again.
      const undone = applyDocumentOps(after, applied.value.inverse);
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toEqual(before);
      if (mode === "direct") return;
      // Rejecting the paste retires the copied thread with its anchors.
      const rejected = applyDocumentOp(after, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: applied.value.revisions,
        decision: REVISION_DECISIONS.REJECT,
      });
      if (rejected.isErr()) throw rejected.error;
      expect(textOf(rejected.value.document)).toBe("abcd");
      expect(anchorsById(rejected.value.document).has(5)).toBe(false);
      expect(
        (rejected.value.document.package.document.comments ?? []).map(
          ({ id }) => id,
        ),
      ).toEqual([1, 2]);
    });

    test(`${mode}: without the comment, leaves the repeated anchors out`, () => {
      const before = documentOf();
      const op = paste(mode, [start(1), run("ab"), end(1), reference(1)]);
      const applied = applyDocumentOp(before, op);
      if (applied.isErr()) throw applied.error;
      const after = applied.value.document;
      expect(textOf(after)).toBe("abcdab");
      expect(anchorsById(after)).toEqual(new Map([[1, RANGE]]));
      expect(commentDocumentIssue(after)).toBeUndefined();
      expect(after.package.document.comments).toEqual([root, reply]);
      const undone = applyDocumentOps(after, applied.value.inverse);
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toEqual(before);
    });

    test(`${mode}: a new comment id goes in as it is`, () => {
      const before = documentOf();
      const op = paste(mode, [start(7), run("ab"), end(7)]);
      const applied = applyDocumentOp(before, op);
      if (applied.isErr()) throw applied.error;
      const after = applied.value.document;
      const anchors = anchorsById(after);
      expect(anchors.get(1)).toEqual(RANGE);
      expect(anchors.get(7)).toEqual(["commentRangeStart", "commentRangeEnd"]);
      expect(after.package.document.comments).toEqual([root, reply]);
    });
  }

  test("a slice holding only a repeated anchor inserts nothing", () => {
    const before = documentOf();
    const applied = applyDocumentOp(before, paste("direct", [start(1)]));
    if (applied.isErr()) throw applied.error;
    expect(applied.value.document).toBe(before);
    expect(applied.value.inverse).toEqual([]);
  });

  test("undoing a deletion puts back an anchor its comment has standing elsewhere", () => {
    const before = documentOf();
    before.package.document.content.push({
      type: "paragraph",
      paraId: "00000002",
      content: [start(1), run("ef"), end(1), reference(1)],
    });
    const original = structuredClone(before);
    const deleted = applyDocumentOp(before, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: {
        story: OP_STORIES.MAIN,
        blockId: "00000002",
        offset: 0,
        zeroWidthBefore: 0,
      },
      to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 3 },
    });
    if (deleted.isErr()) throw deleted.error;
    expect(anchorsById(deleted.value.document)).toEqual(new Map([[1, RANGE]]));
    const undone = applyDocumentOps(
      deleted.value.document,
      JSON.parse(JSON.stringify(deleted.value.inverse)),
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(original);
  });

  test("a part of a standing comment's range is left out even when carried", () => {
    const applied = applyDocumentOp(
      documentOf(),
      paste("direct", [start(1), run("ab")], {
        comments: carried,
        newCommentIds: [5, 6],
      }),
    );
    if (applied.isErr()) throw applied.error;
    expect(anchorsById(applied.value.document)).toEqual(new Map([[1, RANGE]]));
    expect(applied.value.document.package.document.comments).toEqual([
      root,
      reply,
    ]);
  });

  test("an anchor completing a standing half range goes in as it is", () => {
    const before: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [start(1), run("abcd")],
            },
          ],
          comments: [structuredClone(root)],
        },
      },
    };
    const applied = applyDocumentOp(
      before,
      paste("direct", [end(1)], {
        at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 4 },
      }),
    );
    if (applied.isErr()) throw applied.error;
    expect(anchorsById(applied.value.document).get(1)).toEqual([
      "commentRangeStart",
      "commentRangeEnd",
    ]);
  });

  test("a carried thread of a new comment joins the records", () => {
    const thread: Comment[] = [
      { ...structuredClone(root), id: 7, content: body("00000030", "New.") },
    ];
    const before = documentOf();
    const applied = applyDocumentOp(
      before,
      paste("direct", [start(7), run("x"), end(7)], { comments: thread }),
    );
    if (applied.isErr()) throw applied.error;
    const after = applied.value.document;
    expect(after.package.document.comments?.map(({ id }) => id)).toEqual([
      1, 2, 7,
    ]);
    expect(commentDocumentIssue(after)).toBeUndefined();
    const undone = applyDocumentOps(after, applied.value.inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(before);
  });

  test("a pasted copy of a point comment undoes exactly", () => {
    const before: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [run("ab"), reference(1), run("cd")],
            },
          ],
          comments: [{ ...structuredClone(root), anchorKind: "point" }],
        },
      },
    };
    const applied = applyDocumentOp(
      before,
      paste("direct", [reference(1)], {
        at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
        comments: [
          {
            ...root,
            anchorKind: "point",
            content: body("00000020", "Check this."),
          },
        ],
        newCommentIds: [5],
      }),
    );
    if (applied.isErr()) throw applied.error;
    const after = applied.value.document;
    expect(anchorsById(after).get(5)).toEqual(["commentReference"]);
    expect(commentDocumentIssue(after)).toBeUndefined();
    const undone = applyDocumentOps(after, applied.value.inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(before);
  });

  test("a copy without fresh comment ids is refused", () => {
    const refused = applyDocumentOp(
      documentOf(),
      paste("direct", [start(1), run("ab"), end(1)], {
        comments: carried,
        newCommentIds: [5],
      }),
    );
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
    );
  });

  test("a fresh comment id the package uses is refused", () => {
    const refused = applyDocumentOp(
      documentOf(),
      paste("direct", [start(1), run("ab"), end(1)], {
        comments: carried,
        newCommentIds: [2, 6],
      }),
    );
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
    );
  });

  test("a carried comment reusing a package paragraph id is refused", () => {
    const refused = applyDocumentOp(
      documentOf(),
      paste("direct", [start(1), run("ab"), end(1)], {
        comments: [root, { ...reply, content: body("00000021", "Done.") }],
        newCommentIds: [5, 6],
      }),
    );
    expect(refused.isErr() && refused.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
    );
  });
});
