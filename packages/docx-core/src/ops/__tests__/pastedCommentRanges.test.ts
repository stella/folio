/**
 * One range per comment id: a pasted slice repeating the anchors of a
 * comment that stands in the document gets a fresh comment id and a copy of
 * the thread when it carries the comment, and leaves those anchors out when
 * it does not. Anchors of a comment the document does not hold go in as they
 * are.
 */
import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../model/document";
import { COMMENT_PART_RELATIONSHIPS } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { commentDocumentIssue, sourceCommentAnchorsIn } from "../comments";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { BATCH_REJECTION_REASONS } from "../sequencing/envelope";
import { transformByFootprint } from "../sequencing/scope";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
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

  test("a slice holding only a standing range's reference leaves it out", () => {
    const before = documentOf();
    const applied = applyDocumentOp(
      before,
      paste("direct", [run("x"), reference(1)], {
        comments: carried,
        newCommentIds: [5, 6],
      }),
    );
    if (applied.isErr()) throw applied.error;
    const after = applied.value.document;
    expect(textOf(after)).toBe("abcdx");
    expect(anchorsById(after)).toEqual(new Map([[1, RANGE]]));
    expect(after.package.document.comments).toEqual([root, reply]);
    const undone = applyDocumentOps(after, applied.value.inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(before);
  });

  test("a range's reference standing alone is not copied as a point", () => {
    const before = documentOf();
    before.package.document.content = [
      {
        type: "paragraph",
        paraId: "00000001",
        content: [run("ab"), reference(1), run("cd")],
      },
    ];
    const applied = applyDocumentOp(
      before,
      paste("direct", [run("x"), reference(1)], {
        comments: carried,
        newCommentIds: [5, 6],
      }),
    );
    if (applied.isErr()) throw applied.error;
    const after = applied.value.document;
    expect(anchorsById(after)).toEqual(new Map([[1, ["commentReference"]]]));
    expect(after.package.document.comments).toEqual([root, reply]);
  });

  const TWICE = [
    start(7),
    run("x"),
    end(7),
    reference(7),
    run("y"),
    start(7),
    run("z"),
    end(7),
    reference(7),
  ];
  for (const { name, comments } of [
    { name: "uncarried", comments: undefined },
    {
      name: "carried",
      comments: [
        { ...structuredClone(root), id: 7, content: body("00000030", "New.") },
      ],
    },
  ])
    test(`${name}: a comment the document does not hold keeps its first range`, () => {
      const before = documentOf();
      const applied = applyDocumentOp(
        before,
        paste("direct", TWICE, comments === undefined ? {} : { comments }),
      );
      if (applied.isErr()) throw applied.error;
      const after = applied.value.document;
      expect(textOf(after)).toBe("abcdxyz");
      expect(anchorsById(after)).toEqual(
        new Map([
          [1, RANGE],
          [7, RANGE],
        ]),
      );
      const undone = applyDocumentOps(after, applied.value.inverse);
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toEqual(before);
    });

  test("a kept range end follows its start", () => {
    const applied = applyDocumentOp(
      documentOf(),
      paste("direct", [end(7), run("x"), start(7), run("y"), end(7)]),
    );
    if (applied.isErr()) throw applied.error;
    expect(anchorsById(applied.value.document).get(7)).toEqual([
      "commentRangeStart",
      "commentRangeEnd",
    ]);
  });

  test("undoing a deletion puts back every range its slice repeats", () => {
    const before: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [
                run("ab"),
                start(1),
                run("cd"),
                end(1),
                start(1),
                run("ef"),
                end(1),
                run("gh"),
              ],
            },
          ],
          comments: [structuredClone(root)],
        },
      },
    };
    const original = structuredClone(before);
    const deleted = applyDocumentOp(before, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 7 },
    });
    if (deleted.isErr()) throw deleted.error;
    expect(textOf(deleted.value.document)).toBe("ah");
    const undone = applyDocumentOps(
      deleted.value.document,
      JSON.parse(JSON.stringify(deleted.value.inverse)),
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(original);
  });

  test("copies take the fresh ids in carried order, threads interleaved", () => {
    const other: Comment = {
      ...structuredClone(root),
      id: 3,
      content: body("00000012", "Other."),
    };
    const anchored = [
      start(1),
      run("ab"),
      end(1),
      reference(1),
      start(3),
      run("cd"),
      end(3),
      reference(3),
    ];
    const before: Document = {
      package: {
        document: {
          content: [
            { type: "paragraph", paraId: "00000001", content: anchored },
          ],
          comments: [root, reply, other],
        },
      },
    };
    const [carriedRoot, carriedReply] = carried;
    if (carriedRoot === undefined || carriedReply === undefined)
      throw new Error("The thread carries a root and a reply.");
    const applied = applyDocumentOp(
      before,
      paste("direct", anchored, {
        comments: [
          carriedRoot,
          { ...other, content: body("00000022", "Other.") },
          carriedReply,
        ],
        newCommentIds: [5, 6, 7],
      }),
    );
    if (applied.isErr()) throw applied.error;
    const after = applied.value.document;
    expect(
      after.package.document.comments?.map(({ id, parentId, content }) => [
        id,
        parentId,
        content[0]?.paraId,
      ]),
    ).toEqual([
      [1, undefined, "00000010"],
      [2, 1, "00000011"],
      [3, undefined, "00000012"],
      [5, undefined, "00000020"],
      [6, undefined, "00000022"],
      [7, 5, "00000021"],
    ]);
    expect(anchorsById(after)).toEqual(
      new Map([
        [1, RANGE],
        [3, RANGE],
        [5, RANGE],
        [6, RANGE],
      ]),
    );
    const undone = applyDocumentOps(after, applied.value.inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(before);
  });

  for (const mode of MODE_NAMES)
    test(`${mode}: a copy of a proposed comment belongs to the insertion's own suggestion`, () => {
      const before = documentOf();
      before.package.document.comments = [
        { ...root, suggestionId: "source" },
        { ...reply, suggestionId: "source" },
      ];
      const applied = applyDocumentOp(
        before,
        paste(mode, [start(1), run("ab"), end(1), reference(1)], {
          comments: carried.map((comment) => ({
            ...comment,
            suggestionId: "source",
          })),
          newCommentIds: [5, 6],
        }),
      );
      if (applied.isErr()) throw applied.error;
      const copies =
        applied.value.document.package.document.comments?.slice(2) ?? [];
      const owner = mode === "suggested" ? "s1" : undefined;
      expect(copies.map(({ suggestionId }) => suggestionId)).toEqual([
        owner,
        owner,
      ]);
      expect(copies.map((copy) => Object.hasOwn(copy, "suggestionId"))).toEqual(
        [owner !== undefined, owner !== undefined],
      );
    });

  test("a proposed comment joining the records belongs to the insertion's own suggestion", () => {
    const applied = applyDocumentOp(
      documentOf(),
      paste("direct", [start(7), run("x"), end(7)], {
        comments: [
          {
            ...structuredClone(root),
            id: 7,
            suggestionId: "source",
            content: body("00000030", "New."),
          },
        ],
      }),
    );
    if (applied.isErr()) throw applied.error;
    const joined = applied.value.document.package.document.comments?.at(-1);
    expect(joined?.id).toBe(7);
    expect(joined !== undefined && Object.hasOwn(joined, "suggestionId")).toBe(
      false,
    );
  });

  /** The package's comment part relationship targets. */
  const commentParts = (document: Document): string[] =>
    [...(document.package.relationships?.values() ?? [])].flatMap(
      ({ target }) => (target.startsWith("comments") ? [target] : []),
    );
  const joining: readonly Comment[] = [
    { ...structuredClone(root), id: 7, content: body("00000030", "New.") },
    {
      ...structuredClone(reply),
      id: 8,
      parentId: 7,
      content: body("00000031", "Yes."),
    },
  ];
  const created = {
    author: "C",
    content: body("00000040", "Later."),
  } satisfies Omit<Comment, "id" | "parentId">;

  /**
   * Insert the thread 7 into `before`, apply `later`, and undo the insertion
   * alone; the document must keep every comment part its comments need.
   */
  const undoneBeside = (before: Document, later: DocumentOp): Document => {
    const pasted = applyDocumentOp(
      before,
      paste("direct", [start(7), run("x"), end(7)], { comments: joining }),
    );
    if (pasted.isErr()) throw pasted.error;
    expect(commentParts(pasted.value.document)).toEqual([
      "comments.xml",
      "commentsExtended.xml",
    ]);
    const next = applyDocumentOp(pasted.value.document, later);
    if (next.isErr()) throw next.error;
    const undone = applyDocumentOps(
      next.value.document,
      JSON.parse(JSON.stringify(pasted.value.inverse)),
    );
    if (undone.isErr()) throw undone.error;
    expect(anchorsById(undone.value.document).has(7)).toBe(false);
    return undone.value.document;
  };

  test("undoing an insertion keeps the part a comment made since needs", () => {
    const before: Document = {
      package: {
        document: {
          content: [
            { type: "paragraph", paraId: "00000001", content: [run("abcde")] },
            { type: "paragraph", paraId: "00000002", content: [run("fgh")] },
          ],
          comments: [],
        },
        relationships: new Map(),
      },
    };
    const undone = undoneBeside(before, {
      type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      comment: { ...created, id: 9 },
      anchor: {
        kind: "point",
        at: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 1 },
      },
    });
    expect(undone.package.document.comments?.map(({ id }) => id)).toEqual([9]);
    expect(commentParts(undone)).toEqual(["comments.xml"]);
  });

  test("undoing an insertion keeps the part a reply made since needs", () => {
    const before: Document = {
      package: {
        document: {
          content: [
            { type: "paragraph", paraId: "00000001", content: [run("abcde")] },
            {
              type: "paragraph",
              paraId: "00000002",
              content: [start(1), run("fgh"), end(1), reference(1)],
            },
          ],
          comments: [structuredClone(root)],
        },
        relationships: new Map([
          ["rId1", { id: "rId1", ...COMMENT_PART_RELATIONSHIPS.comments }],
        ]),
      },
    };
    const undone = undoneBeside(before, {
      type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      comment: { ...created, id: 9 },
      anchor: { kind: "reply", parentId: 1 },
    });
    expect(undone.package.document.comments?.map(({ id }) => id)).toEqual([
      1, 9,
    ]);
    expect(commentParts(undone)).toEqual([
      "comments.xml",
      "commentsExtended.xml",
    ]);
  });

  /** A run drawing a text box that holds `blocks`. */
  const box = (blocks: BlockContent[]): Run => ({
    type: "run",
    content: [
      {
        type: "shape",
        shape: {
          type: "shape",
          shapeType: "textBox",
          id: "box",
          size: { width: 914_400, height: 914_400 },
          textBody: { content: blocks },
        },
      },
    ],
  });
  const positioned = (marker: string, index: number) => ({
    children: [{ index, xml: `<w:${marker} w:id="1"/>` }],
  });
  /** Comment 1's range in a pasted text box, written each way a box can hold it. */
  const boxed = {
    "in a paragraph": [
      {
        type: "paragraph",
        paraId: "00000030",
        content: [start(1), run("ab"), end(1), reference(1)],
      },
    ],
    "between blocks": [
      start(1),
      { type: "paragraph", paraId: "00000030", content: [run("ab")] },
      end(1),
      { type: "paragraph", paraId: "00000031", content: [reference(1)] },
    ],
    "between a table's rows and cells": [
      {
        type: "table",
        preserved: positioned("commentRangeEnd", 1),
        rows: [
          {
            type: "tableRow",
            preserved: positioned("commentRangeStart", 0),
            cells: [
              {
                type: "tableCell",
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000030",
                    content: [run("ab")],
                  },
                ],
              },
            ],
          },
        ],
      },
      { type: "paragraph", paraId: "00000031", content: [reference(1)] },
    ],
  } as const satisfies Record<string, readonly unknown[]>;
  const sourceAnchors = (document: Document): Map<number, string[]> => {
    const out = new Map<number, string[]>();
    for (const { id, type } of sourceCommentAnchorsIn(document))
      out.set(id, [...(out.get(id) ?? []), type].toSorted());
    return out;
  };
  const WHOLE = RANGE.toSorted();

  for (const [place, blocks] of Object.entries(boxed)) {
    for (const mode of MODE_NAMES) {
      test(`${mode}: a range ${place} of a text box takes a fresh id and a copy of the thread`, () => {
        const before = documentOf();
        const applied = applyDocumentOp(
          before,
          paste(mode, [box(structuredClone(blocks) as BlockContent[])], {
            comments: carried,
            newCommentIds: [5, 6],
          }),
        );
        if (applied.isErr()) throw applied.error;
        const after = applied.value.document;
        expect(sourceAnchors(after)).toEqual(
          new Map([
            [1, WHOLE],
            [5, WHOLE],
          ]),
        );
        expect(commentDocumentIssue(after)).toBeUndefined();
        expect(
          after.package.document.comments?.map(({ id, parentId }) => [
            id,
            parentId,
          ]),
        ).toEqual([
          [1, undefined],
          [2, 1],
          [5, undefined],
          [6, 5],
        ]);
        const undone = applyDocumentOps(
          after,
          JSON.parse(JSON.stringify(applied.value.inverse)),
        );
        if (undone.isErr()) throw undone.error;
        expect(undone.value.document).toEqual(before);
      });

      test(`${mode}: a range ${place} of a text box is left out without a thread`, () => {
        const before = documentOf();
        const applied = applyDocumentOp(
          before,
          paste(mode, [box(structuredClone(blocks) as BlockContent[])]),
        );
        if (applied.isErr()) throw applied.error;
        const after = applied.value.document;
        expect(sourceAnchors(after)).toEqual(new Map([[1, WHOLE]]));
        expect(JSON.stringify(after.package.document.content)).toContain(
          '"textBody"',
        );
        expect(after.package.document.comments).toEqual([root, reply]);
        const undone = applyDocumentOps(
          after,
          JSON.parse(JSON.stringify(applied.value.inverse)),
        );
        if (undone.isErr()) throw undone.error;
        expect(undone.value.document).toEqual(before);
      });
    }
  }

  test("concurrent insertions naming one comment between a text box's rows conflict", () => {
    const [table] = boxed["between a table's rows and cells"];
    const op = paste("direct", [
      box([structuredClone(table) as unknown as BlockContent]),
    ]);
    const decided = transformByFootprint({
      op,
      over: structuredClone(op),
      effect: undefined,
      order: "before",
    });
    expect(decided?.isErr() && decided.error.reason).toBe(
      BATCH_REJECTION_REASONS.PART_CONFLICT,
    );
  });

  test("a start and an end in different text boxes are no range to copy", () => {
    const applied = applyDocumentOp(
      documentOf(),
      paste(
        "direct",
        [
          box([
            {
              type: "paragraph",
              paraId: "00000030",
              content: [start(1), run("x")],
            },
          ]),
          box([
            {
              type: "paragraph",
              paraId: "00000031",
              content: [run("y"), end(1), reference(1)],
            },
          ]),
        ],
        { comments: carried, newCommentIds: [5, 6] },
      ),
    );
    if (applied.isErr()) throw applied.error;
    expect(sourceAnchors(applied.value.document)).toEqual(
      new Map([[1, WHOLE]]),
    );
    expect(applied.value.document.package.document.comments).toEqual([
      root,
      reply,
    ]);
  });

  test("a slice repeating a missing comment's range in a text box keeps the first", () => {
    const applied = applyDocumentOp(
      documentOf(),
      paste("direct", [
        start(7),
        run("x"),
        box([
          {
            type: "paragraph",
            paraId: "00000030",
            content: [start(7), run("y"), end(7)],
          },
        ]),
        end(7),
        reference(7),
      ]),
    );
    if (applied.isErr()) throw applied.error;
    const { document } = applied.value;
    expect(sourceAnchors(document).get(7)).toEqual(WHOLE);
    expect(anchorsById(document).get(7)).toEqual(RANGE);
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
