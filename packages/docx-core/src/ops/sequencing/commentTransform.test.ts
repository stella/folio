import { describe, expect, test } from "bun:test";

import type { BlockContent, Document } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { createEditorIntentIdAllocator } from "../editorIntent";
import { equalForStaleness } from "../equality";
import {
  type CommentOp,
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type TextPosition,
} from "../types";
import {
  canonicalBase,
  compileCanonicalEdit,
} from "./__tests__/canonicalEdits";
import { createClient } from "./client";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  EXCLUSIVE_PARTS,
  type SequencedBatch,
  type SequencedOpEffect,
} from "./envelope";
import { createSequencer } from "./sequencer";
import { transformBatch } from "./transform";

const BLOCK = "10000000";
const at = (offset: number, blockId = BLOCK): TextPosition => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});
const body = [
  {
    type: "paragraph" as const,
    paraId: "50000001",
    content: [
      { type: "run" as const, content: [{ type: "text" as const, text: "A" }] },
    ],
  },
];
const range = (id: number, from: number, to: number): CommentOp => ({
  type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
  comment: { id, author: "A", date: "2026-01-02T00:00:00.000Z", content: body },
  anchor: { kind: "range", from: at(from), to: at(to) },
});
const reply = (id: number, parentId: number): CommentOp => ({
  type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
  comment: { id, author: "B", content: body },
  anchor: { kind: "reply", parentId },
});
const update = (id: number): CommentOp => ({
  type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
  id,
  content: body,
});
const resolution = (id: number, status: "open" | "resolved"): CommentOp => ({
  type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
  id,
  status,
});
const deletion = (id: number): CommentOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
  id,
  scope: "thread",
});
const restoration = (id: number): CommentOp => ({
  type: DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
  ids: [id],
  scaffoldIds: { revision: [], control: [] },
  expected: {
    relationshipPresence: "present",
    relationships: [],
    listPresence: "present",
    records: [],
    anchors: [],
  },
  state: {
    relationshipPresence: "present",
    relationships: [],
    listPresence: "present",
    records: [],
    anchors: [],
  },
});

const batch = (ops: readonly DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "local",
  actor: "local",
  baseRev: 0,
  ops,
});
const sequenced = (
  op: DocumentOp,
  effect: SequencedOpEffect = { type: "none" },
): SequencedBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "remote",
  actor: "remote",
  baseRev: 0,
  revision: 1,
  ops: [op],
  effects: [effect],
});
const touched = (
  blockIds: readonly string[],
  ids: readonly number[],
  relationshipRemoved = false,
): SequencedOpEffect => ({
  type: "touchedBlocks",
  blockIds,
  comments: { ids, relationshipRemoved },
});

const over = (op: DocumentOp, remote: DocumentOp, effect?: SequencedOpEffect) =>
  transformBatch(batch([op]), [sequenced(remote, effect)]);

const refusal = (result: ReturnType<typeof over>) => {
  if (result.isOk()) throw new Error("The pair must refuse.");
  return result.error;
};

describe("concurrent comment operations", () => {
  test("two content edits of one comment refuse naming the comments part", () => {
    const error = refusal(over(update(7), update(7)));
    expect(error.reason).toBe(BATCH_REJECTION_REASONS.PART_CONFLICT);
    expect(error.part).toBe(EXCLUSIVE_PARTS.COMMENTS);
  });

  test("the same resolution twice passes; opposite resolutions refuse", () => {
    expect(
      over(resolution(7, "resolved"), resolution(7, "resolved")).isOk(),
    ).toBe(true);
    expect(
      refusal(over(resolution(7, "open"), resolution(7, "resolved"))).part,
    ).toBe(EXCLUSIVE_PARTS.COMMENTS);
  });

  test("a content edit and a resolution of one comment pass in either order", () => {
    expect(over(update(7), resolution(7, "resolved")).isOk()).toBe(true);
    expect(over(resolution(7, "resolved"), update(7)).isOk()).toBe(true);
  });

  test("a deletion refuses every other use of its thread", () => {
    // The deleted thread's replies come from the sequenced effect.
    const removed = touched([BLOCK], [7, 8]);
    for (const op of [
      reply(9, 8),
      update(8),
      resolution(7, "resolved"),
      deletion(8),
    ])
      expect(refusal(over(op, deletion(7), removed)).part).toBe(
        EXCLUSIVE_PARTS.COMMENTS,
      );
    expect(refusal(over(deletion(7), resolution(7, "open"))).part).toBe(
      EXCLUSIVE_PARTS.COMMENTS,
    );
    // Another thread passes.
    expect(over(update(3), deletion(7), removed).isOk()).toBe(true);
  });

  test("a reply passes edits of its parent and another reply to it", () => {
    expect(over(reply(9, 7), update(7)).isOk()).toBe(true);
    expect(over(reply(9, 7), resolution(7, "resolved")).isOk()).toBe(true);
    expect(over(reply(9, 7), reply(10, 7), touched([BLOCK], [10])).isOk()).toBe(
      true,
    );
    expect(over(reply(10, 7), reply(9, 7), touched([BLOCK], [9])).isOk()).toBe(
      true,
    );
  });

  test("a resolution after a deletion that removed a part relationship refuses", () => {
    const error = refusal(
      over(resolution(3, "resolved"), deletion(7), touched([], [7], true)),
    );
    expect(error.part).toBe(EXCLUSIVE_PARTS.RELATIONSHIPS);
    expect(
      over(resolution(3, "resolved"), deletion(7), touched([], [7])).isOk(),
    ).toBe(true);
  });

  test("a restoration passes only content edits of other comments", () => {
    expect(over(restoration(7), update(3)).isOk()).toBe(true);
    for (const remote of [
      update(7),
      range(3, 1, 2),
      deletion(3),
      resolution(3, "resolved"),
    ])
      expect(refusal(over(restoration(7), remote, touched([], [3]))).part).toBe(
        EXCLUSIVE_PARTS.COMMENTS,
      );
  });
});

describe("comment anchors", () => {
  const anchorOf = (result: ReturnType<typeof over>) => {
    if (result.isErr()) throw result.error;
    const op = result.value.ops.at(0);
    if (op?.type !== DOCUMENT_OP_TYPES.CREATE_COMMENT)
      throw new Error("A creation must remain.");
    return op.anchor;
  };

  test("a range moves like text over a deletion before it and stays over one after it", () => {
    const erase = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at(0),
      to: at(2),
    } as const;
    expect(anchorOf(over(range(7, 4, 6), erase))).toEqual({
      kind: "range",
      from: at(2),
      to: at(4),
    });
    expect(
      anchorOf(over(range(7, 4, 6), { ...erase, from: at(8), to: at(9) })),
    ).toEqual({ kind: "range", from: at(4), to: at(6) });
  });

  test("a deletion reaching an anchor refuses rather than collapsing the range", () => {
    for (const [from, to] of [
      [3, 7],
      [5, 6],
      [6, 8],
      [2, 4],
    ] as const) {
      const remote = {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: at(from),
        to: at(to),
      } as const;
      expect(refusal(over(range(7, 4, 6), remote)).reason).toBe(
        BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      );
      expect(
        refusal(over(remote, range(7, 4, 6), touched([BLOCK], [7]))).reason,
      ).toBe(BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR);
    }
  });

  test("text after a comment's reference moves past it; text at an anchor refuses", () => {
    const typed = (offset: number) =>
      ({
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at(offset),
        text: "x",
        runProps: "inherit",
      }) as const;
    const effect = touched([BLOCK], [7]);
    const moved = over(typed(8), range(7, 4, 6), effect);
    if (moved.isErr()) throw moved.error;
    expect(moved.value.ops).toEqual([typed(9)]);
    const inside = over(typed(5), range(7, 4, 6), effect);
    if (inside.isErr()) throw inside.error;
    expect(inside.value.ops).toEqual([typed(5)]);
    for (const offset of [4, 6])
      expect(refusal(over(typed(offset), range(7, 4, 6), effect)).reason).toBe(
        BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      );
  });

  test("two ranges on one paragraph move past each other's reference", () => {
    expect(
      anchorOf(over(range(8, 7, 9), range(7, 2, 4), touched([BLOCK], [7]))),
    ).toEqual({ kind: "range", from: at(8), to: at(10) });
    expect(
      refusal(over(range(8, 4, 9), range(7, 2, 4), touched([BLOCK], [7])))
        .reason,
    ).toBe(BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR);
  });

  test("a range refuses a concurrent reply that changed its paragraph", () => {
    expect(
      refusal(over(range(8, 7, 9), reply(9, 3), touched([BLOCK], [9]))).reason,
    ).toBe(BATCH_REJECTION_REASONS.BLOCK_CONFLICT);
    expect(
      over(range(8, 7, 9), reply(9, 3), touched(["10000001"], [9])).isOk(),
    ).toBe(true);
  });
});

test("a comment keeps the author and date it was sent with on every receiver", () => {
  const base = canonicalBase();
  const sequencer = createSequencer(base);
  const lease = sequencer.grantIdLease("actor-0");
  if (lease.isErr()) throw lease.error;
  const ops = compileCanonicalEdit(
    base,
    {
      actor: "actor-0",
      tag: "0",
      lease: lease.value,
      allocate: createEditorIntentIdAllocator(),
    },
    { kind: "comment", a: 24, b: 0, suggest: false },
    5,
  );
  const created = ops?.at(0);
  if (created?.type !== DOCUMENT_OP_TYPES.CREATE_COMMENT)
    throw new Error("The comment edit did not compile.");
  const submitted = sequencer.submit(
    JSON.parse(JSON.stringify({ ...batch([created]), actor: "actor-0" })),
  );
  expect(submitted.type).toBe("ack");
  const receiver = createClient(base);
  const broadcast = sequencer.broadcasts.at(0);
  if (broadcast === undefined)
    throw new Error("The comment was not sequenced.");
  const received = receiver.receiveBroadcast(
    JSON.parse(JSON.stringify(broadcast)),
  );
  if (received.isErr()) throw received.error;
  for (const document of [sequencer.document, receiver.document]) {
    const comment = document.package.document.comments?.at(0);
    expect(comment?.author).toBe("actor-0");
    expect(comment?.date).toBe(created.comment.date);
  }
});

test("two replies to one parent take the same places in either order", () => {
  const base = canonicalBase();
  const sequencer = createSequencer(base);
  const author = (index: number) => {
    const lease = sequencer.grantIdLease(`actor-${index}`);
    if (lease.isErr()) throw lease.error;
    return {
      actor: `actor-${index}`,
      tag: String(index),
      lease: lease.value,
      allocate: createEditorIntentIdAllocator(),
    };
  };
  const compiled = (
    document: typeof base,
    index: number,
    kind: "comment" | "reply",
    serial: number,
  ) => {
    const ops = compileCanonicalEdit(
      document,
      author(index),
      { kind, a: 24, b: 0, suggest: false },
      serial,
    );
    if (ops === undefined) throw new Error(`The ${kind} did not compile.`);
    return ops;
  };
  const apply = (document: typeof base, ops: readonly DocumentOp[]) => {
    const applied = applyDocumentOps(document, ops);
    if (applied.isErr()) throw applied.error;
    return applied.value.document;
  };
  const commented = apply(base, compiled(base, 0, "comment", 5));
  const first = compiled(commented, 1, "reply", 6);
  const second = compiled(commented, 2, "reply", 7);
  // Neither reply refuses the other, and both orders give one document.
  expect(
    transformBatch(batch(second), [
      { ...batch(first), opId: "first", revision: 1 },
    ]).isOk(),
  ).toBe(true);
  const left = apply(apply(commented, first), second);
  const right = apply(apply(commented, second), first);
  expect(equalForStaleness(left, right)).toBe(true);
  // The same with the parent's range ending between blocks, after the
  // paragraph that holds its replies' end markers.
  const between = endBetweenBlocks(commented);
  const early = compiled(between, 1, "reply", 6);
  const late = compiled(between, 2, "reply", 7);
  expect(
    equalForStaleness(
      apply(apply(between, early), late),
      apply(apply(between, late), early),
    ),
  ).toBe(true);
});

/** The document with each inline comment range end moved after its paragraph. */
const endBetweenBlocks = (document: Document): Document => {
  const content = document.package.document.content.flatMap(
    (block): BlockContent[] => {
      if (block.type !== "paragraph") return [block];
      const ends = block.content.filter(
        (node) => node.type === "commentRangeEnd",
      );
      if (ends.length === 0) return [block];
      return [
        {
          ...block,
          content: block.content.filter(
            (node) => node.type !== "commentRangeEnd",
          ),
        },
        ...ends.map((end): BlockContent => end),
      ];
    },
  );
  expect(content.length).toBeGreaterThan(
    document.package.document.content.length,
  );
  return {
    ...document,
    package: {
      ...document.package,
      document: { ...document.package.document, content },
    },
  };
};
