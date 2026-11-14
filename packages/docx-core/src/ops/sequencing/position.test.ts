import { expect, test } from "bun:test";

import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type TextPosition,
} from "../types";
import type { SequencedBatch } from "./envelope";
import { transformPosition } from "./position";

const story = { kind: "footnote", id: 1 } as const;
const at = (offset: number, zeroWidthBefore?: number): TextPosition => ({
  story,
  blockId: "0000000A",
  offset,
  ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
});
const batch = (ops: readonly DocumentOp[]): SequencedBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "peer",
  actor: "peer",
  baseRev: 0,
  revision: 1,
  ops,
  // What the authority records for inline content: the block it touched.
  effects: ops.map(() => ({
    type: "touchedBlocks" as const,
    blockIds: ["0000000A"],
  })),
});
const typed = (offset: number, zeroWidthBefore: number): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
  at: at(offset, zeroWidthBefore),
  slice: {
    openStart: 0,
    openEnd: 0,
    content: [{ type: "run", content: [{ type: "text", text: "Re" }] }],
  },
});

test("inline content inserted before a position moves it on by its width", () => {
  expect(
    transformPosition(at(13, 0), [batch([typed(0, 0)])], "before"),
  ).toEqual(at(15, 0));
  expect(transformPosition(at(13, 0), [batch([typed(0, 0)])], "after")).toEqual(
    at(15, 0),
  );
});

test("inline content inserted after a position leaves it", () => {
  expect(transformPosition(at(2, 0), [batch([typed(5, 0)])], "before")).toEqual(
    at(2, 0),
  );
});

test("at the content's own place, before keeps a position ahead of it", () => {
  expect(transformPosition(at(4, 0), [batch([typed(4, 0)])], "before")).toEqual(
    at(4, 0),
  );
  expect(transformPosition(at(4, 0), [batch([typed(4, 1)])], "before")).toEqual(
    at(4, 0),
  );
  // Past the zero-width children the content goes after, or moved along by
  // `after`: the content's own zero-width children decide, so it is unmapped.
  expect(
    transformPosition(at(4, 2), [batch([typed(4, 1)])], "before"),
  ).toBeUndefined();
  expect(
    transformPosition(at(4, 0), [batch([typed(4, 0)])], "after"),
  ).toBeUndefined();
  expect(transformPosition(at(4), [batch([typed(4, 0)])], "before")).toBe(
    undefined,
  );
});

test("records cut or merged at a position leave its offset", () => {
  for (const type of [
    DOCUMENT_OP_TYPES.SPLIT_INLINE,
    DOCUMENT_OP_TYPES.JOIN_INLINE,
  ] as const)
    expect(
      transformPosition(
        at(7, 0),
        [batch([{ type, at: at(3, 0), depth: 2 }])],
        "before",
      ),
    ).toEqual(at(7, 0));
});

test("inline content in another paragraph or story leaves a position", () => {
  expect(
    transformPosition(
      { ...at(9, 0), blockId: "0000000B" },
      [batch([typed(0, 0)])],
      "before",
    ),
  ).toEqual({ ...at(9, 0), blockId: "0000000B" });
  expect(
    transformPosition(
      { story: OP_STORIES.MAIN, blockId: "0000000A", offset: 9 },
      [batch([typed(0, 0)])],
      "before",
    ),
  ).toEqual({ story: OP_STORIES.MAIN, blockId: "0000000A", offset: 9 });
});
