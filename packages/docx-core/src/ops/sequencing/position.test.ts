import { expect, test } from "bun:test";

import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type TextPosition,
} from "../types";
import type { SequencedBatch } from "./envelope";
import { transformPosition } from "./position";

// A note's paragraph: its own mark is a zero-width child at offset 0, and a
// caret just after the mark stands at offset 0 past one zero-width child.
const story = { kind: "footnote", id: 1 } as const;
const at = (offset: number, zeroWidthBefore?: number): TextPosition => ({
  story,
  blockId: "0000000A",
  offset,
  ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
});
const afterMark = at(0, 1);
const batch = (op: DocumentOp): SequencedBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "peer",
  actor: "peer",
  baseRev: 0,
  revision: 1,
  ops: [op],
});

test("a split later in the paragraph leaves a caret after the note's mark in the first half", () => {
  expect(
    transformPosition(
      afterMark,
      [
        {
          ...batch({
            type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
            at: at(5, 0),
            newBlockId: "0000000B",
          }),
          effects: [{ type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, newHalf: "first" }],
        },
      ],
      "before",
    ),
  ).toEqual({ ...afterMark, blockId: "0000000B" });
});

test("a split ahead of a position among zero-width children moves it with them", () => {
  expect(
    transformPosition(
      at(7, 1),
      [
        {
          ...batch({
            type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
            at: at(5, 0),
            newBlockId: "0000000B",
          }),
          effects: [{ type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, newHalf: "first" }],
        },
      ],
      "before",
    ),
  ).toEqual(at(2, 1));
});

test("a direct deletion apart from the note's mark leaves a caret after it", () => {
  expect(
    transformPosition(
      afterMark,
      [
        batch({
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: at(2, 0),
          to: at(4, 0),
        }),
      ],
      "before",
    ),
  ).toEqual(afterMark);
});

test("a direct deletion covering a position among zero-width children leaves it where the deletion starts", () => {
  expect(
    transformPosition(
      at(3, 1),
      [
        batch({
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: at(1, 0),
          to: at(6, 0),
        }),
      ],
      "before",
    ),
  ).toEqual(at(1, 0));
});

test("a split or deletion edge at the position's own offset still leaves it unmapped", () => {
  expect(
    transformPosition(
      afterMark,
      [
        {
          ...batch({
            type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
            at: at(0, 0),
            newBlockId: "0000000B",
          }),
          effects: [{ type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, newHalf: "first" }],
        },
      ],
      "before",
    ),
  ).toBeUndefined();
  expect(
    transformPosition(
      afterMark,
      [
        batch({
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: at(0, 0),
          to: at(3, 0),
        }),
      ],
      "before",
    ),
  ).toBeUndefined();
});
