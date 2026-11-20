import { expect, spyOn, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type {
  Document,
  ParagraphContent,
  RunContent,
} from "../../model/document";
import * as apply from "../apply";
import { appendTrackedDeletion, createTrackedPlan } from "../plan";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, type TextPosition } from "../types";

const PARA_ID = "0A000001";
const REVISION = { id: 31, author: "Writer", date: "2026-01-01T00:00:00Z" };
const NEW_IDS = { revision: [32, 33, 34, 35] };
const run = (...content: RunContent[]): ParagraphContent => ({
  type: "run",
  content,
});
const text = (value: string): RunContent => ({ type: "text", text: value });
const fixture = (content: ParagraphContent[]): Document => ({
  package: {
    document: { content: [{ type: "paragraph", paraId: PARA_ID, content }] },
  },
});
const at = (offset: number): TextPosition => ({
  story: OP_STORIES.MAIN,
  blockId: PARA_ID,
  offset,
});
type PlanOptions = {
  document: Document;
  from: TextPosition;
  to: TextPosition;
  countEveryPiece: boolean;
};
const planWithCount = ({
  document,
  from,
  to,
  countEveryPiece,
}: PlanOptions) => {
  const applyOp = apply.applyDocumentOp;
  let direct = 0;
  let tracked = 0;
  const applications = spyOn(apply, "applyDocumentOp").mockImplementation(
    (before, op) => {
      if (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE) {
        if (op.revision === undefined) direct += 1;
        else tracked += 1;
      }
      return applyOp(before, op);
    },
  );
  try {
    const options = { from, to, revision: REVISION, newIds: NEW_IDS };
    const plan = createTrackedPlan({
      document,
      revision: REVISION,
      newIds: NEW_IDS,
      countEveryPiece,
    });
    const result = appendTrackedDeletion({ document, options, plan });
    return { result, ops: plan.ops, direct, tracked };
  } finally {
    applications.mockRestore();
  }
};

const alias = (position: TextPosition): TextPosition => ({
  ...position,
  blockId: position.blockId.toLowerCase(),
});

for (const countEveryPiece of [false, true]) {
  for (const endpoints of ["default", "explicit"] as const) {
    test(`reuses one direct check for an unchanged range (${endpoints}, countEveryPiece=${countEveryPiece})`, () => {
      const document = fixture([
        run(
          text("a"),
          { type: "renderedPageBreak" },
          text("bc"),
          { type: "renderedPageBreak" },
          text("d"),
        ),
      ]);
      const from =
        endpoints === "default" ? at(1) : { ...at(1), zeroWidthBefore: 1 };
      const to =
        endpoints === "default" ? at(3) : { ...at(3), zeroWidthBefore: 0 };
      const current = planWithCount({ document, from, to, countEveryPiece });
      const fallback = planWithCount({
        document,
        from: alias(from),
        to: alias(to),
        countEveryPiece,
      });
      expect(current.result.isOk()).toBe(true);
      expect(fallback.result.isOk()).toBe(true);
      expect(current.direct).toBe(1);
      expect(fallback.direct).toBe(2);
      expect(current.tracked).toBe(countEveryPiece ? 2 : 1);
      expect(fallback.tracked).toBe(current.tracked);
      assertExactModel(current.ops, fallback.ops);
      const applied = apply.applyDocumentOps(document, current.ops);
      const legacy = apply.applyDocumentOps(document, fallback.ops);
      if (applied.isErr()) throw applied.error;
      if (legacy.isErr()) throw legacy.error;
      assertExactModel(applied.value, legacy.value);
      const undone = apply.applyDocumentOps(
        applied.value.document,
        applied.value.inverse,
      );
      if (undone.isErr()) throw undone.error;
      assertExactModel(undone.value.document, document);
    });
  }
}

test("a comment boundary changes the surviving segment's gaps", () => {
  const document = fixture([
    { type: "commentRangeStart", id: 1 },
    run(text("abcd")),
    { type: "commentRangeEnd", id: 1 },
    run({ type: "commentReference", id: 1 }),
  ]);
  document.package.document.comments = [
    {
      id: 1,
      author: "Reader",
      date: REVISION.date,
      anchorKind: "range",
      content: [
        { type: "paragraph", paraId: "0A000002", content: [run(text("Note"))] },
      ],
    },
  ];
  const from = { ...at(0), zeroWidthBefore: 0 };
  const to = at(1);
  const outcome = planWithCount({ document, from, to, countEveryPiece: false });
  expect(outcome.result.isOk()).toBe(true);
  expect(outcome.ops.length).toBe(1);
  expect(outcome.ops.at(0)).toMatchObject({
    from: { offset: 0, zeroWidthBefore: 1 },
  });
  expect(outcome.direct).toBe(2);
});

test("mixed tracking and retraction still simulates each distinct segment", () => {
  const document = fixture([
    run(text("a")),
    {
      type: "insertion",
      info: { ...REVISION, id: 20 },
      content: [run(text("b"))],
    },
    run(text("c")),
  ]);
  const outcome = planWithCount({
    document,
    from: at(0),
    to: at(3),
    countEveryPiece: false,
  });
  expect(outcome.result.isOk()).toBe(true);
  expect(outcome.ops.length).toBe(3);
  expect(outcome.direct).toBe(5);
  expect(outcome.tracked).toBe(2);
});

test("invalid zero-width and surrogate endpoints refuse before a tracked edit", () => {
  const document = fixture([run(text("a😀b"))]);
  for (const { from, reason } of [
    {
      from: { ...at(0), zeroWidthBefore: 1 },
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
    },
    { from: at(2), reason: DOCUMENT_OP_REFUSAL_REASONS.SPLITS_SURROGATE_PAIR },
  ]) {
    const { result, direct, tracked, ops } = planWithCount({
      document,
      from,
      to: at(3),
      countEveryPiece: false,
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.reason).toBe(reason);
    expect(direct).toBe(1);
    expect(tracked).toBe(0);
    expect(ops).toEqual([]);
  }
});

test("tracked refusal is still reported after the direct preflight succeeds", () => {
  const document = fixture([
    run(text("abcd")),
    {
      type: "insertion",
      info: { ...REVISION, author: "Reader" },
      content: [run(text("e"))],
    },
  ]);
  for (const countEveryPiece of [false, true]) {
    const current = planWithCount({
      document,
      from: at(1),
      to: at(3),
      countEveryPiece,
    });
    const fallback = planWithCount({
      document,
      from: alias(at(1)),
      to: alias(at(3)),
      countEveryPiece,
    });
    expect(current.result.isErr()).toBe(true);
    expect(fallback.result.isErr()).toBe(true);
    if (current.result.isErr() && fallback.result.isErr()) {
      expect(current.result.error).toMatchObject({
        opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
        reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        message: "Revision id 31 is already used in the package.",
      });
      assertExactModel(current.result.error, fallback.result.error);
    }
    expect(current.direct).toBe(1);
    expect(fallback.direct).toBe(1);
    expect(current.tracked).toBe(countEveryPiece ? 1 : 2);
    expect(fallback.tracked).toBe(current.tracked);
    expect(current.ops).toEqual([]);
  }
});

test("paragraph joins still consume the direct simulations", () => {
  const document = fixture([run(text("ab"))]);
  const follower = "0A000002";
  document.package.document.content.push({
    type: "paragraph",
    paraId: follower,
    content: [run(text("cd"))],
  });
  const outcome = planWithCount({
    document,
    from: at(1),
    to: { ...at(1), blockId: follower },
    countEveryPiece: false,
  });
  if (outcome.result.isErr()) throw outcome.result.error;

  expect(outcome.direct).toBe(4);
  expect(outcome.ops.map(({ type }) => type)).toEqual([
    DOCUMENT_OP_TYPES.DELETE_RANGE,
    DOCUMENT_OP_TYPES.DELETE_RANGE,
    DOCUMENT_OP_TYPES.JOIN_BLOCKS,
  ]);
});
