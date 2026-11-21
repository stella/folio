import { Result } from "better-result";
import { expect, spyOn, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as blocks from "../blocks";
import { applyCommentOp } from "../comments";
import { formatParaId } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "../refusal";
import {
  type CreateCommentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type TextPosition,
} from "../types";

const paragraphId = (index: number) => formatParaId(0x10000000 + index);
const paragraph = (paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: [{ type: "run", content: [{ type: "text", text: "Clause" }] }],
});
const at = (index: number, offset = 1): TextPosition => ({
  blockId: paragraphId(index),
  story: OP_STORIES.MAIN,
  offset,
});
const source = (width: number): Document => ({
  package: {
    document: {
      content: Array.from({ length: width }, (_, index) =>
        paragraph(paragraphId(index)),
      ),
    },
  },
});

type Range = { from: TextPosition; to: TextPosition };
const comment = ({ from, to }: Range): CreateCommentOp => ({
  type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
  comment: { id: 1, author: "Reader", content: [paragraph("1000FFF0")] },
  anchor: { kind: "range", from, to },
});

/** Count the validation's own full walks before any marker is inserted. */
const boundaryWork = (document: Document, op: CreateCommentOp) => {
  const enumerate = blocks.storyParagraphs;
  let visited = 0;
  const walk = spyOn(blocks, "storyParagraphs").mockImplementation((body) => {
    const located = enumerate(body);
    if (body.content === document.package.document.content)
      visited += located.length;
    return located;
  });
  const stop = new DocumentOpRefusal({
    reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
    message: "Stop at first marker insert",
    opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
  });
  let reachedInsertion = false;
  try {
    const result = applyCommentOp(document, op, (_current, operations) => {
      reachedInsertion = true;
      expect(operations).toHaveLength(1);
      expect(operations.at(0)?.type).toBe(DOCUMENT_OP_TYPES.INSERT_CONTENT);
      return Result.err(stop);
    });
    if (reachedInsertion) {
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error).toBe(stop);
    }
    return { visited, result, reachedInsertion };
  } finally {
    walk.mockRestore();
  }
};

test.each([16, 64])(
  "named range validation does not enumerate %i story paragraphs",
  (width) => {
    const document = source(width);
    const op = comment({ from: at(width - 2), to: at(width - 1) });
    const { visited, reachedInsertion } = boundaryWork(document, op);
    expect(reachedInsertion).toBe(true);
    expect(visited).toBe(0);
  },
);

const assertRestored = (document: Document, op: CreateCommentOp) => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  expect(
    applied.value.document.package.document.comments?.map(({ id }) => id),
  ).toEqual([1]);
  const reverted = applyDocumentOps(
    applied.value.document,
    applied.value.inverse,
  );
  if (reverted.isErr()) throw reverted.error;
  assertExactModel(reverted.value.document, document);
  return applied.value;
};

test("aliases of named boundaries have the same comment and inverse", () => {
  const document = source(16);
  const from = at(10, 0);
  const to = at(10, 3);
  const aliased = {
    from: { ...from, blockId: from.blockId.toLowerCase() },
    to,
  };
  const expected = assertRestored(document, comment({ from, to }));
  const actual = assertRestored(document, comment(aliased));
  assertExactModel(actual, expected);
});

test("cross-list boundaries preserve physical order and the full walk", () => {
  const document = source(3);
  document.package.document.content.splice(1, 1, {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [{ type: "tableCell", content: [paragraph(paragraphId(1))] }],
      },
    ],
  });
  const from = at(0);
  const to = at(1);
  const op = comment({ from, to });
  expect(boundaryWork(document, op)).toMatchObject({
    reachedInsertion: true,
    visited: 3,
  });
  assertRestored(document, op);
  const backwards = comment({ from: to, to: from });
  expect(boundaryWork(document, backwards)).toMatchObject({
    reachedInsertion: false,
    visited: 3,
  });
  const refused = applyDocumentOp(document, backwards);
  expect(refused.isErr()).toBe(true);
  if (refused.isErr())
    expect(refused.error).toMatchObject({
      opType: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
      message: "A comment range cannot run backwards.",
    });
});

test("unnamed endpoints retain document-order lookup before native insertion", () => {
  const document = source(2);
  document.package.document.content.unshift(paragraph());
  const op = comment({ from: { ...at(0), blockId: "" }, to: at(1) });
  expect(boundaryWork(document, op)).toMatchObject({
    reachedInsertion: true,
    visited: 3,
  });
  const refused = applyDocumentOp(document, op);
  expect(refused.isErr()).toBe(true);
  if (refused.isErr())
    expect(refused.error).toMatchObject({
      opType: DOCUMENT_OP_TYPES.CREATE_COMMENT,
      reason: DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID,
      message:
        "A story paragraph has no paraId; ensureParaIds establishes one.",
    });
});

test("a missing name and a backwards gap keep their original native refusal", () => {
  const document = source(16);
  const from = at(10, 3);
  const options = [
    {
      to: at(100),
      reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      message: "A comment boundary paragraph does not exist.",
    },
    {
      to: {
        ...at(10),
        blockId: paragraphId(10).toLowerCase(),
        zeroWidthBefore: 0,
      },
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
      message: "A comment range cannot run backwards.",
    },
  ];
  for (const { to, reason, message } of options) {
    const op = comment({ from, to });
    expect(boundaryWork(document, op)).toMatchObject({
      reachedInsertion: false,
      visited: 0,
    });
    const refused = applyDocumentOp(document, op);
    expect(refused.isErr()).toBe(true);
    if (refused.isErr())
      expect(refused.error).toMatchObject({
        opType: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        reason,
        message,
      });
  }
});

test("same-cell boundaries compare their indexed positions", () => {
  const document = source(3);
  document.package.document.content.splice(1, 1, {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: [paragraph(paragraphId(1)), paragraph(paragraphId(3))],
          },
        ],
      },
    ],
  });
  const op = comment({ from: at(1), to: at(3) });
  expect(boundaryWork(document, op)).toMatchObject({
    reachedInsertion: true,
    visited: 0,
  });
  assertRestored(document, op);
});

test("a repeated named endpoint keeps the first normalized occurrence", () => {
  const document = source(16);
  document.package.document.content.push(
    paragraph(paragraphId(10).toLowerCase()),
  );
  // The original first occurrence comes before paragraph 11. The later
  // occurrence would make the range run backwards.
  const op = comment({
    from: { ...at(10), blockId: paragraphId(10).toLowerCase() },
    to: at(11),
  });
  expect(boundaryWork(document, op)).toMatchObject({
    reachedInsertion: true,
    visited: 17,
  });
});
