import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { prepareTrackedParagraphInsertion } from "../blockInsertion";
import * as blocks from "../blocks";
import { formatParaId } from "../idLease";
import {
  identityKeysIn,
  idKey,
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
  paragraphIdsIn,
} from "../ids";
import * as primitives from "../paragraphPrimitive";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type InsertBlocksOp, OP_STORIES } from "../types";
import { captureDocumentOp } from "../wire";

const paragraph = (id: number): Paragraph => ({
  type: "paragraph",
  paraId: formatParaId(id),
  content: [{ type: "run", content: [{ type: "text", text: String(id) }] }],
});

const insertion = (anchor: number, id: number): InsertBlocksOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: "after", blockId: formatParaId(anchor) },
  blocks: [paragraph(id)],
  revision: { id: id * 10, author: "Author", date: "2026-01-11T00:00:00.000Z" },
  newIds: { revision: [id * 10 + 1, id * 10 + 2] },
});

const checkout = (content: Paragraph[]): Document => {
  const document: Document = {
    package: { document: { content, sections: [{ content, properties: {} }] } },
  };
  return blocks
    .replaceParagraphs({
      document,
      story: OP_STORIES.MAIN,
      at: { list: [], index: 0 },
      count: 0,
      replacement: [],
    })
    .unwrap();
};

const countsOver = (index: { count: (key: string) => number }) => {
  const delta = new Map<string, number>();
  const bump = (keys: readonly string[], by: number) => {
    for (const key of keys) delta.set(key, (delta.get(key) ?? 0) + by);
  };
  return {
    count: (key: string) => index.count(key) + (delta.get(key) ?? 0),
    add: (keys: readonly string[]) => bump(keys, 1),
    remove: (keys: readonly string[]) => bump(keys, -1),
  };
};

test.each(["tail", "afterWithFollower", "before"])(
  "pure insertion preparation uses no story splices and preserves exact native effects (%s)",
  (side) => {
    const tail = side === "tail";
    const insertingBefore = side === "before";
    const first = paragraph(1);
    const second = paragraph(2);
    const virtual = tail ? [first] : [first, second];
    const baseline = checkout(virtual.slice());
    let sequential = baseline;
    const count = 64;
    const ops = Array.from({ length: count }, (_, step) => {
      const native = insertion(tail && step > 0 ? step + 2 : 1, step + 3);
      const op = insertingBefore
        ? ({
            ...native,
            at: { type: "before", blockId: formatParaId(1) },
          } satisfies InsertBlocksOp)
        : native;
      if (step === count - 1) {
        const incoming = op.blocks.at(0);
        if (!incoming) throw new Error("Missing test paragraph");
        return {
          ...op,
          blocks: [{ ...incoming, formatting: { alignment: "center" } }],
        } satisfies InsertBlocksOp;
      }
      return op;
    });
    const expected = ops.map((op) => {
      const applied = applyDocumentOp(sequential, op).unwrap();
      sequential = applied.document;
      return applied;
    });

    const sectionSplices = spyOn(blocks, "replaceParagraphs");
    const validators = spyOn(primitives, "validateBlockReplacement");
    try {
      const paragraphIds = countsOver(
        packageParagraphIdIndex(baseline.package),
      );
      const recordIds = countsOver(packageIdentityKeyIndex(baseline.package));
      for (const [step, op] of ops.entries()) {
        const firstIndex = insertingBefore ? step : 0;
        const index = tail ? virtual.length - 1 : firstIndex;
        const anchor = virtual[index];
        if (!anchor) throw new Error("Missing virtual anchor");
        const following = virtual[index + 1];
        const planned = prepareTrackedParagraphInsertion({
          document: baseline,
          op,
          anchor,
          ...(following === undefined ? {} : { following }),
          paragraphIds,
          recordIds,
        }).unwrap();
        if (planned.type !== "replacement")
          throw new Error("Unsupported test operation");
        expect(planned.before).toBe(
          insertingBefore ? anchor : (following ?? anchor),
        );
        const native = expected[step];
        if (!native) throw new Error("Missing sequential test result");
        expect(planned.effect.inverse.map(captureDocumentOp)).toEqual(
          native.inverse,
        );
        expect(planned.effect.touched).toEqual(native.touched);
        const putAt =
          insertingBefore || following === undefined ? index : index + 1;
        paragraphIds.remove(paragraphIdsIn([planned.before]).map(idKey));
        paragraphIds.add(paragraphIdsIn(planned.operation.blocks).map(idKey));
        recordIds.remove(identityKeysIn([planned.before]));
        recordIds.add(identityKeysIn(planned.operation.blocks));
        virtual.splice(putAt, 1, ...planned.operation.blocks);
        expect(virtual).toEqual(native.document.package.document.content);
      }
      expect(sectionSplices).toHaveBeenCalledTimes(0);
      expect(validators).toHaveBeenCalledTimes(count * 2);
      for (const [index, native] of expected.entries()) {
        const draft = validators.mock.calls[index * 2]?.[0].op;
        const tracked = validators.mock.calls[index * 2 + 1]?.[0].op;
        const inverse = native.inverse.at(0);
        if (
          draft === undefined ||
          tracked === undefined ||
          inverse?.type !== DOCUMENT_OP_TYPES.REPLACE_BLOCKS
        )
          throw new Error("Missing draft or tracked validation");
        expect(draft.expected).toEqual(tracked.expected);
        expect(tracked.blocks).toEqual(inverse.expected);
        expect(draft.blocks).not.toEqual(tracked.blocks);
      }
    } finally {
      validators.mockRestore();
      sectionSplices.mockRestore();
    }
  },
);

test("the pure seam leaves unsupported work to native insertion and retains collision precedence", () => {
  const anchor = paragraph(1);
  const follower = paragraph(2);
  const baseline = checkout([anchor, follower]);
  const paragraphIds = packageParagraphIdIndex(baseline.package);
  const recordIds = packageIdentityKeyIndex(baseline.package);
  const prepare = (op: InsertBlocksOp, following?: Paragraph) =>
    prepareTrackedParagraphInsertion({
      document: baseline,
      op,
      anchor,
      ...(following === undefined ? {} : { following }),
      paragraphIds,
      recordIds,
    });
  let fallbacks = 0;
  const unsupported = {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: formatParaId(1) },
    blocks: [paragraph(3)],
  } satisfies InsertBlocksOp;
  if (prepare(unsupported).unwrap().type === "unsupported") {
    fallbacks++;
    expect(applyDocumentOp(baseline, unsupported).isOk()).toBe(true);
  }
  expect(fallbacks).toBe(1);
  const colliding = insertion(1, 2);
  const native = applyDocumentOp(baseline, colliding);
  const planned = prepare(colliding, follower);
  expect(planned.isErr() && planned.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
  );
  expect(native.isErr()).toBe(true);
  if (planned.isErr() && native.isErr())
    expect(planned.error).toEqual(native.error);
});
