/**
 * Planning a piece of a tracked edit applies it with no new ids first and
 * counts the ids it needs only when it took some. The plan that comes out is
 * the plan that counting every piece gives: the same operations, the same
 * document after each, and the same refusals, over random runs of direct and
 * tracked pieces on documents with and without pending changes, with pools
 * of new ids from empty to ample.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { createTrackedPlan } from "../plan";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "../types";
import {
  documentArbitrary,
  opFor,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  stampFor,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

type Piece = Parameters<ReturnType<typeof createTrackedPlan>["append"]>[0];

/** The piece a plan takes for a drawn operation, without its own new ids. */
const pieceOf = (op: DocumentOp): Piece | undefined => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS: {
      const piece = { ...op };
      Reflect.deleteProperty(piece, "newIds");
      return piece;
    }
    default:
      return undefined;
  }
};

const ids = (base: number, length: number): number[] =>
  Array.from({ length }, (_, offset) => base + offset);

/** What a plan holds after a piece: its outcome, operations and document. */
const stateOf = (
  plan: ReturnType<typeof createTrackedPlan>,
  outcome: ReturnType<ReturnType<typeof createTrackedPlan>["append"]>,
): { outcome: unknown; ops: DocumentOp[]; document: Document } => ({
  outcome: outcome.isOk()
    ? "applied"
    : { reason: outcome.error.reason, message: outcome.error.message },
  ops: plan.ops,
  document: plan.document(),
});

describe("planning tracked pieces", () => {
  test("gives the plan that counting every piece's ids gives", async () => {
    await assertProperty(
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        opSeedArbitrary,
        fc.nat({ max: 8 }),
        fc.nat({ max: 3 }),
        fc.array(fc.record({ seed: opSeedArbitrary, tracked: fc.boolean() }), {
          minLength: 1,
          maxLength: 10,
        }),
        (document, stampSeed, revisionCount, controlCount, pieces) => {
          const options = {
            document,
            revision: stampFor(document, stampSeed, 0),
            newIds: {
              revision: ids(40_000, revisionCount),
              control: ids(30_000, controlCount),
            },
          };
          const planned = createTrackedPlan(options);
          const counted = createTrackedPlan({
            ...options,
            countEveryPiece: true,
          });
          for (const [index, { seed, tracked }] of pieces.entries()) {
            const current = planned.document();
            const piece = pieceOf(
              tracked
                ? trackedOpFor(current, seed, index)
                : opFor(current, seed),
            );
            if (piece === undefined) continue;
            const outcome = planned.append(piece);
            expect(stateOf(planned, outcome)).toEqual(
              stateOf(counted, counted.append(piece)),
            );
            if (outcome.isErr()) return;
          }
        },
      ),
      { id: "tracked plan allocation equals counting", numRuns: 300 },
    );
  });
});
