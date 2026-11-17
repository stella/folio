/** Exact inverse and journal sequence laws; fixtures do not register tests. */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpType,
} from "../types";
import {
  applyAll,
  count,
  COVERAGE_RUNS,
  expectEveryKindApplied,
  expectRestores,
  inTurn,
  INVERSE_KINDS,
  NUM_RUNS,
  SEQUENCE_COVERAGE_RUNS,
  type Tally,
} from "./applyPropertySupport";
import {
  documentArbitrary,
  independentCopy,
  opFor,
  type OpSeed,
  opSeedArbitrary,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

describe("document operations", () => {
  test("an operation's inverse restores the document exactly", () => {
    const law = (tally: Tally) =>
      fc.property(documentArbitrary, opSeedArbitrary, (document, seed) => {
        const op = opFor(document, seed);
        const original = structuredClone(document);
        const applied = applyDocumentOp(document, op);
        // Applying, refused or not, never writes into its input.
        expect(document).toStrictEqual(original);
        if (applied.isErr()) {
          count(tally, "refused");
          return;
        }
        count(tally, op.type);
        const allowed: readonly DocumentOpType[] = INVERSE_KINDS[op.type];
        for (const inverse of applied.value.inverse) {
          expect(allowed).toContain(inverse.type);
        }
        // Note cleanup composes with the primary inverse: one delta can undo
        // removed bodies, another can preserve prior orphan declarations.
        let noteLifecycle = 0;
        const primaryInverse = applied.value.inverse.filter((inverse) => {
          if (inverse.type !== DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS)
            return true;
          const fields = Object.keys(inverse.parts).filter(
            (key) => key !== "undefinedFields",
          );
          if (!fields.every((key) => key === "footnotes" || key === "endnotes"))
            return true;
          expect(
            Object.keys(inverse.expected)
              .filter((key) => key !== "undefinedFields")
              .sort(),
          ).toEqual(fields.toSorted());
          for (const capture of [inverse.parts, inverse.expected])
            for (const presence of capture.undefinedFields ?? []) {
              expect(presence.target).toBe("package");
              expect(presence.keys.every((key) => fields.includes(key))).toBe(
                true,
              );
            }
          noteLifecycle++;
          return false;
        });
        expect(noteLifecycle).toBeLessThanOrEqual(2);
        if (
          op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK ||
          op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS
        ) {
          // One structural inverse, at most one review restoration, and at most
          // one content restoration per original paragraph; no duplicate targets.
          const restoring = primaryInverse.filter(
            (inverse) => inverse.type === DOCUMENT_OP_TYPES.REPLACE_INLINE,
          );
          const review = primaryInverse.filter(
            (inverse) =>
              inverse.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          );
          expect(new Set(restoring.map(({ blockId }) => blockId)).size).toBe(
            restoring.length,
          );
          expect(restoring.length).toBeLessThanOrEqual(
            op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK ? 1 : 2,
          );
          expect(review.length).toBeLessThanOrEqual(1);
          expect(primaryInverse.length - restoring.length - review.length).toBe(
            1,
          );
        } else if (op.type !== DOCUMENT_OP_TYPES.SET_RUN_PROPS) {
          expect(primaryInverse.length).toBeLessThanOrEqual(1);
        }
        expectRestores(applied.value, original);
      });
    assertProperty(law(new Map()), {
      numRuns: NUM_RUNS,
      id: "an operation's inverse restores the document exactly",
    });
    const tally: Tally = new Map();
    checkCoverageSample(law(tally), COVERAGE_RUNS, {
      map: ([document, seed], index) => [document, inTurn(seed, index)],
    });
    expectEveryKindApplied(tally);
  });

  test(
    "a sequence's inverses in reverse, and a batch's, restore the document exactly",
    () => {
      const law = (tally: Tally) =>
        fc.property(
          documentArbitrary,
          fc.array(opSeedArbitrary, { minLength: 2, maxLength: 8 }),
          (document, seeds: OpSeed[]) => {
            const original = structuredClone(document);
            let current = document;
            const ops: DocumentOp[] = [];
            const inverses: (readonly DocumentOp[])[] = [];
            for (const seed of seeds) {
              const op = opFor(current, seed);
              const applied = applyDocumentOp(current, op);
              if (applied.isErr()) {
                count(tally, "refused");
                continue;
              }
              count(tally, op.type);
              ops.push(op);
              inverses.push(applied.value.inverse);
              current = applied.value.document;
            }
            expect(
              applyAll(current, inverses.toReversed().flat()),
            ).toStrictEqual(original);

            const batch = applyDocumentOps(document, ops);
            if (batch.isErr()) {
              throw batch.error;
            }
            expect(batch.value.document).toStrictEqual(current);
            expectRestores(batch.value, original);

            // JSON is the journal wire format; IDs and section/list model state survive both ways.
            const wireOps = JSON.parse(JSON.stringify(ops)) as DocumentOp[];
            const wireBatch = applyDocumentOps(
              independentCopy(document),
              wireOps,
            );
            if (wireBatch.isErr()) throw wireBatch.error;
            expect(wireBatch.value.document).toStrictEqual(current);
            const wireInverse = JSON.parse(
              JSON.stringify(wireBatch.value.inverse),
            ) as DocumentOp[];
            const wireUndo = applyDocumentOps(
              independentCopy(wireBatch.value.document),
              wireInverse,
            );
            if (wireUndo.isErr()) throw wireUndo.error;
            expect(wireUndo.value.document).toStrictEqual(original);
          },
        );
      assertProperty(law(new Map()), {
        numRuns: NUM_RUNS,
        id: "a sequence's inverses in reverse, and a batch's, restore the document exactly",
      });
      const tally: Tally = new Map();
      checkCoverageSample(law(tally), SEQUENCE_COVERAGE_RUNS, {
        map: ([document, seeds], index) => [
          document,
          seeds.map((seed, offset) => inTurn(seed, index + offset)),
        ],
      });
      expectEveryKindApplied(tally);
    },
    propertyTestTimeout(480_000),
  );
});
