import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document, ParagraphContent } from "../../../model/document";
import { applyDocumentOp } from "../../apply";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
} from "../../types";
import { applyBatch } from "../applyBatch";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  type SequencedBatch,
} from "../envelope";
import { isBlockScoped } from "../scope";
import { TRACKED_ORDER_PAIRS, trackedOrderPair } from "../trackedTransform";
import { transformBatch } from "../transform";

const at = {
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  offset: 0,
  zeroWidthBefore: 0,
} as const;
const batchOf = (opId: string, ops: readonly DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops,
});
const tracked = {
  type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
  at,
  slice: {
    content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
    openStart: 0,
    openEnd: 0,
  },
  revision: { id: 20, author: "B", date: "2026-01-01T00:00:00Z" },
  newIds: { revision: [], control: [] },
} as const satisfies DocumentOp;

// Open deletion inverses restore a source seam; a concurrent insertion changes
// that seam. The enforced footprint refuses them instead of promising that the
// authority can choose an order for two ordinary tracked positional edits.
test(
  "source-seam inverses keep their scoped refusal beside tracked insertions",
  async () => {
    await assertProperty(
      fc.property(
        fc.array(fc.constantFrom("a", "β", "中"), {
          minLength: 2,
          maxLength: 12,
        }),
        fc.nat(),
        fc.integer({ min: 0, max: 4 }),
        (characters, seed, depth) => {
          let content: ParagraphContent[] = [
            {
              type: "run",
              content: [{ type: "text", text: characters.join("") }],
            },
          ];
          for (let index = 0; index < depth; index += 1)
            content = [
              {
                type: "inlineWrapper",
                kind: "bidi",
                control: "embedding",
                content,
              },
            ];
          const document: Document = {
            package: {
              document: {
                content: [{ type: "paragraph", paraId: at.blockId, content }],
              },
            },
          };
          const deleted = applyDocumentOp(document, {
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: at,
            to: { ...at, offset: 1 + (seed % (characters.length - 1)) },
          }).unwrap();
          const inverse = deleted.inverse.at(0);
          expect(inverse?.type).toBe(DOCUMENT_OP_TYPES.INSERT_CONTENT);
          if (inverse?.type !== DOCUMENT_OP_TYPES.INSERT_CONTENT)
            throw new TypeError("Expected the producer's source-seam inverse");
          expect(
            inverse.slice.openStart + inverse.slice.openEnd,
          ).toBeGreaterThan(0);
          expect(isBlockScoped(inverse)).toBe(true);
          const inserted = applyBatch(deleted.document, [tracked]).unwrap();
          const remote = {
            ...batchOf("tracked", [tracked]),
            revision: 1,
            effects: inserted.effects,
          } as const satisfies SequencedBatch;
          const refused = transformBatch(batchOf("undo", deleted.inverse), [
            remote,
          ]);
          expect(refused.isErr()).toBe(true);
          if (refused.isErr())
            expect(refused.error.reason).toBe(
              BATCH_REJECTION_REASONS.BLOCK_CONFLICT,
            );
          expect(trackedOrderPair(inverse, tracked)).toBeUndefined();
          expect(trackedOrderPair(tracked, inverse)).toBeUndefined();
          const direct = {
            type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
            at,
            slice: {
              content: [
                {
                  type: "run",
                  content: [{ type: "text", text: characters.join("") }],
                },
              ],
              openStart: 0,
              openEnd: 0,
            },
          } as const satisfies DocumentOp;
          expect(isBlockScoped(direct)).toBe(false);
          expect(trackedOrderPair(direct, tracked)).toBe(
            TRACKED_ORDER_PAIRS.CUT_AT_TRACKED_CHANGE,
          );
          expect(trackedOrderPair(tracked, direct)).toBe(
            TRACKED_ORDER_PAIRS.CUT_AT_TRACKED_CHANGE,
          );
          expect(
            transformBatch(batchOf("direct", [direct]), [remote]).isOk(),
          ).toBe(true);
        },
      ),
      {
        id: "tracked insertion policy preserves source seam ownership",
        numRuns: 60,
      },
    );
  },
  propertyTestTimeout(30_000),
);
