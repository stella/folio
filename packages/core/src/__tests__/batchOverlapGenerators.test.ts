/** Frozen pre-extraction generator from batchOverlap.test.ts; this test checks
 * generation and shrink sequences rather than reimplementing the batch oracle. */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import pinnedSeeds from "../../../../test/property-seeds/packages%2Fcore%2Fsrc%2Fai-edits%2FbatchOverlap.test.ts.json";

import {
  blockText as extractedBlockText,
  createBatchOverlapBatchArbitrary,
  token as extractedToken,
} from "./batchOverlapGenerators";

const BLOCK_COUNT = 4;
const TOKEN_COUNT = 4;

const token = (block: number, index: number): string => `t${String(block)}${String(index)}x`;
const blockText = (block: number): string =>
  Array.from({ length: TOKEN_COUNT }, (_, index) => token(block, index)).join(" ");

type Span = { block: number; first: number; last: number };

type GeneratedOperation =
  | ({ kind: "replaceInBlock" | "replaceRange" | "formatRange" | "commentOnRange" } & Span)
  | { kind: "splitBlock"; block: number; before: number }
  | { kind: "mergeBlockWithNext"; block: number; separator?: string }
  | {
      kind:
        | "commentOnBlock"
        | "deleteBlock"
        | "replaceBlock"
        | "setBlockParagraphProperties"
        | "insertAfterBlock"
        | "insertBeforeBlock";
      block: number;
    };

const spanArbitrary = fc
  .record({
    block: fc.nat({ max: BLOCK_COUNT - 1 }),
    first: fc.nat({ max: TOKEN_COUNT - 1 }),
    length: fc.nat({ max: 1 }),
  })
  .map(({ block, first, length }) => ({
    block,
    first,
    last: Math.min(TOKEN_COUNT - 1, first + length),
  }));

const operationArbitrary: fc.Arbitrary<GeneratedOperation> = fc.oneof(
  fc
    .tuple(
      fc.constantFrom(
        "replaceInBlock" as const,
        "replaceRange" as const,
        "formatRange" as const,
        "commentOnRange" as const,
      ),
      spanArbitrary,
    )
    .map(([kind, { block, first, last }]) => ({ kind, block, first, last })),
  fc
    .record({
      block: fc.nat({ max: BLOCK_COUNT - 1 }),
      before: fc.integer({ min: 1, max: TOKEN_COUNT - 1 }),
    })
    .map(({ block, before }) => ({ kind: "splitBlock" as const, block, before })),
  fc
    .tuple(
      fc.constantFrom(
        "commentOnBlock" as const,
        "deleteBlock" as const,
        "replaceBlock" as const,
        "setBlockParagraphProperties" as const,
        "insertAfterBlock" as const,
        "insertBeforeBlock" as const,
      ),
      fc.nat({ max: BLOCK_COUNT - 1 }),
    )
    .map(([kind, block]) => ({ kind, block })),
  fc.nat({ max: BLOCK_COUNT - 2 }).map((block) => ({ kind: "mergeBlockWithNext" as const, block })),
);

const SEEDS = [
  ...new Set([
    0,
    1,
    -1,
    42,
    1142,
    2147483647,
    -2147483648,
    ...Object.values(pinnedSeeds).flatMap((records) => records.map(({ seed }) => seed)),
  ]),
];
const MODES = ["direct", "tracked-changes"] as const;
const originalBatchArbitrary = () => fc.array(operationArbitrary, { minLength: 2, maxLength: 7 });

const batchAndMode = (batch: fc.Arbitrary<GeneratedOperation[]>) =>
  fc.tuple(batch, fc.constantFrom(...MODES));

describe("batch-overlap generator extraction", () => {
  test("preserves the token document byte for byte", () => {
    for (let block = 0; block < BLOCK_COUNT; block++) {
      expect(extractedBlockText(block)).toBe(blockText(block));
      for (let index = 0; index < TOKEN_COUNT; index++) {
        expect(extractedToken(block, index)).toBe(token(block, index));
      }
    }
  });

  test.each(SEEDS)(
    "generates byte-identical ordered batch-and-mode sequences (seed %i)",
    (seed) => {
      const original = fc.sample(batchAndMode(originalBatchArbitrary()), { seed, numRuns: 150 });
      const extracted = fc.sample(batchAndMode(createBatchOverlapBatchArbitrary()), {
        seed,
        numRuns: 150,
      });
      expect(JSON.stringify(extracted)).toBe(JSON.stringify(original));
    },
  );

  test.each(SEEDS)("preserves the complete shrink trace and replay path (seed %i)", (seed) => {
    // Deliberately fail every sample to exercise the complete shrink sequence.
    const check = (batch: fc.Arbitrary<GeneratedOperation[]>, path?: string) => {
      const result = fc.check(
        fc.property(batchAndMode(batch), () => false),
        { seed, numRuns: 1, verbose: 2, ...(path === undefined ? {} : { path }) },
      );
      expect(result.failed).toBe(true);
      return {
        counterexample: result.counterexample,
        counterexamplePath: result.counterexamplePath,
        numShrinks: result.numShrinks,
        failures: result.failures,
        executionSummary: result.executionSummary,
      };
    };
    const original = check(originalBatchArbitrary());
    const extracted = check(createBatchOverlapBatchArbitrary());
    expect(JSON.stringify(extracted)).toBe(JSON.stringify(original));
    const path = original.counterexamplePath;
    expect(path).not.toBeNull();
    if (path === null) panic("The deliberately failing reference has no replay path.");
    expect(JSON.stringify(check(createBatchOverlapBatchArbitrary(), path))).toBe(
      JSON.stringify(check(originalBatchArbitrary(), path)),
    );
  });
});
