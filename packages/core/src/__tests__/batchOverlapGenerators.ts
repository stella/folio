/** Test-only batch-overlap fixtures and arbitraries; preserve random draw and shrink order. */
import fc from "fast-check";

export const BLOCK_COUNT = 4;
export const TOKEN_COUNT = 4;

export const token = (block: number, index: number): string => `t${String(block)}${String(index)}x`;
export const blockText = (block: number): string =>
  Array.from({ length: TOKEN_COUNT }, (_, index) => token(block, index)).join(" ");

export type Span = { block: number; first: number; last: number };

export type GeneratedOperation =
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

export const operationArbitrary: fc.Arbitrary<GeneratedOperation> = fc.oneof(
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

/** Construct at the caller's original array boundary; no extra filtering or mapping. */
export const createBatchOverlapBatchArbitrary = () =>
  fc.array(operationArbitrary, { minLength: 2, maxLength: 7 });
