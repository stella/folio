/**
 * Property-based tests for incremental paragraph measurement: re-measuring only
 * the blocks a dirty range touches yields the measures a full pass would.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { FlowBlock, Measure, ParagraphBlock } from "../layout-engine/types";
import { findDirtyBlockIndexes, tryBuildIncrementalMeasures } from "./incrementalMeasure";
import type { DirtyRange } from "./incrementalMeasure";

setDefaultTimeout(propertyTestTimeout(30_000));

type ParagraphSpec = {
  text: string;
};

const paragraphSpec = fc.record({
  text: fc.string({ minLength: 0, maxLength: 120 }),
});

describe("incremental paragraph measurement", () => {
  test("remeasures only blocks touched by arbitrary edit ranges", () => {
    fc.assert(
      fc.property(
        fc.array(paragraphSpec, { minLength: 1, maxLength: 200 }),
        fc.integer({ min: 0, max: 20_000 }),
        fc.integer({ min: 0, max: 20_000 }),
        (paragraphs, firstPosition, secondPosition) => {
          const previousBlocks = makeParagraphBlocks(paragraphs);
          const nextBlocks = previousBlocks.map((block) => ({ ...block }));
          const previousMeasures = previousBlocks.map(fakeMeasureBlock);
          const fullNextMeasures = nextBlocks.map(fakeMeasureBlock);
          const widths = Array.from({ length: nextBlocks.length }, () => 624);
          const dirtyRange = normalizeDirtyRange({
            from: firstPosition,
            to: secondPosition,
          });
          const expectedDirtyIndexes = findDirtyBlockIndexes(nextBlocks, dirtyRange);

          const result = tryBuildIncrementalMeasures({
            previousBlocks,
            previousMeasures,
            previousBlockWidths: widths,
            nextBlocks,
            nextBlockWidths: widths,
            dirtyRange,
            measureBlock: fakeMeasureBlock,
          });

          if (expectedDirtyIndexes.length === 0) {
            expect(result).toBeNull();
            return;
          }

          expect(result?.measuredBlockIndexes).toEqual(expectedDirtyIndexes);
          expect(result?.measures).toEqual(fullNextMeasures);
        },
      ),
      propertyConfig({ numRuns: 1000 }),
    );
  });

  test("matches full measurement after a localized paragraph edit shifts later positions", () => {
    fc.assert(
      fc.property(
        fc.array(paragraphSpec, { minLength: 2, maxLength: 200 }),
        fc.integer({ min: 0, max: 20_000 }),
        fc.string({ minLength: 0, maxLength: 160 }),
        (paragraphs, dirtyIndexSeed, replacementText) => {
          const previousBlocks = makeParagraphBlocks(paragraphs);
          const dirtyIndex = dirtyIndexSeed % paragraphs.length;
          const nextParagraphs = paragraphs.map((paragraph, index) =>
            index === dirtyIndex ? { text: replacementText } : paragraph,
          );
          const nextBlocks = makeParagraphBlocks(nextParagraphs);
          const previousMeasures = previousBlocks.map(fakeMeasureBlock);
          const fullNextMeasures = nextBlocks.map(fakeMeasureBlock);
          const widths = Array.from({ length: nextBlocks.length }, () => 624);
          const dirtyBlock = nextBlocks[dirtyIndex];
          if (!dirtyBlock) {
            return;
          }

          const result = tryBuildIncrementalMeasures({
            previousBlocks,
            previousMeasures,
            previousBlockWidths: widths,
            nextBlocks,
            nextBlockWidths: widths,
            dirtyRange: {
              from: dirtyBlock.pmStart,
              to: dirtyBlock.pmEnd,
            },
            measureBlock: fakeMeasureBlock,
          });

          expect(result?.measuredBlockIndexes).toEqual([dirtyIndex]);
          expect(result?.measures).toEqual(fullNextMeasures);
        },
      ),
      propertyConfig({ numRuns: 1000 }),
    );
  });
});

function normalizeDirtyRange(dirtyRange: DirtyRange): DirtyRange {
  return {
    from: Math.min(dirtyRange.from, dirtyRange.to),
    to: Math.max(dirtyRange.from, dirtyRange.to),
  };
}

function makeParagraphBlocks(specs: ParagraphSpec[]): ParagraphBlock[] {
  const blocks: ParagraphBlock[] = [];
  let pmStart = 0;

  for (let i = 0; i < specs.length; i += 1) {
    const text = specs[i]?.text ?? "";
    const pmEnd = pmStart + text.length + 2;
    blocks.push({
      kind: "paragraph",
      id: `block-${i}`,
      runs: [
        {
          kind: "text",
          text,
          pmStart: pmStart + 1,
          pmEnd: pmStart + 1 + text.length,
        },
      ],
      pmStart,
      pmEnd,
    });
    pmStart = pmEnd + 1;
  }

  return blocks;
}

function fakeMeasureBlock(block: FlowBlock): Measure {
  if (block.kind !== "paragraph") {
    throw new Error("Expected paragraph block");
  }

  const textLength = block.runs.reduce(
    (sum, run) => sum + (run.kind === "text" ? run.text.length : 1),
    0,
  );
  const lineCount = Math.max(1, Math.ceil(textLength / 60));

  return {
    kind: "paragraph",
    lines: Array.from({ length: lineCount }, (_, index) => ({
      ascent: 12,
      descent: 4,
      fromChar: index * 60,
      fromRun: 0,
      lineHeight: 16,
      toChar: Math.min((index + 1) * 60, textLength),
      toRun: 0,
      width: Math.min(624, textLength * 7),
    })),
    totalHeight: lineCount * 16,
  };
}
