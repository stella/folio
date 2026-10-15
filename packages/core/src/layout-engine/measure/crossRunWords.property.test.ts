import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { ParagraphBlock } from "../types";
import { collectCrossRunWord, computeFollowingTextLeads } from "./crossRunWords";
import { MAX_HYPHENATION_WORD_LENGTH } from "./lineBreakProvider";
import { getMeasureProvider, setMeasureProvider } from "./measureProvider";

const paragraph = (texts: string[]): ParagraphBlock => ({
  kind: "paragraph",
  id: "00000001",
  runs: texts.map((text) => ({ kind: "text", text })),
});

test(
  "following-run leads retain complete supplementary characters",
  async () => {
    await assertProperty(
      fc.property(fc.integer({ min: 0x10000, max: 0x10ffff }), (point) => {
        const character = String.fromCodePoint(point);
        expect(computeFollowingTextLeads(paragraph(["a", "", `${character}tail`]))).toEqual([
          character,
          character,
          undefined,
        ]);
      }),
      { numRuns: 100, id: "following run leads retain complete Unicode characters" },
    );
  },
  propertyTestTimeout(10_000),
);

test(
  "complete-scalar lookahead retains the cross-run word budget",
  async () => {
    const previous = getMeasureProvider();
    setMeasureProvider({ ...previous, measureTextWidth: (text) => text.length });
    try {
      await assertProperty(
        fc.property(fc.integer({ min: 0x1d400, max: 0x1d419 }), (point) => {
          const character = String.fromCodePoint(point);
          const block = paragraph([
            "a",
            `${"a".repeat(MAX_HYPHENATION_WORD_LENGTH - 1)}${character}tail`,
          ]);
          expect(collectCrossRunWord({ block, startRunIndex: 0, startChar: 0 })).toBeUndefined();
        }),
        { numRuns: 30, id: "Unicode lookahead preserves cross-run word limits" },
      );
    } finally {
      setMeasureProvider(previous);
    }
  },
  propertyTestTimeout(10_000),
);
