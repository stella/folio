import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  FolioContentBlock,
  FolioContentRun,
} from "../../compare/content-types";
import { inlineFormattingSegments } from "../../compare/formatting";
import {
  PROPERTY_DIFFERENCES,
  type RunPresentation,
} from "./__fixtures__/inlinePropertyDifferences";
import {
  CANONICAL_INLINE_PRESENTATION_PROPERTIES,
  sameCanonicalInlinePresentation,
} from "./inline-presentation";

setDefaultTimeout(propertyTestTimeout(30_000));

const block = (
  text: string,
  runs: readonly FolioContentRun[],
): FolioContentBlock => ({
  id: "block",
  kind: "paragraph",
  text,
  previewRuns: runs,
});

const singleRunBlock = (presentation: RunPresentation): FolioContentBlock =>
  block("Contract", [{ text: "Contract", ...presentation }]);

const PRESENTATION_STATES = [
  {},
  { bold: true },
  { italic: true, directFormatting: { italic: true } },
  { underline: true, directFormatting: { underline: false } },
  { strike: true },
  { fontFamily: "Noto Sans", directFormatting: { fontFamily: "Noto Sans" } },
  { fontSizePt: 10.5, directFormatting: { fontSizePt: null } },
  { color: "#c00000", directFormatting: { color: "C00000" } },
] as const satisfies readonly RunPresentation[];

describe("canonical inline presentation ownership", () => {
  test("the descriptor grammar and exercised property matrix are the same set", () => {
    expect(CANONICAL_INLINE_PRESENTATION_PROPERTIES.toSorted()).toEqual(
      Object.keys(PROPERTY_DIFFERENCES).toSorted(),
    );
  });

  test("every effective and authored property difference drives planning and verification", () => {
    for (const property of CANONICAL_INLINE_PRESENTATION_PROPERTIES) {
      for (const fixture of PROPERTY_DIFFERENCES[property]) {
        const base = singleRunBlock(fixture.base);
        const revised = singleRunBlock(fixture.revised);
        expect({
          property,
          case: fixture.label,
          comparison: inlineFormattingSegments({
            baseBlock: base,
            targetBlock: revised,
            maxSegments: 1,
          }),
          equivalent: sameCanonicalInlinePresentation(base, revised),
        }).toEqual({
          property,
          case: fixture.label,
          comparison: {
            status: "compared",
            segments: [
              {
                startOffset: 0,
                endOffset: "Contract".length,
                formatting: fixture.expected,
              },
            ],
          },
          equivalent: false,
        });
      }
    }
  });

  test("equivalent color spellings and effective false defaults have one projection", () => {
    const base = singleRunBlock({
      color: "#abc",
      directFormatting: { color: "#aBc" },
    });
    const revised = singleRunBlock({
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      color: "ABC",
      directFormatting: { color: "ABC" },
    });

    expect(sameCanonicalInlinePresentation(base, revised)).toBe(true);
    expect(
      inlineFormattingSegments({
        baseBlock: base,
        targetBlock: revised,
        maxSegments: 1,
      }),
    ).toEqual({ status: "compared", segments: [] });
  });

  test("an authored null remains distinct from an absent authored value", () => {
    const base = singleRunBlock({ color: "red", directFormatting: {} });
    const revised = singleRunBlock({
      color: "red",
      directFormatting: { color: null },
    });

    expect(sameCanonicalInlinePresentation(base, revised)).toBe(false);
    expect(
      inlineFormattingSegments({
        baseBlock: base,
        targetBlock: revised,
        maxSegments: 1,
      }),
    ).toEqual({
      status: "compared",
      segments: [
        {
          startOffset: 0,
          endOffset: "Contract".length,
          formatting: { color: null },
        },
      ],
    });
  });

  test("distinct non-hex color tokens remain distinct in planning and verification", () => {
    const base = singleRunBlock({ color: "red" });
    const revised = singleRunBlock({ color: "blue" });

    expect(sameCanonicalInlinePresentation(base, revised)).toBe(false);
    expect(
      inlineFormattingSegments({
        baseBlock: base,
        targetBlock: revised,
        maxSegments: 1,
      }),
    ).toEqual({
      status: "compared",
      segments: [
        {
          startOffset: 0,
          endOffset: "Contract".length,
          formatting: { color: "blue" },
        },
      ],
    });
  });

  test("arbitrary UTF-16 repartitioning cannot change presentation", () => {
    const symbol = fc.constantFrom(
      "A",
      "é",
      "e\u0301",
      "中",
      "😀",
      "\uD83D",
      "\uDE00",
    );
    fc.assert(
      fc.property(
        fc
          .array(symbol, { minLength: 1, maxLength: 40 })
          .map((parts) => parts.join("")),
        fc.array(fc.nat(), { maxLength: 30 }),
        (text, rawCuts) => {
          const cuts = [
            0,
            ...new Set(rawCuts.map((cut) => cut % (text.length + 1))),
            text.length,
          ].toSorted((left, right) => left - right);
          const richPresentation = {
            bold: true,
            italic: true,
            underline: true,
            strike: true,
            fontFamily: "Noto Sans",
            fontSizePt: 10.5,
            color: "#c00000",
            directFormatting: {
              bold: true,
              italic: true,
              underline: true,
              strike: true,
              fontFamily: "Noto Sans",
              fontSizePt: 10.5,
              color: "C00000",
            },
          } as const satisfies RunPresentation;
          const revisedRuns: FolioContentRun[] = [
            { text: "", color: "FFFFFF" },
          ];
          for (let index = 1; index < cuts.length; index++) {
            const start = cuts[index - 1];
            const end = cuts[index];
            if (start === undefined || end === undefined || start === end) {
              continue;
            }
            revisedRuns.push({
              text: text.slice(start, end),
              ...richPresentation,
            });
          }
          revisedRuns.push({ text: "", bold: false });
          const base = block(text, [{ text, ...richPresentation }]);
          const revised = block(text, revisedRuns);

          expect(sameCanonicalInlinePresentation(base, revised)).toBe(true);
          expect(
            inlineFormattingSegments({
              baseBlock: base,
              targetBlock: revised,
              maxSegments: 0,
            }),
          ).toEqual({ status: "compared", segments: [] });
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("formatting segments and their budget threshold survive UTF-16 repartitioning", () => {
    const prefix = "A😀e\u0301Z";
    const prefixUnits = Array.from(
      { length: prefix.length },
      (_unused, index) => prefix.slice(index, index + 1),
    );
    const presentationIndex = fc.integer({
      min: 0,
      max: PRESENTATION_STATES.length - 1,
    });
    const buildPartitionedRuns = (
      text: string,
      presentationIndexes: readonly number[],
      cuts: readonly number[],
      noisySeams: ReadonlySet<number>,
    ): FolioContentRun[] => {
      const boundaries = [...new Set([0, ...cuts, text.length])].toSorted(
        (left, right) => left - right,
      );
      const runs: FolioContentRun[] = [];
      for (let index = 1; index < boundaries.length; index++) {
        const start = boundaries[index - 1];
        const end = boundaries[index];
        if (start === undefined || end === undefined || start === end) continue;
        if (noisySeams.has(start)) runs.push({ text: "", color: "FFFFFF" });
        const presentation =
          PRESENTATION_STATES[presentationIndexes[start] ?? 0];
        runs.push({ text: text.slice(start, end), ...presentation });
      }
      runs.push({ text: "", bold: false });
      return runs;
    };
    fc.assert(
      fc.property(
        fc.array(fc.tuple(presentationIndex, presentationIndex), {
          maxLength: 40,
        }),
        fc.array(fc.nat(), { maxLength: 30 }),
        fc.array(fc.nat(), { maxLength: 30 }),
        (suffixPresentations, rawBaseCuts, rawRevisedCuts) => {
          const suffix = Array.from(
            { length: suffixPresentations.length },
            (_unused, index) => String.fromCharCode(0x61 + (index % 26)),
          );
          const text = [...prefixUnits, ...suffix].join("");
          const basePresentationIndexes = [
            ...Array.from({ length: prefix.length }, () => 0),
            ...suffixPresentations.map(([base]) => base),
          ];
          const revisedPresentationIndexes = [
            1,
            ...Array.from({ length: prefix.length - 1 }, () => 0),
            ...suffixPresentations.map(([_base, revised]) => revised),
          ];
          const mandatoryCuts = new Set([0, 1, text.length]);
          for (let index = 1; index < text.length; index++) {
            if (
              basePresentationIndexes[index] !==
                basePresentationIndexes[index - 1] ||
              revisedPresentationIndexes[index] !==
                revisedPresentationIndexes[index - 1]
            ) {
              mandatoryCuts.add(index);
            }
          }
          const baseCuts = [
            ...mandatoryCuts,
            2,
            ...rawBaseCuts.map((cut) => cut % (text.length + 1)),
          ];
          const revisedCuts = [
            ...mandatoryCuts,
            4,
            ...rawRevisedCuts.map((cut) => cut % (text.length + 1)),
          ];
          const atomicBase = block(
            text,
            buildPartitionedRuns(
              text,
              basePresentationIndexes,
              Array.from(
                { length: text.length + 1 },
                (_unused, index) => index,
              ),
              new Set(),
            ),
          );
          const atomicRevised = block(
            text,
            buildPartitionedRuns(
              text,
              revisedPresentationIndexes,
              Array.from(
                { length: text.length + 1 },
                (_unused, index) => index,
              ),
              new Set(),
            ),
          );
          const splitBase = block(
            text,
            buildPartitionedRuns(
              text,
              basePresentationIndexes,
              baseCuts,
              new Set([2]),
            ),
          );
          const splitRevised = block(
            text,
            buildPartitionedRuns(
              text,
              revisedPresentationIndexes,
              revisedCuts,
              new Set([4]),
            ),
          );
          const reference = inlineFormattingSegments({
            baseBlock: atomicBase,
            targetBlock: atomicRevised,
            maxSegments: Number.MAX_SAFE_INTEGER,
          });
          expect(reference.status).toBe("compared");
          if (reference.status !== "compared")
            throw new Error("The unbounded oracle did not compare.");
          expect(reference.segments.length).toBeGreaterThan(0);

          expect(
            sameCanonicalInlinePresentation(atomicBase, atomicRevised),
          ).toBe(false);
          expect(sameCanonicalInlinePresentation(splitBase, splitRevised)).toBe(
            false,
          );
          expect(
            inlineFormattingSegments({
              baseBlock: splitBase,
              targetBlock: splitRevised,
              maxSegments: Number.MAX_SAFE_INTEGER,
            }),
          ).toEqual(reference);
          for (const maxSegments of new Set([
            0,
            Math.max(0, reference.segments.length - 1),
            reference.segments.length,
            reference.segments.length + 1,
          ])) {
            expect(
              inlineFormattingSegments({
                baseBlock: splitBase,
                targetBlock: splitRevised,
                maxSegments,
              }),
            ).toEqual(
              maxSegments < reference.segments.length
                ? { status: "budget-exceeded", maximum: maxSegments }
                : reference,
            );
          }
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a malformed run stream cannot be accepted as an equal projection", () => {
    const valid = block("Contract", [{ text: "Contract", bold: true }]);
    const missingText = block("Contract", [{ text: "Contrac", bold: true }]);

    expect(sameCanonicalInlinePresentation(valid, missingText)).toBe(false);
    expect(
      inlineFormattingSegments({
        baseBlock: valid,
        targetBlock: missingText,
        maxSegments: 1,
      }),
    ).toEqual({ status: "unalignable", side: "revised" });
  });

  test("a valid early difference cannot escape a malformed trailing run stream", () => {
    const valid = block("Contract", [{ text: "Contract" }]);
    const malformedAfterDifference = block("Contract", [
      { text: "Con", bold: true },
      { text: "trac" },
    ]);

    expect(
      sameCanonicalInlinePresentation(valid, malformedAfterDifference),
    ).toBe(false);
    for (const maxSegments of [0, 1, 2]) {
      expect(
        inlineFormattingSegments({
          baseBlock: valid,
          targetBlock: malformedAfterDifference,
          maxSegments,
        }),
      ).toEqual({ status: "unalignable", side: "revised" });
    }
  });
});
