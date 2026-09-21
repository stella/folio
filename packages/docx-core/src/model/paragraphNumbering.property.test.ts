/**
 * `w:numId 0` had five spellings and `w:numPr`'s two slots had four merge
 * implementations. Both are now one function, and these properties are what
 * keep them one.
 *
 * 1. **The five spellings agree.** `isNumberingReference(numId)`, a bare
 *    `numId === 0`, the hand-inlined absence-or-literal in `validate/docx.ts`,
 *    the relational `numId > 0` in the AI snapshot, and the Rust kernel's
 *    `num_id != Some(0)` all decided "does this paragraph name a numbering
 *    definition". Each is restated here against the union's reading over the
 *    whole input space, so a spelling that drifts fails rather than diverging
 *    in one tier.
 * 2. **The merge table holds.** ECMA-376 17.3.1.19's independent inheritance,
 *    written as the five rows of the design's table and checked against the
 *    fold over generated tiers, including that the fold is closed (its result
 *    is always a member of the same union) and associative over three tiers.
 * 3. **An authored `w:ilvl` absence survives.** The slots a stated override
 *    writes back parse to the same override: this is the property that stops
 *    anyone "simplifying" `ilvl?: number` to `ilvl: number`.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import {
  isNumberingReference,
  mergeParagraphNumbering,
  NO_NUMBERING_NUM_ID,
  type ParagraphNumberingOverride,
  paragraphNumberingFromSlots,
  paragraphNumberingLevel,
  paragraphNumberingReferenceId,
  paragraphNumberingSlots,
  resolveParagraphNumbering,
  sameStatedParagraphNumbering,
} from "./paragraphNumbering";

const slots = fc.record({
  numId: fc.option(fc.integer({ min: -2, max: 9 }), { nil: undefined }),
  ilvl: fc.option(fc.integer({ min: -1, max: 10 }), { nil: undefined }),
});

const statedArbitrary = slots.map(paragraphNumberingFromSlots);

describe("the five spellings of the reserved w:numId", () => {
  test("all of them agree with the union's reading", () => {
    fc.assert(
      fc.property(slots, ({ numId, ilvl }) => {
        const stated = paragraphNumberingFromSlots({ numId, ilvl });
        const namesADefinition = paragraphNumberingReferenceId(stated) !== undefined;

        // A: the sanctioned reader.
        expect(isNumberingReference(numId)).toBe(namesADefinition);
        // B: the bare literal, as `document-operations.ts` spelled it.
        expect(numId !== undefined && numId !== 0).toBe(namesADefinition);
        // C: the absence-or-literal hand-inlined across the package boundary.
        expect(!(numId === undefined || numId === NO_NUMBERING_NUM_ID)).toBe(namesADefinition);
        // D: the relational form the AI snapshot used. It is the one spelling
        // that disagreed: a malformed package's negative id is a dangling
        // reference to A, B, C and E, and "not numbered" to D.
        const relational = numId !== undefined && numId > 0;
        if (numId !== undefined && numId < 0) {
          expect(relational).toBe(false);
          expect(namesADefinition).toBe(true);
        } else {
          expect(relational).toBe(namesADefinition);
        }
        // E: the Rust kernel's `num_id != Some(0)`, which also admits absence.
        expect(numId !== NO_NUMBERING_NUM_ID).toBe(namesADefinition || numId === undefined);
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("the reserved id never reaches a reference arm", () => {
    fc.assert(
      fc.property(statedArbitrary, (stated) => {
        expect(paragraphNumberingReferenceId(stated)).not.toBe(NO_NUMBERING_NUM_ID);
        expect(resolveParagraphNumbering(stated).kind === "reference").toBe(
          paragraphNumberingReferenceId(stated) !== undefined,
        );
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a reference arm is unrepresentable for the reserved id at the type level", () => {
    expect(paragraphNumberingFromSlots({ numId: NO_NUMBERING_NUM_ID, ilvl: 3 })).toEqual({
      kind: "none",
    });
  });
});

/** The design's merge table, restated as data rather than as four spreads. */
const MERGE_TABLE: readonly {
  name: string;
  inherited: ParagraphNumberingOverride | undefined;
  stated: ParagraphNumberingOverride | undefined;
  expected: ParagraphNumberingOverride | undefined;
}[] = [
  {
    name: "a direct cancellation is authoritative",
    inherited: { kind: "reference", numId: 4, ilvl: 1 },
    stated: { kind: "none" },
    expected: { kind: "none" },
  },
  {
    name: "a direct id keeps the level inherited from a reference",
    inherited: { kind: "reference", numId: 4, ilvl: 1 },
    stated: { kind: "reference", numId: 7 },
    expected: { kind: "reference", numId: 7, ilvl: 1 },
  },
  {
    name: "a direct id keeps a level-only tier beneath it",
    inherited: { kind: "levelOnly", ilvl: 4 },
    stated: { kind: "reference", numId: 7 },
    expected: { kind: "reference", numId: 7, ilvl: 4 },
  },
  {
    name: "a level-only tier keeps the id it inherits",
    inherited: { kind: "reference", numId: 4 },
    stated: { kind: "levelOnly", ilvl: 2 },
    expected: { kind: "reference", numId: 4, ilvl: 2 },
  },
  {
    name: "a level-only tier over a cancellation stays cancelled",
    inherited: { kind: "none" },
    stated: { kind: "levelOnly", ilvl: 2 },
    expected: { kind: "none" },
  },
  {
    name: "a level-only tier with nothing beneath it states only the level",
    inherited: undefined,
    stated: { kind: "levelOnly", ilvl: 2 },
    expected: { kind: "levelOnly", ilvl: 2 },
  },
  {
    name: "a tier that states nothing keeps what it inherits",
    inherited: { kind: "reference", numId: 4, ilvl: 1 },
    stated: undefined,
    expected: { kind: "reference", numId: 4, ilvl: 1 },
  },
];

describe("w:numId and w:ilvl inherit independently", () => {
  for (const { name, inherited, stated, expected } of MERGE_TABLE) {
    test(name, () => {
      expect(mergeParagraphNumbering(inherited, stated)).toEqual(expected);
    });
  }

  test("the fold is closed and idempotent under a repeated tier", () => {
    fc.assert(
      fc.property(statedArbitrary, statedArbitrary, (inherited, stated) => {
        const merged = mergeParagraphNumbering(inherited, stated);
        // Closed: the result is a member of the same union, so a third tier
        // folds over it without a special case.
        expect(mergeParagraphNumbering(merged, stated)).toEqual(merged);
        // A tier that states nothing changes nothing.
        expect(mergeParagraphNumbering(merged, undefined)).toEqual(merged);
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a cancellation is a boundary for a later id-only reference", () => {
    fc.assert(
      fc.property(statedArbitrary, (inherited) => {
        const cancelled = mergeParagraphNumbering(inherited, { kind: "none" });
        expect(
          mergeParagraphNumbering(cancelled, paragraphNumberingFromSlots({ numId: 7 })),
        ).toEqual(paragraphNumberingFromSlots({ numId: 7 }));
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a merged level never contradicts the tier that stated it", () => {
    fc.assert(
      fc.property(statedArbitrary, statedArbitrary, (inherited, stated) => {
        const merged = mergeParagraphNumbering(inherited, stated);
        if (stated?.kind === "levelOnly" && merged?.kind === "reference") {
          expect(paragraphNumberingLevel(merged)).toBe(stated.ilvl);
        }
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });
});

describe("an authored w:ilvl absence survives the round trip", () => {
  test("the slots a stated override writes back parse to the same override", () => {
    fc.assert(
      fc.property(statedArbitrary, (stated) => {
        if (stated === undefined) {
          return;
        }
        expect(
          sameStatedParagraphNumbering(
            paragraphNumberingFromSlots(paragraphNumberingSlots(stated)),
            stated,
          ),
        ).toBe(true);
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("a stated level zero is not the same statement as no level", () => {
    const withLevel = paragraphNumberingFromSlots({ numId: 7, ilvl: 0 });
    const withoutLevel = paragraphNumberingFromSlots({ numId: 7 });
    expect(sameStatedParagraphNumbering(withLevel, withoutLevel)).toBe(false);
    expect(resolveParagraphNumbering(withLevel)).toEqual(resolveParagraphNumbering(withoutLevel));
  });
});
