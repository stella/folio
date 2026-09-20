/**
 * `w:outlineLvl` 9 is not a tenth heading level, and after the union it cannot
 * be read as one anywhere.
 *
 * Three properties over the whole stated range, plus the values outside it that
 * real packages carry:
 *
 * 1. **Fidelity.** Parse, force a real save (the captured markup dropped so the
 *    serializer actually runs), parse again: what the document stated comes
 *    back, and what it did not state stays unstated.
 * 2. **The sentinel is never a level.** No parsed value produces a heading arm
 *    at all for `w:val="9"`, and no heading arm anywhere carries a level above
 *    eight. `typecheck/model/outlineLevel.typecheck.ts` in `@stll/docx-core`
 *    pins the same claim at the type level, which is the half a runtime
 *    property cannot reach.
 * 3. **Out of range is dropped, not stored.** The Rust kernel refuses
 *    `w:val="10"` outright; the TypeScript parse boundary drops it so a
 *    document Word opens still opens, and the union is what makes "drops" mean
 *    "cannot be stored" rather than "nobody looks".
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  BODY_TEXT_OUTLINE_LEVEL,
  headingLevelOf,
  headingOutlineLevel,
  type OutlineLevel,
  outlineLevelFromStatedValue,
  outlineLevelStatedValue,
} from "@stll/docx-core/model";

import { propertyConfig } from "../../../../test/property-testing";

import { modelParagraphFormattingEmission } from "../internal/paragraphFormattingSerialization";
import { parseParagraphProperties } from "./paragraphProperties";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** The stated range, the sentinel, and what producers write outside both. */
const STATED_VALUES = [
  "0",
  "1",
  "2",
  "3",
  "4",
  "5",
  "6",
  "7",
  "8",
  "9",
  "10",
  "-1",
  "-3",
  "99",
  "",
  "one",
] as const;

const pPrXml = (stated: string | undefined): string =>
  `<w:pPr ${NS}>${stated === undefined ? "" : `<w:outlineLvl w:val="${stated}"/>`}</w:pPr>`;

const parsePPr = (xml: string): OutlineLevel | undefined => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("w:pPr did not parse");
  }
  return parseParagraphProperties(root, null)?.outlineLevel;
};

/**
 * Parse, serialize the modeled part, parse again. `w:pPr` has no captured
 * markup to drop here, because `modelParagraphFormattingEmission` is the
 * rebuild path itself.
 */
const throughRealSave = (stated: string | undefined): OutlineLevel | undefined => {
  const parsed = parsePPr(pPrXml(stated));
  const written = modelParagraphFormattingEmission(
    parsed === undefined ? {} : { outlineLevel: parsed },
  ).propertiesXml;
  return parsePPr(`<w:pPr ${NS}>${written ?? ""}</w:pPr>`);
};

/** What the source stated, read off the generated case rather than the parser. */
const expectedOutlineLevel = (stated: string | undefined): OutlineLevel | undefined => {
  if (stated === undefined) {
    return undefined;
  }
  const value = Number(stated);
  if (stated.trim() === "" || !Number.isInteger(value)) {
    return undefined;
  }
  return value === 9 ? BODY_TEXT_OUTLINE_LEVEL : headingOutlineLevel(value);
};

const statedArbitrary = fc.option(fc.constantFrom(...STATED_VALUES), { nil: undefined });

describe("a stated outline level survives a save", () => {
  test("what the document stated comes back, and what it did not stays unstated", () => {
    fc.assert(
      fc.property(statedArbitrary, (stated) => {
        const expected = expectedOutlineLevel(stated);
        expect(parsePPr(pPrXml(stated))).toEqual(expected);
        expect(throughRealSave(stated)).toEqual(expected);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("a save is a fixed point: a second round trip changes nothing", () => {
    fc.assert(
      fc.property(statedArbitrary, (stated) => {
        const once = throughRealSave(stated);
        const twice = throughRealSave(
          once === undefined ? undefined : String(outlineLevelStatedValue(once)),
        );
        expect(twice).toEqual(once);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});

describe("the body-text sentinel is never a heading level", () => {
  test("no parsed outline level names a heading above eight", () => {
    fc.assert(
      fc.property(statedArbitrary, (stated) => {
        const level = headingLevelOf(parsePPr(pPrXml(stated)));
        expect(level === undefined || level <= 8).toBe(true);
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test('`w:val="9"` is body text, and body text writes nine back', () => {
    expect(parsePPr(pPrXml("9"))).toEqual(BODY_TEXT_OUTLINE_LEVEL);
    expect(headingLevelOf(BODY_TEXT_OUTLINE_LEVEL)).toBeUndefined();
    expect(outlineLevelStatedValue(BODY_TEXT_OUTLINE_LEVEL)).toBe(9);
  });

  test("a value outside the stated range is dropped rather than stored", () => {
    for (const stated of ["10", "-1", "99", "one", ""]) {
      expect(parsePPr(pPrXml(stated))).toBeUndefined();
    }
  });

  test("every stated value round-trips through the one reader and the one writer", () => {
    for (let value = 0; value <= 9; value += 1) {
      const level = outlineLevelFromStatedValue(value);
      expect(level).toBeDefined();
      expect(level === undefined ? undefined : outlineLevelStatedValue(level)).toBe(value);
    }
    expect(outlineLevelFromStatedValue(10)).toBeUndefined();
    expect(outlineLevelFromStatedValue(-1)).toBeUndefined();
    expect(outlineLevelFromStatedValue(1.5)).toBeUndefined();
  });
});
