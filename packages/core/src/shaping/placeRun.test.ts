/**
 * What placement has to get right, in the scripts that need it.
 *
 * Each case asserts something a `cmap` lookup per code point cannot produce, so
 * a regression to unshaped output fails here rather than reaching a reader.
 */

import { describe, expect, test } from "bun:test";

import {
  readShapingTestFont,
  SHAPING_TEST_FONTS_INSTALLED,
  SHAPING_TEST_FONTS_SKIP_REASON,
  type ShapingTestScript,
} from "../fonts/sfnt/__tests__/testFonts";
import { toSfntBytes } from "../fonts/sfnt/woff";
import { advancesPerCodePointPx, directionOf, needsShaping, placeRun } from "./placeRun";
import { getShaper, SHAPING_DIRECTION, type Shaper } from "./shaper";

const FONT_SIZE_PX = 16;

/** beh twice: the first takes an initial form, the last a final one. */
const BEH_TWICE = "بب";
/** beh, yeh, teh: a word whose every letter joins. */
const ARABIC_WORD = "بيت";
/** lam then alef. */
const LAM_ALEF = "لا";
/** ka, virama, ssa: a Devanagari cluster that forms a conjunct. */
const DEVANAGARI_CLUSTER = "क्ष";
/** shin with its dot and a vowel point, both positioned on the letter. */
const HEBREW_POINTED = "שָׁ";

const faceOf = async (script: ShapingTestScript): Promise<Uint8Array> => {
  const woff = await readShapingTestFont(script);
  const sfnt = toSfntBytes(woff);
  if (sfnt.isErr()) {
    throw sfnt.error;
  }
  return sfnt.value;
};

const glyphIdsOf = (shaper: Shaper, font: Uint8Array, text: string): readonly number[] =>
  placeRun({ shaper, font, text, fontSizePx: FONT_SIZE_PX }).map(({ glyphId }) => glyphId);

describe("which script a run is in", () => {
  test("Latin, Cyrillic and Greek need no shaping", () => {
    expect(needsShaping("Smlouva")).toBe(false);
    expect(needsShaping("Договор")).toBe(false);
    expect(needsShaping("Σύμβαση")).toBe(false);
  });

  test("the scripts that shape are recognised as such", () => {
    expect(needsShaping(ARABIC_WORD)).toBe(true);
    expect(needsShaping(DEVANAGARI_CLUSTER)).toBe(true);
    expect(needsShaping(HEBREW_POINTED)).toBe(true);
  });

  test("only the bidirectional scripts run right to left", () => {
    expect(directionOf(ARABIC_WORD)).toBe(SHAPING_DIRECTION.rightToLeft);
    expect(directionOf(HEBREW_POINTED)).toBe(SHAPING_DIRECTION.rightToLeft);
    // Devanagari shapes, and reads left to right.
    expect(directionOf(DEVANAGARI_CLUSTER)).toBe(SHAPING_DIRECTION.leftToRight);
    expect(directionOf("Smlouva")).toBe(SHAPING_DIRECTION.leftToRight);
  });
});

describe.skipIf(!SHAPING_TEST_FONTS_INSTALLED)(
  `placement (${SHAPING_TEST_FONTS_SKIP_REASON})`,
  () => {
    test("a letter takes a different glyph from its position in a word", async () => {
      const shaper = await getShaper();
      const font = await faceOf("arabic");

      // The sharpest test of joining: one letter, twice. A per-code-point
      // lookup gives the same glyph both times.
      const pair = glyphIdsOf(shaper, font, BEH_TWICE);
      expect(new Set(pair).size).toBeGreaterThan(1);

      // And the form a letter takes alone is not the one it takes in a word.
      const alone = glyphIdsOf(shaper, font, "ل");
      const joined = glyphIdsOf(shaper, font, LAM_ALEF);
      expect(joined).not.toEqual(expect.arrayContaining([...alone]));
    });

    test("a right-to-left run comes back in visual order", async () => {
      const shaper = await getShaper();
      const placed = placeRun({
        shaper,
        font: await faceOf("arabic"),
        text: ARABIC_WORD,
        fontSizePx: FONT_SIZE_PX,
      });

      // Clusters index the source text, so a run painted left to right walks
      // them backwards: the first glyph on the page is the last character read.
      const clusters = placed.map(({ clusterIndex }) => clusterIndex);
      expect(clusters).toEqual([...clusters].sort((left, right) => right - left));
    });

    test("a Devanagari conjunct is one cluster over several characters", async () => {
      const shaper = await getShaper();
      const placed = placeRun({
        shaper,
        font: await faceOf("devanagari"),
        text: DEVANAGARI_CLUSTER,
        fontSizePx: FONT_SIZE_PX,
      });

      // Three characters, one cluster: the conjunct is a glyph no single
      // character maps to, and the whole cluster's width belongs to it.
      expect(new Set(placed.map(({ clusterIndex }) => clusterIndex)).size).toBe(1);
      const advances = advancesPerCodePointPx(placed, [...DEVANAGARI_CLUSTER].length);
      expect(advances).toHaveLength(3);
      expect(advances.at(0)).toBeGreaterThan(0);
      expect(advances.at(1)).toBe(0);
      expect(advances.at(2)).toBe(0);
    });

    test("a Hebrew point is placed on its letter rather than after it", async () => {
      const shaper = await getShaper();
      const placed = placeRun({
        shaper,
        font: await faceOf("hebrew"),
        text: HEBREW_POINTED,
        fontSizePx: FONT_SIZE_PX,
      });

      // A point advances the pen by nothing and sits where the shaper puts it,
      // which is what makes it a point rather than a following character.
      const marks = placed.filter(({ xAdvancePx }) => xAdvancePx === 0);
      expect(marks.length).toBeGreaterThan(0);
      expect(marks.some(({ xOffsetPx, yOffsetPx }) => xOffsetPx !== 0 || yOffsetPx !== 0)).toBe(
        true,
      );
    });

    test("advances sum to the run's own width whatever the clusters", async () => {
      const shaper = await getShaper();
      const placed = placeRun({
        shaper,
        font: await faceOf("arabic"),
        text: ARABIC_WORD,
        fontSizePx: FONT_SIZE_PX,
      });

      const shaped = placed.reduce((total, { xAdvancePx }) => total + xAdvancePx, 0);
      const perCodePoint = advancesPerCodePointPx(placed, [...ARABIC_WORD].length).reduce(
        (total, advance) => total + advance,
        0,
      );
      // Line breaking reads the per-code-point numbers, so they have to add up
      // to what the run actually occupies.
      expect(perCodePoint).toBeCloseTo(shaped, 10);
    });

    test("a run scales with its size", async () => {
      const shaper = await getShaper();
      const font = await faceOf("arabic");
      const options = { shaper, font, text: ARABIC_WORD } as const;

      const small = placeRun({ ...options, fontSizePx: 10 });
      const large = placeRun({ ...options, fontSizePx: 20 });

      const width = (glyphs: typeof small): number =>
        glyphs.reduce((total, { xAdvancePx }) => total + xAdvancePx, 0);
      expect(width(large)).toBeCloseTo(width(small) * 2, 10);
    });
  },
);
