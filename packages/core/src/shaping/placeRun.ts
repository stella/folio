/**
 * Placement: shaped glyphs in the units the rest of folio works in.
 *
 * The shaper answers in font design units and byte offsets. Both consumers want
 * the same answer in CSS pixels and code-point indices, and they must want it
 * from the same code: a measurer that scaled one way and a painter that scaled
 * another would lay out and paint two different pages.
 */

import { hasComplexScript, hasRightToLeft } from "../utils/scriptSegments";
import { SHAPING_DIRECTION, type Shaper, type ShapingDirection } from "./shaper";

/**
 * One glyph, placed. Advances and offsets are in CSS pixels at the run's size;
 * the glyph id belongs to the face that was shaped, not to any subset of it.
 */
export type PlacedGlyph = {
  readonly glyphId: number;
  /**
   * Index into the run's code points of the cluster this glyph belongs to.
   * Several glyphs share it when a cluster shapes to more than one glyph, and
   * one glyph carries it for every code point of a ligature.
   */
  readonly clusterIndex: number;
  readonly xAdvancePx: number;
  readonly xOffsetPx: number;
  readonly yOffsetPx: number;
};

export type PlaceRunOptions = {
  readonly shaper: Shaper;
  /** sfnt bytes of the binary that covers this text. */
  readonly font: Uint8Array;
  readonly text: string;
  readonly fontSizePx: number;
  /** Omitted means: from the text's own script. */
  readonly direction?: ShapingDirection;
};

/** Whether a run's glyphs can be chosen from its code points alone. */
export const needsShaping = (text: string): boolean => hasComplexScript(text);

export const directionOf = (text: string): ShapingDirection =>
  hasRightToLeft(text) ? SHAPING_DIRECTION.rightToLeft : SHAPING_DIRECTION.leftToRight;

/**
 * Code-point index for each UTF-8 byte offset the shaper can report as a
 * cluster. Only cluster starts are ever looked up, so offsets inside a
 * character are left absent rather than filled in with a neighbour's index.
 */
const codePointIndexByByteOffset = (text: string): ReadonlyMap<number, number> => {
  const byOffset = new Map<number, number>();
  let offset = 0;
  let index = 0;
  for (const character of text) {
    byOffset.set(offset, index);
    // SAFETY: iterating a string yields whole code points.
    const codePoint = character.codePointAt(0)!;
    offset += utf8Length(codePoint);
    index += 1;
  }
  return byOffset;
};

const UTF8_TWO_BYTE_MINIMUM = 0x80;
const UTF8_THREE_BYTE_MINIMUM = 0x800;
const UTF8_FOUR_BYTE_MINIMUM = 0x1_00_00;

const utf8Length = (codePoint: number): number => {
  if (codePoint < UTF8_TWO_BYTE_MINIMUM) {
    return 1;
  }
  if (codePoint < UTF8_THREE_BYTE_MINIMUM) {
    return 2;
  }
  return codePoint < UTF8_FOUR_BYTE_MINIMUM ? 3 : 4;
};

/**
 * Shape and place one run against one binary.
 *
 * The text must be served by the face's bytes throughout: a run split across
 * subsets of a family is placed one segment at a time, because a glyph id only
 * means anything in the binary it came from.
 */
export const placeRun = ({
  shaper,
  font,
  text,
  fontSizePx,
  direction,
}: PlaceRunOptions): readonly PlacedGlyph[] => {
  const shaped = shaper.shapeRun({
    font,
    text,
    direction: direction ?? directionOf(text),
  });
  if (shaped.unitsPerEm <= 0) {
    return [];
  }
  const scale = fontSizePx / shaped.unitsPerEm;
  const indexByOffset = codePointIndexByByteOffset(text);
  return shaped.glyphs.map(({ glyphId, cluster, xAdvance, xOffset, yOffset }) => ({
    glyphId,
    clusterIndex: indexByOffset.get(cluster) ?? 0,
    xAdvancePx: xAdvance * scale,
    xOffsetPx: xOffset * scale,
    yOffsetPx: yOffset * scale,
  }));
};

/**
 * The run's advance spread over its code points, one number each.
 *
 * A cluster's whole advance sits on the first code point that produced it and
 * the rest carry zero. Shaping is what decides how many glyphs a cluster
 * becomes, so there is no per-code-point advance to report: two code points
 * that ligated into one glyph have one width between them, and splitting it
 * would invent a boundary the font does not have. The sum is exact, which is
 * what line breaking and every caller downstream reads.
 */
export const advancesPerCodePointPx = (
  glyphs: readonly PlacedGlyph[],
  codePointCount: number,
): readonly number[] => {
  const advances = Array.from({ length: codePointCount }, () => 0);
  for (const { clusterIndex, xAdvancePx } of glyphs) {
    const at = Math.min(clusterIndex, codePointCount - 1);
    if (at < 0) {
      continue;
    }
    advances[at] = (advances[at] ?? 0) + xAdvancePx;
  }
  return advances;
};
