/**
 * A `MeasureProvider` that needs no canvas.
 *
 * The layout engine measures through a swappable seam so a server, a CI job or
 * an export can lay out pages without a browser. This backend supplies the
 * seam from parsed font binaries: advances come from `hmtx`, ascent and
 * descent from the ink bounds of `H` and `g` (which is what the canvas backend
 * reads out of `measureText("Hg")`), and the single-line ratio from the same
 * static table the canvas backend uses, so that number cannot differ at all.
 *
 * The arithmetic on top of those advances is shared with the canvas backend
 * through `advanceComposition.ts`. What remains free to differ is the raw
 * advance of a glyph: a canvas applies kerning and ligatures, `hmtx` does not.
 * That residue is measured rather than assumed, by the differential tests
 * beside this file and by the paint-equivalence harness.
 */

import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import {
  composeRunMeasurement,
  composeTextWidth,
  type GlyphAdvancesFn,
} from "../layout-engine/measure/advanceComposition";
import {
  DEFAULT_FONT_FAMILY,
  DEFAULT_FONT_SIZE,
  getResolvedData,
  ptToPx,
} from "../layout-engine/measure/measureHelpers";
import type { MeasureProvider } from "../layout-engine/measure/measureProvider";
import { setMeasureProvider } from "../layout-engine/measure/measureProvider";
import type {
  FontMetrics,
  FontStyle,
  RunMeasurement,
  TextMeasurement,
} from "../layout-engine/measure/measureTypes";
import {
  advancesPerCodePointPx,
  needsShaping,
  placeRun,
} from "../shaping/placeRun";
import { resolvedShaper, type Shaper } from "../shaping/shaper";
import { parseSfnt } from "./sfnt/parse";
import type { SfntFont } from "./sfnt/parse";
import { toSfntBytes } from "./sfnt/woff";

/** One concrete face, as a font source is asked for it. */
export type HeadlessFontRequest = {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
};

/**
 * Where face binaries come from. A server reads them from disk, a browser
 * fetches them, a package supplies its own embedded faces; none of that is
 * this module's business.
 *
 * A face is a *list* of binaries rather than one, because a family is
 * routinely shipped split by script: `@fontsource` cuts every family into
 * disjoint `latin`, `latin-ext`, `cyrillic` and `greek` subsets, and a Czech
 * or Polish document needs two of them at once. Returning one binary per face
 * would mean picking a subset per document and painting `.notdef` for
 * everything outside it. The list is in priority order, and a code point is
 * served by the first binary whose `cmap` covers it.
 */
export type HeadlessFontSource = {
  /** Plain sfnt or WOFF binaries, empty when the source has no such face. */
  readonly load: (request: HeadlessFontRequest) => readonly Uint8Array[];
};

/** A face the source could not supply, measured with a stand-in instead. */
export type HeadlessFontSubstitution = {
  readonly requested: HeadlessFontRequest;
  readonly reason: string;
};

const HeadlessMeasureErrorBase: TaggedErrorClass<"HeadlessMeasureError"> =
  TaggedError("HeadlessMeasureError");

export class HeadlessMeasureError extends HeadlessMeasureErrorBase<{
  message: string;
}> {}

export type HeadlessMeasureOptions = {
  /**
   * The shaper to measure shaping scripts with.
   *
   * Omitted, the provider uses whichever shaper has been loaded, which is what
   * the layout entry points resolve once they find a document that needs one.
   * Pass one explicitly to measure against a shaper of your own; pass `null` to
   * measure as if none existed.
   */
  readonly shaper?: Shaper | null;
};

export type HeadlessMeasureProvider = {
  readonly provider: MeasureProvider;
  /**
   * Faces measured with a stand-in, in first-request order. A caller that
   * ships the result to a user must surface these: a substituted face
   * measures differently, so the pagination is not the authored one.
   */
  readonly substitutions: () => readonly HeadlessFontSubstitution[];
};

/**
 * Ink-bounds probe characters. The canvas backend takes its ink ascent and
 * descent from `measureText("Hg")`, so this backend reads those same two
 * glyphs. The font's own hhea box is wider and is reported separately, as
 * `fontBoxAscent`/`fontBoxDescent`; substituting one for the other shifts
 * every baseline.
 */
const ASCENT_PROBE = "H";
const DESCENT_PROBE = "g";

/** Ratios the canvas backend falls back to when it cannot measure ink bounds. */
const FALLBACK_ASCENT_RATIO = 0.8;
const FALLBACK_DESCENT_RATIO = 0.2;

const faceKey = ({ family, bold, italic }: HeadlessFontRequest): string =>
  `${family}\u0000${bold ? "b" : ""}${italic ? "i" : ""}`;

const requestFor = (style: FontStyle): HeadlessFontRequest => ({
  family: style.fontFamily ?? DEFAULT_FONT_FAMILY,
  bold: style.bold ?? false,
  italic: style.italic ?? false,
});

/**
 * The parsed binaries of one face, in the source's priority order.
 *
 * `primary` is the first that parsed and is where face-wide metrics come from:
 * ascent, descent and the font box are properties of the design, and the
 * subsets of one family share them. `covering` picks the binary that actually
 * has a glyph for a code point.
 */
type LoadedFace = {
  readonly fonts: readonly SfntFont[];
  readonly primary: SfntFont | null;
};

const coveringFont = (face: LoadedFace, codePoint: number): SfntFont | null => {
  for (const font of face.fonts) {
    if (font.glyphIdFor(codePoint) !== 0) {
      return font;
    }
  }
  return null;
};

/**
 * Maximal stretches of text served by one binary.
 *
 * A family split into `latin` and `latin-ext` puts a Czech `ř` and its
 * neighbouring `p` in different binaries, and a shaper works on one face at a
 * time: a glyph id only means anything in the binary it came from.
 */
type FaceSegment = { readonly font: SfntFont | null; readonly text: string };

const segmentByBinary = (
  face: LoadedFace,
  text: string,
): readonly FaceSegment[] => {
  const segments: FaceSegment[] = [];
  let current: { font: SfntFont | null; text: string } | null = null;
  for (const character of text) {
    // SAFETY: iterating a string yields whole code points.
    const font = coveringFont(face, character.codePointAt(0)!) ?? face.primary;
    if (current === null || current.font !== font) {
      current = { font, text: character };
      segments.push(current);
      continue;
    }
    current.text += character;
  }
  return segments;
};

/**
 * Resolve, decode, parse and cache faces. Kept separate from the provider so
 * the cache is per-provider: two providers over different sources must not
 * share a face table.
 */
const createFaceCache = (source: HeadlessFontSource) => {
  const faces = new Map<string, LoadedFace>();
  const substitutions: HeadlessFontSubstitution[] = [];

  const record = (request: HeadlessFontRequest, reason: string): void => {
    substitutions.push({ requested: request, reason });
  };

  const decode = (
    request: HeadlessFontRequest,
    bytes: Uint8Array,
  ): SfntFont | null => {
    const sfnt = toSfntBytes(bytes);
    if (sfnt.isErr()) {
      record(request, sfnt.error.message);
      return null;
    }
    const parsed = parseSfnt(sfnt.value);
    if (parsed.isErr()) {
      record(request, parsed.error.message);
      return null;
    }
    return parsed.value;
  };

  const load = (request: HeadlessFontRequest): LoadedFace => {
    const key = faceKey(request);
    const cached = faces.get(key);
    if (cached !== undefined) {
      return cached;
    }
    const binaries = source.load(request);
    if (binaries.length === 0) {
      record(request, "The font source has no binary for this face.");
    }
    const fonts = binaries.flatMap((bytes) => {
      const font = decode(request, bytes);
      return font === null ? [] : [font];
    });
    const loaded = { fonts, primary: fonts.at(0) ?? null };
    faces.set(key, loaded);
    return loaded;
  };

  const readSubstitutions = (): readonly HeadlessFontSubstitution[] =>
    substitutions;

  return { load, substitutions: readSubstitutions };
};

type FaceCache = ReturnType<typeof createFaceCache>;

/**
 * Advance of a string in px, summed per code point.
 *
 * Each code point is measured in the subset that actually covers it, because
 * a family split into `latin` and `latin-ext` puts a Czech `ř` and its
 * neighbouring `p` in different binaries, and measuring both against one of
 * them charges `.notdef` for whichever it lacks. A code point no subset covers
 * falls back to the primary's `.notdef` advance: that is the glyph a renderer
 * paints, so measuring it as anything else lays out a page nobody sees.
 */
const advanceOfText = (
  face: LoadedFace,
  text: string,
  fontSizePx: number,
  shaper: Shaper | null,
): number => {
  if (shaper !== null && needsShaping(text)) {
    return shapedAdvanceOfText(face, text, fontSizePx, shaper);
  }
  let total = 0;
  for (const char of text) {
    // SAFETY: iterating a string yields whole code points.
    const cp = char.codePointAt(0)!;
    const font = coveringFont(face, cp) ?? face.primary;
    if (font === null) {
      // No binary at all. Half the em is roughly a digit's width in a text
      // face, enough to keep a page from collapsing, and the substitution is
      // reported so nobody trusts the number.
      total += fontSizePx * 0.5;
      continue;
    }
    total +=
      (font.advanceWidthFor(font.glyphIdFor(cp)) * fontSizePx) /
      font.unitsPerEm;
  }
  return total;
};

/**
 * The advance of text that shapes, summed from the glyphs shaping produced.
 *
 * Summing `hmtx` per code point is the wrong number here and not by a rounding:
 * a lam-alef pair is one glyph with one advance, an Arabic letter's width
 * depends on the form its neighbours give it, and a Devanagari cluster reorders
 * before it is measured at all.
 */
const shapedAdvanceOfText = (
  face: LoadedFace,
  text: string,
  fontSizePx: number,
  shaper: Shaper,
): number => {
  let total = 0;
  for (const segment of segmentByBinary(face, text)) {
    if (segment.font === null) {
      // No binary at all: the same half-em stand-in the unshaped path uses,
      // and the substitution is already reported.
      total += fontSizePx * 0.5 * [...segment.text].length;
      continue;
    }
    if (!needsShaping(segment.text)) {
      total += advanceOfText(
        { fonts: [segment.font], primary: segment.font },
        segment.text,
        fontSizePx,
        null,
      );
      continue;
    }
    for (const glyph of placeRun({
      shaper,
      font: segment.font.bytes,
      text: segment.text,
      fontSizePx,
    })) {
      total += glyph.xAdvancePx;
    }
  }
  return total;
};

/**
 * Advances for shaped text, one per code point of the input.
 *
 * A cluster's advance sits on the code point that opens it; see
 * `advancesPerCodePointPx`. Segments the run does not need shaping for keep
 * their per-code-point widths, so a run of Arabic with a Latin word in it is
 * right on both sides.
 */
const advancesOfText = (
  face: LoadedFace,
  text: string,
  fontSizePx: number,
  shaper: Shaper,
): readonly number[] => {
  const advances: number[] = [];
  for (const segment of segmentByBinary(face, text)) {
    const characters = [...segment.text];
    if (segment.font === null || !needsShaping(segment.text)) {
      const single = {
        fonts: segment.font === null ? [] : [segment.font],
        primary: segment.font,
      };
      for (const character of characters) {
        advances.push(advanceOfText(single, character, fontSizePx, null));
      }
      continue;
    }
    const placed = placeRun({
      shaper,
      font: segment.font.bytes,
      text: segment.text,
      fontSizePx,
    });
    advances.push(...advancesPerCodePointPx(placed, characters.length));
  }
  return advances;
};

const createAdvancesOf =
  (faces: FaceCache, shaperNow: () => Shaper | null): GlyphAdvancesFn =>
  (text: string, style: FontStyle): readonly number[] => {
    const shaper = shaperNow();
    const face = faces.load(requestFor(style));
    const fontSizePx = ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE);
    if (shaper === null) {
      return [...text].map((character) =>
        advanceOfText(face, character, fontSizePx, null),
      );
    }
    return advancesOfText(face, text, fontSizePx, shaper);
  };

const createAdvanceOf =
  (faces: FaceCache, shaperNow: () => Shaper | null) =>
  (text: string, style: FontStyle): number =>
    advanceOfText(
      faces.load(requestFor(style)),
      text,
      ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE),
      shaperNow(),
    );

const probeInk = (
  font: SfntFont,
  probe: string,
): { yMin: number; yMax: number } | null => {
  const cp = probe.codePointAt(0);
  if (cp === undefined) {
    return null;
  }
  const bounds = font.glyphBoundsFor(font.glyphIdFor(cp));
  return bounds === null ? null : { yMin: bounds.yMin, yMax: bounds.yMax };
};

const createGetFontMetrics =
  (faces: FaceCache) =>
  (style: FontStyle): FontMetrics => {
    const fontSize = style.fontSize ?? DEFAULT_FONT_SIZE;
    const fontFamily = style.fontFamily ?? DEFAULT_FONT_FAMILY;
    const fontSizePx = ptToPx(fontSize);
    const loaded = faces.load(requestFor(style));

    let ascent = fontSizePx * FALLBACK_ASCENT_RATIO;
    let descent = fontSizePx * FALLBACK_DESCENT_RATIO;
    let fontBoxAscent = ascent;
    let fontBoxDescent = descent;
    // Face-wide metrics come from the primary binary: ascent, descent and the
    // font box are properties of the design, and the subsets of one family
    // share them, so reading whichever subset happened to cover a code point
    // would make a line's height depend on its script.
    const primary = loaded.primary;
    if (primary !== null) {
      const scale = fontSizePx / primary.unitsPerEm;
      const ascentInk = probeInk(primary, ASCENT_PROBE);
      const descentInk = probeInk(primary, DESCENT_PROBE);
      if (ascentInk !== null) {
        ascent = ascentInk.yMax * scale;
      }
      if (descentInk !== null) {
        descent = -descentInk.yMin * scale;
      }
      // A browser's `fontBoundingBox*` comes from hhea, so read hhea rather
      // than OS/2: the two disagree in many faces and a DOM painter positions
      // its text boxes against whichever one the browser used.
      fontBoxAscent = primary.ascender * scale;
      fontBoxDescent = -primary.descender * scale;
    }

    return {
      fontSize,
      ascent,
      descent,
      fontBoxAscent,
      fontBoxDescent,
      lineHeight: Math.max(fontSizePx, ascent + descent),
      fontFamily,
      singleLineRatio: getResolvedData(fontFamily, style).singleLineRatio,
    };
  };

/**
 * Build a provider over a font source. The provider is stateful only in its
 * face cache and its substitution log, both scoped to the returned object, so
 * two providers never contaminate each other.
 */
export const createHeadlessMeasureProvider = (
  source: HeadlessFontSource,
  options: HeadlessMeasureOptions = {},
): HeadlessMeasureProvider => {
  const faces = createFaceCache(source);
  // Read per measurement rather than captured: a provider is installed before
  // the document has been read, and whether that document needs a shaper is
  // not known until it has been.
  const shaperNow = (): Shaper | null =>
    options.shaper === undefined ? resolvedShaper() : options.shaper;
  const advanceOf = createAdvanceOf(faces, shaperNow);
  const advancesOf = createAdvancesOf(faces, shaperNow);
  const getFontMetrics = createGetFontMetrics(faces);

  const measureTextWidth = (text: string, style: FontStyle): number =>
    composeTextWidth({ text, style, advanceOf });

  const measureText = (text: string, style: FontStyle): TextMeasurement => {
    const metrics = getFontMetrics(style);
    return {
      width: measureTextWidth(text, style),
      height: metrics.ascent + metrics.descent,
      ascent: metrics.ascent,
      descent: metrics.descent,
    };
  };

  const measureRun = (text: string, style: FontStyle): RunMeasurement =>
    composeRunMeasurement({
      text,
      style,
      metrics: getFontMetrics(style),
      advanceOf,
      advancesOf,
    });

  return {
    provider: { getFontMetrics, measureTextWidth, measureText, measureRun },
    substitutions: faces.substitutions,
  };
};

/**
 * Install a headless backend as the active provider. An explicit call, never a
 * module side effect, so importing this module cannot silently change how a
 * document measures.
 */
export const installHeadlessMeasureProvider = (
  source: HeadlessFontSource,
  options: HeadlessMeasureOptions = {},
): HeadlessMeasureProvider => {
  const headless = createHeadlessMeasureProvider(source, options);
  setMeasureProvider(headless.provider);
  return headless;
};

/**
 * Refuse to measure when a source cannot supply a face that the document
 * actually uses. Callers that would rather fail than paginate against a
 * stand-in call this after layout.
 */
export const assertNoSubstitutions = (
  headless: HeadlessMeasureProvider,
): Result<void, HeadlessMeasureError> => {
  const substitutions = headless.substitutions();
  if (substitutions.length === 0) {
    return Result.ok(undefined);
  }
  const listed = substitutions
    .map(({ requested, reason }) => `${requested.family} (${reason})`)
    .join(", ");
  return Result.err(
    new HeadlessMeasureError({
      message: `The font source could not supply ${substitutions.length} face(s): ${listed}`,
    }),
  );
};
