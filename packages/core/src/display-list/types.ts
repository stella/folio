/**
 * The display list: folio's painter-neutral paint IR.
 *
 * This is seam 4 of `docs/seam-architecture.md` made explicit. The layout
 * engine decides *where* things go; this module states *what is drawn* in
 * terms no backend can look behind. A DOM backend and a PDF backend both
 * consume it, so the two cannot drift: a backend that wants a fact the
 * display list does not carry has to add it here, where both backends see it.
 *
 * ## Invariants
 *
 * - Pure data. No behaviour, no DOM, no font binaries, no layout types. The
 *   whole structure survives `JSON.parse(JSON.stringify(...))`.
 * - Units are CSS pixels at 96 dpi, matching the layout engine
 *   (`measureHelpers.ts`). Backends convert at emit time; a PDF backend
 *   multiplies by 0.75 for points.
 * - Coordinates are page-absolute with the origin at the page's top-left and
 *   y growing downwards, again matching the layout engine. A backend whose
 *   device space differs (PDF's origin is bottom-left) flips once, at its own
 *   edge.
 * - `primitives` is painted in order, back to front. There is no z-index:
 *   the producer has already resolved stacking, because two backends sorting
 *   independently is exactly the divergence this module exists to prevent.
 * - Colours are resolved. CSS custom properties (`--doc-canvas-text`) and
 *   `currentColor` are producer concerns; a backend never resolves a colour.
 *
 * ## What is deliberately absent
 *
 * Glyph ids. The display list carries code points and the *advances layout
 * actually used*, so a backend positions text on the measurer's numbers
 * rather than on its own idea of the font's metrics. Mapping code points to
 * glyphs needs the font binary, which only the PDF backend has.
 */

/**
 * Straight sRGB. `r`, `g` and `b` are integers in 0..255 and `a` is a fraction
 * in 0..1; a producer that emits a fractional channel produces CSS a backend
 * cannot serialize.
 */
export type DisplayColor = {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a: number;
};

/** A rectangle in page coordinates. */
export type DisplayRect = {
  readonly xPx: number;
  readonly yPx: number;
  readonly widthPx: number;
  readonly heightPx: number;
};

/**
 * Stroke patterns folio paints. `double` and `wavy` are patterns rather than
 * widths: both backends draw them from `thicknessPx`, so neither has to guess
 * what the other did. Dash and gap lengths live in `primitives.ts` and are
 * imported by both backends, so neither invents a period of its own.
 */
export type DisplayStrokePattern = "solid" | "dashed" | "dotted" | "double" | "wavy";

/**
 * A stroke is centred on its path: half the thickness falls on each side.
 * Stated because the backends default differently (a CSS border sits inside
 * its box, a PDF stroke is centred), and an unstated default is a divergence
 * that a reader finds rather than a test.
 */
export type DisplayStroke = {
  readonly color: DisplayColor;
  readonly thicknessPx: number;
  readonly pattern: DisplayStrokePattern;
};

/**
 * A face embedded in the source package (`fonts/embeddedFonts.ts`). Its bytes
 * travel with the display list because they exist nowhere else: a backend
 * cannot resolve them from a host font list, and a face the measurer used but
 * a backend cannot obtain is exactly the divergence this carries bytes to
 * prevent.
 */
export type DisplayEmbeddedFont = {
  readonly id: string;
  /** Deobfuscated sfnt or WOFF bytes. */
  readonly bytes: Uint8Array;
};

/**
 * A font face as the *measurer* resolved it. `family` is the resolved family
 * the measurement was taken against, not the authored OOXML name, so a
 * backend embedding a face embeds the face the widths came from.
 */
export type DisplayFontFace = {
  readonly family: string;
  /** CSS weight; folio paints 400 and 700 only. */
  readonly weight: number;
  readonly italic: boolean;
  /** Category to fall back on when the family itself is unavailable. */
  readonly generic: "serif" | "sans-serif" | "monospace" | "cursive" | "fantasy";
  /**
   * The face's own box, above and below the baseline, as fractions of the font
   * size, exactly as the measurer reported it (`FontMetrics.fontBox*`).
   *
   * These are the font-box metrics, not the ink extents, because a DOM backend
   * places a text box and a browser builds that box from the font box: given
   * ink extents it would put every baseline out by half the difference between
   * the two, per face, with nothing to attribute the error to. A backend
   * holding the font binary ignores them and positions glyphs directly.
   */
  readonly fontBoxAscentRatio: number;
  readonly fontBoxDescentRatio: number;
  /** Present only for a face embedded in the source package. */
  readonly embedded?: DisplayEmbeddedFont;
};

/** Index into {@link DisplayList.fonts}. */
export type DisplayFontRef = number;

/** Index into {@link DisplayList.images}. */
export type DisplayImageRef = number;

/**
 * Raw image bytes plus the format they are already in. `png` and `jpeg`
 * pass through to PDF unchanged; every other source format is transcoded by
 * the producer, so a backend never owns a codec.
 */
export type DisplayImageSource = {
  readonly format: "png" | "jpeg";
  readonly bytes: Uint8Array;
  readonly pixelWidth: number;
  readonly pixelHeight: number;
};

/**
 * Text drawn from one face at one size in one colour.
 *
 * `advancesPx` holds one advance per code point of `text` (not per UTF-16
 * unit) and sums to the run's painted width. These are the numbers line
 * breaking and pagination were decided on: a backend that positions glyphs any
 * other way paints a page the engine did not lay out.
 *
 * ## Ordering, stated so two backends cannot read it differently
 *
 * `text` is in logical order and `advancesPx` matches it index for index. The
 * run occupies `[xPx, xPx + sum(advancesPx)]` whatever its direction. Under
 * `ltr`, code point `i` starts at `xPx + sum(advancesPx[0..i))`; under `rtl`
 * it *ends* at `xPx + sum(advancesPx) - sum(advancesPx[0..i))`. A producer
 * never emits a run that spans a direction change, so no backend runs the bidi
 * algorithm and no two backends can disagree about its result.
 */
export type DisplayGlyphRun = {
  readonly kind: "glyphRun";
  readonly font: DisplayFontRef;
  readonly fontSizePx: number;
  readonly color: DisplayColor;
  /** Left edge of the run, before any horizontal scale. */
  readonly xPx: number;
  readonly baselineYPx: number;
  readonly text: string;
  readonly advancesPx: readonly number[];
  readonly direction: "ltr" | "rtl";
  /**
   * Glyph outline stroke, for `w:outline` and the emboss/imprint effects.
   * Absent means fill only.
   */
  readonly stroke?: DisplayStroke;
};

/** An axis-aligned rectangle, filled and/or stroked. At least one is set. */
export type DisplayRectPrimitive = {
  readonly kind: "rect";
  readonly rect: DisplayRect;
  readonly fill?: DisplayColor;
  readonly stroke?: DisplayStroke;
};

/**
 * A stroked segment. Borders, rules, column separators, table diagonals and
 * text decorations are all this one primitive, so a dashed underline and a
 * dashed cell border cannot be drawn by two different rules.
 */
export type DisplayLine = {
  readonly kind: "line";
  readonly x1Px: number;
  readonly y1Px: number;
  readonly x2Px: number;
  readonly y2Px: number;
  readonly stroke: DisplayStroke;
};

export type DisplayImagePrimitive = {
  readonly kind: "image";
  readonly image: DisplayImageRef;
  /** Destination box, after cropping and rotation have been resolved. */
  readonly rect: DisplayRect;
  /** Source crop as 0..1 fractions of the decoded image, if any. */
  readonly crop?: {
    readonly l: number;
    readonly t: number;
    readonly r: number;
    readonly b: number;
  };
  readonly opacity: number;
};

/** Clips its children to `rect`. Used for cropped images and cell overflow. */
export type DisplayClipGroup = {
  readonly kind: "clipGroup";
  readonly rect: DisplayRect;
  readonly children: readonly DisplayPrimitive[];
};

/**
 * Rotates its children about `originXPx`/`originYPx`. Rotated images,
 * watermark text and vertical (`w:textDirection`) table text are all this.
 * Rotation is the only transform folio paints; scale is resolved into
 * geometry by the producer.
 */
export type DisplayRotateGroup = {
  readonly kind: "rotateGroup";
  readonly degrees: number;
  readonly originXPx: number;
  readonly originYPx: number;
  readonly children: readonly DisplayPrimitive[];
};

/** Uniform alpha over a subtree, for watermark washout and image opacity. */
export type DisplayOpacityGroup = {
  readonly kind: "opacityGroup";
  readonly opacity: number;
  readonly children: readonly DisplayPrimitive[];
};

export type DisplayPrimitive =
  | DisplayGlyphRun
  | DisplayRectPrimitive
  | DisplayLine
  | DisplayImagePrimitive
  | DisplayClipGroup
  | DisplayRotateGroup
  | DisplayOpacityGroup;

/**
 * A clickable region. Annotations are not paint: they sit beside the
 * primitive list so a backend that cannot express them (a raster) drops them
 * without silently losing a drawing instruction.
 */
export type DisplayLink = {
  readonly rect: DisplayRect;
  readonly target: DisplayLinkTarget;
  readonly tooltip?: string;
};

export type DisplayLinkTarget =
  | { readonly kind: "external"; readonly href: string }
  /**
   * `pageIndex` indexes {@link DisplayList.pages}. It is not
   * {@link DisplayPage.pageNumber}, which is the engine's physical number and
   * may restart or skip, so a backend rendering one page in isolation cannot
   * resolve an internal target and must be handed the whole list.
   */
  | { readonly kind: "page"; readonly pageIndex: number; readonly yPx: number };

export type DisplayPage = {
  /** Physical 1-based page number, as the layout engine numbered it. */
  readonly pageNumber: number;
  readonly widthPx: number;
  readonly heightPx: number;
  readonly orientation: "portrait" | "landscape";
  /** Painted back to front. */
  readonly primitives: readonly DisplayPrimitive[];
  readonly links: readonly DisplayLink[];
};

/** One heading, for a document outline. `pageIndex` is 0-based. */
export type DisplayOutlineEntry = {
  readonly title: string;
  /** OOXML outline level, 0 for Heading 1. */
  readonly level: number;
  readonly pageIndex: number;
  readonly yPx: number;
};

export type DisplayMetadata = {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
};

/**
 * One construct the producer could not turn into primitives.
 *
 * Reported rather than dropped, on the same principle as `compareDocx`'s
 * `unsupported`: a page that is missing paint must say so, because a backend
 * cannot tell an empty region from an unimplemented one.
 */
export type DisplayUnsupported = {
  readonly construct: string;
  readonly pageIndex: number;
  readonly detail: string;
};

export type DisplayList = {
  readonly pages: readonly DisplayPage[];
  readonly fonts: readonly DisplayFontFace[];
  readonly images: readonly DisplayImageSource[];
  readonly outline: readonly DisplayOutlineEntry[];
  readonly metadata: DisplayMetadata;
  readonly unsupported: readonly DisplayUnsupported[];
};
