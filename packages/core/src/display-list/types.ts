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
   * The concrete families between `family` and `generic`, in the order the
   * measurer would fall through them.
   *
   * A face is a stack, not a name: the measurer resolves one and measures with
   * whichever entry the host actually has. A backend handed only the first
   * entry falls straight to the generic when that entry is missing, and paints
   * a different face from the one the page was laid out in. Empty when the
   * resolved stack is one family and a generic.
   */
  readonly fallbacks: readonly string[];
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
 * Which document a run's model positions address.
 *
 * A page is not one document. Its body, each header and footer part, and each
 * note are separate stories with separate position spaces, and the same number
 * means a different character in each. A range that did not say which one it
 * belonged to could only be used for the body, which is why the producer used
 * to drop the others and the painter used to strip them; naming the story is
 * what lets an editing surface route a click into the story it landed in.
 */
export type DisplayStoryRef =
  | { readonly kind: "body" }
  /** `rId` is the relationship that names the part, as the section selects it. */
  | { readonly kind: "header" | "footer"; readonly rId: string }
  /** `id` is the note's own `w:footnote`/`w:endnote` id. */
  | { readonly kind: "footnote" | "endnote"; readonly id: number };

export type DisplayModelRange = {
  readonly start: number;
  readonly end: number;
  readonly story: DisplayStoryRef;
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
/**
 * What the measurer added on top of the glyphs' own advances.
 *
 * `advancesPx` already carries all three, because that is what line breaking
 * decided the page on. They are named again here because a backend that hands
 * the text to a shaper cannot recover them from it: a shaper advances glyphs by
 * what the font says, and letter spacing, a horizontal scale and the space
 * compression a justified line was fitted with are not in the font. Such a
 * backend reapplies them and lands on the same extent; one that positions every
 * glyph itself ignores this and reads `advancesPx`.
 *
 * Absent means none of the three applies, which is the common run.
 *
 * Every number is in the same painted pixels as `advancesPx`: `letterSpacingPx`
 * and `wordSpacingPx` are what a code point actually gained on the page, after
 * `horizontalScale`. A backend that applies the scale as a transform divides by
 * it first, because a transform scales the spacing it is given.
 */
export type DisplayRunAdjustments = {
  /** Added after every code point of the run but the last. */
  readonly letterSpacingPx: number;
  /** Multiplies every advance. 1 when the run carries no `w:w`. */
  readonly horizontalScale: number;
  /** Added to each compressible space, negative on a contracted line. */
  readonly wordSpacingPx: number;
};

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
   * Whether the glyphs kern.
   *
   * Stated rather than left to the backend's default, because the two defaults
   * disagree: OOXML kerns only above the `w:kern` size and a run without it
   * does not kern at all, while a browser asked nothing kerns whenever the face
   * has the table. A run measured unkerned and painted kerned is narrower on
   * the page than the line it was fitted into.
   */
  readonly kerning: boolean;
  /**
   * Whether the run's lowercase letters are drawn as small capitals.
   *
   * Carried because it changes every advance in the run: a backend that paints
   * the text at full size occupies a different width from the one the line was
   * fitted to. Which glyphs a face uses for them, real `smcp` forms or scaled
   * capitals, is the backend's business.
   */
  readonly smallCaps: boolean;
  /** What the measurer added to the advances; absent when it added nothing. */
  readonly adjustments?: DisplayRunAdjustments;
  /**
   * A run of spaces at a line edge that the line was fitted without.
   *
   * Word keeps such spaces addressable and paints them at no width, so the
   * caret can sit among them while the line breaks as though they were not
   * there. The advances are therefore all zero, and `spaceAdvancePx` is what
   * one of those spaces would have been worth: an editing surface steps the
   * caret by it, because there is nothing painted to step over.
   */
  readonly collapsedEdge?: {
    readonly side: "leading" | "trailing";
    readonly spaceAdvancePx: number;
  };
  /**
   * Glyph outline stroke, for `w:outline` and the emboss/imprint effects.
   * Absent means fill only.
   */
  readonly stroke?: DisplayStroke;
  /**
   * Editable-model range this run's text occupies, when the producer knows it.
   *
   * Not paint, and a backend that only draws ignores it, the way a rasterizer
   * ignores `links`. It is here because an editing surface has to map a click
   * and a selection back to a position in the document, and the only structure
   * that knows which glyphs came from which characters is the one that placed
   * them. Absent when a run has no counterpart in the model: a list marker, a
   * substituted field value, a tab leader.
   */
  readonly pmRange?: DisplayModelRange;
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

/**
 * What kind of thing a click landed in.
 *
 * A surface that edits a page has to answer more than "which character": which
 * story, which paragraph, which cell, whether it hit a picture or the space
 * after a line. The painter answered those from its own element structure, and
 * every reader learned that structure; naming them here makes the answer a
 * property of the page rather than of the markup that happened to paint it.
 */
export type DisplayHitRegionKind =
  /** The page's content area, inside its margins. */
  | "pageContent"
  | "headerSlot"
  | "footerSlot"
  /** One note's body in the band at the foot of the page. */
  | "note"
  | "paragraph"
  /** One laid-out line of a paragraph, full content width. */
  | "line"
  /** A line with no runs: the caret still has to land somewhere. */
  | "emptyRun"
  /** A tab's advance, which is space rather than characters. */
  | "tab"
  | "image"
  | "table"
  | "tableRow"
  | "tableCell"
  | "textBox";

/**
 * What a region resolves to besides its box.
 *
 * Every field is something a reader of the painted markup used to get from an
 * attribute the painter wrote. They are optional because they are properties of
 * particular kinds: a cell has a column index and a paragraph does not.
 */
export type DisplayHitRegionModel = {
  /** Model range and story, for a region a caret can land in. */
  readonly pmRange?: DisplayModelRange;
  /** The story a slot or note region *is*, whether or not it has a range. */
  readonly story?: DisplayStoryRef;
  /** Durable block identity, for scrolling to and highlighting a block. */
  readonly blockId?: string;
  /** Comment threads anchored on this region's text. */
  readonly commentIds?: readonly number[];
  /** Row and column, for a cell; row only, for a row. */
  readonly rowIndex?: number;
  readonly columnIndex?: number;
  /** Leading and trailing spaces the line collapsed, and what they advanced. */
  readonly collapsedLeadingSpaces?: boolean;
  readonly collapsedTrailingSpaces?: boolean;
  readonly collapsedSpaceAdvancePx?: number;
};

/**
 * A box a click can land in, and what it means.
 *
 * Regions form a tree over the page's primitives rather than a copy of them:
 * `from` and `to` are a half-open range into `DisplayPage.primitives`, and a
 * child's range lies inside its parent's. The producer emits a line's
 * primitives together, so every region owns a contiguous slice; a backend that
 * paints structure walks the tree and a backend that only draws ignores it.
 *
 * Primitives no region claims are the page's own furniture: its background, its
 * borders, a watermark. Nothing about them is editable, so nothing has to
 * resolve them.
 */
export type DisplayHitRegion = {
  readonly kind: DisplayHitRegionKind;
  readonly rect: DisplayRect;
  /** First primitive of this region, as an index into the page's list. */
  readonly from: number;
  /** One past its last. */
  readonly to: number;
  readonly children: readonly DisplayHitRegion[];
  readonly model?: DisplayHitRegionModel;
};

export type DisplayPage = {
  /** Physical 1-based page number, as the layout engine numbered it. */
  readonly pageNumber: number;
  readonly widthPx: number;
  readonly heightPx: number;
  readonly orientation: "portrait" | "landscape";
  /** Painted back to front. */
  readonly primitives: readonly DisplayPrimitive[];
  /**
   * Where a click can land on this page, over the same primitives.
   *
   * Not paint: a rasterizer ignores it, as it ignores `links`. It is here
   * because the surface that edits a page and the backend that paints it must
   * agree about what is where, and they can only do that from one structure.
   */
  readonly regions: readonly DisplayHitRegion[];
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
