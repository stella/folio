/**
 * The DOM backend for the display list.
 *
 * Every mark is positioned absolutely from the coordinates the producer
 * already decided. Nothing here re-derives a position from inline flow
 * (`text-align`, `text-indent`, floats, line boxes): a second opinion about
 * where a glyph goes is the divergence the display list exists to remove.
 *
 * A glyph run is one element, and the display list places its origin. Where
 * each glyph lands inside the run is the browser's, because shaping is the
 * only thing that forms a ligature, applies a kern, or picks a cursive
 * letter's initial, medial or final form, and none of those can be recovered
 * from advances alone. The painted extent of a run can therefore differ from
 * its declared extent by the shaper's rounding.
 *
 * That residue is measured rather than assumed. It is small when the list was
 * built by the same engine that paints it, which is the editor's own path, and
 * larger when a list built from font-table advances is painted by a shaper
 * that kerns; the harness reports both and gates only on the first.
 *
 * This module takes a type-only edge to the paint IR and a runtime edge to its
 * companions in `../primitives`, which carry no layout fact. It has no access
 * to the layout engine, so a fact it needs and the display list does not carry
 * has to be added to the contract, where the PDF backend sees it too.
 *
 * `DisplayList.unsupported` is not paint and is not consumed here: the caller
 * surfaces it, because a page missing marks must say so.
 *
 * Object URLs minted for images and embedded fonts outlive this call, one per
 * image and one per embedded face for the whole render. A caller that discards
 * the returned elements should revoke them (`img[src^="blob:"]` and the `src`
 * of each `@font-face` rule).
 */

import { panic } from "better-result";

import {
  DOUBLE_STROKE_GAP_FACTOR,
  STROKE_DASH_FACTORS,
  WAVY_STROKE_AMPLITUDE_FACTOR,
  WAVY_STROKE_PERIOD_FACTOR,
} from "../primitives";
import type {
  DisplayClipGroup,
  DisplayColor,
  DisplayEmbeddedFont,
  DisplayFontFace,
  DisplayFontRef,
  DisplayGlyphRun,
  DisplayRunAdjustments,
  DisplayImagePrimitive,
  DisplayImageRef,
  DisplayImageSource,
  DisplayLine,
  DisplayLink,
  DisplayLinkTarget,
  DisplayList,
  DisplayOpacityGroup,
  DisplayPage,
  DisplayPrimitive,
  DisplayRect,
  DisplayRectPrimitive,
  DisplayRotateGroup,
  DisplayStroke,
  DisplayStrokePattern,
} from "../types";

const PAGE_CLASS_NAME = "layout-page";

/**
 * Anchor ids are the display list's own page indices, not `pageNumber`, which
 * the engine may restart or skip. One numbering, so a link target and the id
 * it addresses cannot drift apart.
 */
const PAGE_ELEMENT_ID_PREFIX = "page-";

const DEGREES_PER_RADIAN = 180 / Math.PI;

const BASE64_CHUNK_SIZE = 0x8000;

const IMAGE_MIME_TYPES = {
  png: "image/png",
  jpeg: "image/jpeg",
} as const satisfies Record<DisplayImageSource["format"], string>;

/**
 * The display list does not say whether an embedded face is sfnt or WOFF, so a
 * data URL declares neither and the browser sniffs; an `@font-face` `src` needs
 * no `format()` hint to load.
 */
const FONT_MIME_TYPE = "application/octet-stream";

/**
 * Which implementation draws a rect's stroke. Every pattern with a period or a
 * second line goes through the `line` renderer, so a dashed cell border and a
 * dashed underline cannot come out as two different dashes. `solid` has no
 * pattern to get wrong, so a CSS border expresses it exactly and cheaply.
 */
const RECT_STROKE_PAINTERS = {
  solid: "border",
  dashed: "edges",
  dotted: "edges",
  double: "edges",
  wavy: "edges",
} as const satisfies Record<DisplayStrokePattern, "border" | "edges">;

const px = (value: number) => `${value}px`;

/**
 * `rgb()` when opaque so serialized styles stay stable and diffable against
 * the PDF backend's colour dump.
 */
const cssColor = ({ r, g, b, a }: DisplayColor) =>
  a === 1 ? `rgb(${r}, ${g}, ${b})` : `rgba(${r}, ${g}, ${b}, ${a})`;

const quoteFontFamily = (family: string) => `"${family.replaceAll('"', '\\"')}"`;

/**
 * An embedded face wins over any installed face of the same name by being
 * named first, under the id its `@font-face` rule was registered with.
 */
/**
 * The stack the measurer resolved, rebuilt in CSS.
 *
 * Every entry, not just the first: the browser picks the first family it has,
 * and so did the measurement. A shorter stack here paints a different face on
 * any host missing the first one.
 */
const fontFamilyStack = ({ family, fallbacks, generic, embedded }: DisplayFontFace) => {
  const names = [family, ...fallbacks].map(quoteFontFamily).join(", ");
  const resolved = `${names}, ${generic}`;
  return embedded === undefined ? resolved : `${quoteFontFamily(embedded.id)}, ${resolved}`;
};

const resolveFont = (ref: DisplayFontRef, fonts: readonly DisplayFontFace[]) => {
  const face = ref < 0 ? undefined : fonts.at(ref);
  if (face === undefined) {
    panic(`renderDisplayListToDom: font ref ${ref} is out of range (${fonts.length} faces)`);
  }
  return face;
};

const resolveImage = (ref: DisplayImageRef, images: readonly DisplayImageSource[]) => {
  const source = ref < 0 ? undefined : images.at(ref);
  if (source === undefined) {
    panic(`renderDisplayListToDom: image ref ${ref} is out of range (${images.length} images)`);
  }
  return source;
};

const toBase64 = (bytes: Uint8Array) => {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += BASE64_CHUNK_SIZE) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + BASE64_CHUNK_SIZE));
  }
  return btoa(binary);
};

/** A `Blob` part cannot be backed by shared memory, so prove it is not. */
const isBlobPart = (bytes: Uint8Array): bytes is Uint8Array<ArrayBuffer> =>
  bytes.buffer instanceof ArrayBuffer;

const binarySrc = (bytes: Uint8Array, mimeType: string) => {
  if (
    isBlobPart(bytes) &&
    typeof Blob === "function" &&
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function"
  ) {
    return URL.createObjectURL(new Blob([bytes], { type: mimeType }));
  }
  if (typeof btoa === "function") {
    return `data:${mimeType};base64,${toBase64(bytes)}`;
  }
  panic("renderDisplayListToDom: no Blob or btoa to build a binary source from");
};

const fontFaceRule = (face: DisplayFontFace, embedded: DisplayEmbeddedFont, src: string) =>
  `@font-face { font-family: ${quoteFontFamily(embedded.id)}; src: url("${src}"); font-weight: ${face.weight}; font-style: ${face.italic ? "italic" : "normal"}; }`;

/**
 * `@font-face` rules for the faces whose bytes travel with the list. Resolved
 * once per render and shared by every page: an id maps to one blob, so the
 * bytes are held once and one revoke releases them.
 */
const embeddedFontFaceCss = (fonts: readonly DisplayFontFace[]) => {
  const srcById = new Map<string, string>();
  const rules: string[] = [];
  for (const face of fonts) {
    const embedded = face.embedded;
    if (embedded === undefined) {
      continue;
    }
    const cached = srcById.get(embedded.id);
    const src = cached ?? binarySrc(embedded.bytes, FONT_MIME_TYPE);
    if (cached === undefined) {
      srcById.set(embedded.id, src);
    }
    rules.push(fontFaceRule(face, embedded, src));
  }
  return rules.join("\n");
};

/**
 * The paint target plus the page-absolute coordinate of its origin. Primitive
 * coordinates are page-absolute; a container that establishes its own origin
 * (a clip box) subtracts it once, here, rather than at every call site.
 */
type PaintContext = {
  readonly doc: Document;
  readonly fonts: readonly DisplayFontFace[];
  readonly images: readonly DisplayImageSource[];
  readonly parent: HTMLElement;
  readonly originXPx: number;
  readonly originYPx: number;
};

const createAbsoluteDiv = (context: PaintContext, rect: DisplayRect) => {
  const element = context.doc.createElement("div");
  element.style.position = "absolute";
  element.style.left = px(rect.xPx - context.originXPx);
  element.style.top = px(rect.yPx - context.originYPx);
  element.style.width = px(rect.widthPx);
  element.style.height = px(rect.heightPx);
  return element;
};

const createFilledDiv = (
  context: PaintContext,
  rect: DisplayRect,
  fill: DisplayColor | undefined,
) => {
  const element = createAbsoluteDiv(context, rect);
  if (fill !== undefined) {
    element.style.backgroundColor = cssColor(fill);
  }
  return element;
};

/**
 * A group that transforms its children without changing their coordinates: a
 * zero-sized positioned box at the container's origin is the containing block,
 * so a child keeps the coordinates it had one level up.
 */
const createTransparentGroupDiv = (context: PaintContext) => {
  const element = context.doc.createElement("div");
  element.style.position = "absolute";
  element.style.left = "0px";
  element.style.top = "0px";
  element.style.width = "0px";
  element.style.height = "0px";
  return element;
};

/** Gradient directions in a bar's local space: along it, then across it. */
type BarAxes = {
  readonly main: string;
  readonly cross: string;
};

const HORIZONTAL_BAR_AXES = { main: "to right", cross: "to bottom" } as const;
const VERTICAL_BAR_AXES = { main: "to bottom", cross: "to right" } as const;

/**
 * How far a stroke reaches across its path. A dash pattern stays within the
 * thickness; `double` spans two lines plus the shared gap, and `wavy` the peak
 * to peak amplitude plus the thickness the curve itself is drawn with.
 */
const strokeCrossExtentPx = (stroke: DisplayStroke) => {
  switch (stroke.pattern) {
    case "solid":
    case "dashed":
    case "dotted":
      return stroke.thicknessPx;
    case "double":
      return stroke.thicknessPx * (2 + DOUBLE_STROKE_GAP_FACTOR);
    case "wavy":
      return stroke.thicknessPx * (WAVY_STROKE_AMPLITUDE_FACTOR + 1);
    default: {
      const unreachable: never = stroke.pattern;
      panic(`renderDisplayListToDom: unhandled stroke pattern ${JSON.stringify(unreachable)}`);
    }
  }
};

const repeatingBars = (color: string, axis: string, barPx: number, gapPx: number) =>
  `repeating-linear-gradient(${axis}, ${color} 0px, ${color} ${barPx}px, transparent ${barPx}px, transparent ${barPx + gapPx}px)`;

const strokeBackground = (stroke: DisplayStroke, axes: BarAxes) => {
  const color = cssColor(stroke.color);
  switch (stroke.pattern) {
    case "solid":
      return color;
    case "dashed":
    case "dotted": {
      // Dotted dots are square: a round dot needs a repeating radial gradient,
      // whose period a PDF dash array cannot reproduce.
      const { dash, gap } = STROKE_DASH_FACTORS[stroke.pattern];
      return repeatingBars(color, axes.main, stroke.thicknessPx * dash, stroke.thicknessPx * gap);
    }
    case "double": {
      // Two lines and one gap fill the cross extent, so each is one third of it
      // whenever the gap is a single thickness.
      const rulePercent = 100 / (2 + DOUBLE_STROKE_GAP_FACTOR);
      return `linear-gradient(${axes.cross}, ${color} 0%, ${color} ${rulePercent}%, transparent ${rulePercent}%, transparent ${100 - rulePercent}%, ${color} ${100 - rulePercent}%)`;
    }
    case "wavy": {
      // Approximate. No CSS gradient draws a sine wave, so this is a diagonal
      // hatch of the contract's period across the contract's amplitude; a
      // backend comparison must treat `wavy` as inexact, not as a mismatch.
      const halfPeriod = (stroke.thicknessPx * WAVY_STROKE_PERIOD_FACTOR) / 2;
      return repeatingBars(color, "45deg", halfPeriod, halfPeriod);
    }
    default: {
      const unreachable: never = stroke.pattern;
      panic(`renderDisplayListToDom: unhandled stroke pattern ${JSON.stringify(unreachable)}`);
    }
  }
};

/** Strokes are centred on their path, so half the cross extent falls each side. */
const paintLineSegment = (line: DisplayLine, context: PaintContext) => {
  const { x1Px, y1Px, x2Px, y2Px, stroke } = line;
  const crossExtentPx = strokeCrossExtentPx(stroke);
  const halfCrossExtentPx = crossExtentPx / 2;
  const element = context.doc.createElement("div");
  element.style.position = "absolute";
  element.style.boxSizing = "content-box";

  if (y1Px === y2Px) {
    element.style.left = px(Math.min(x1Px, x2Px) - context.originXPx);
    element.style.top = px(y1Px - halfCrossExtentPx - context.originYPx);
    element.style.width = px(Math.abs(x2Px - x1Px));
    element.style.height = px(crossExtentPx);
    element.style.background = strokeBackground(stroke, HORIZONTAL_BAR_AXES);
  } else if (x1Px === x2Px) {
    element.style.left = px(x1Px - halfCrossExtentPx - context.originXPx);
    element.style.top = px(Math.min(y1Px, y2Px) - context.originYPx);
    element.style.width = px(crossExtentPx);
    element.style.height = px(Math.abs(y2Px - y1Px));
    element.style.background = strokeBackground(stroke, VERTICAL_BAR_AXES);
  } else {
    const dx = x2Px - x1Px;
    const dy = y2Px - y1Px;
    element.style.left = px(x1Px - context.originXPx);
    element.style.top = px(y1Px - halfCrossExtentPx - context.originYPx);
    element.style.width = px(Math.hypot(dx, dy));
    element.style.height = px(crossExtentPx);
    element.style.transformOrigin = "0 50%";
    element.style.transform = `rotate(${Math.atan2(dy, dx) * DEGREES_PER_RADIAN}deg)`;
    element.style.background = strokeBackground(stroke, HORIZONTAL_BAR_AXES);
  }

  context.parent.append(element);
};

/**
 * Reapply what the measurer added to the advances.
 *
 * The browser shapes the run's text and advances it by what the font says, and
 * none of these three is in the font: without them the painted run is as wide
 * as the glyphs alone, which is not the width the line was fitted at. The CSS
 * is the same CSS the legacy painter emits, so the two agree by construction
 * rather than by coincidence.
 *
 * Spacing goes on before the transform scales it, so the painted-pixel numbers
 * the list carries are divided by the scale first.
 */
const applyAdjustments = (
  span: HTMLElement,
  adjustments: DisplayRunAdjustments | undefined,
): void => {
  if (adjustments === undefined) {
    return;
  }
  const { letterSpacingPx, horizontalScale, wordSpacingPx } = adjustments;
  if (letterSpacingPx !== 0) {
    span.style.letterSpacing = px(letterSpacingPx / horizontalScale);
  }
  if (wordSpacingPx !== 0) {
    span.style.wordSpacing = px(wordSpacingPx / horizontalScale);
  }
  if (horizontalScale !== 1) {
    // The run's own box is already the scaled width, so the transform must not
    // widen it again: it scales the glyphs inside a box the list sized.
    span.style.transform = `scaleX(${String(horizontalScale)})`;
    span.style.transformOrigin = "left center";
  }
};

const paintGlyphRun = (run: DisplayGlyphRun, context: PaintContext) => {
  if ([...run.text].length !== run.advancesPx.length) {
    panic(
      `renderDisplayListToDom: glyph run carries ${run.advancesPx.length} advances for ${[...run.text].length} code points`,
    );
  }
  const face = resolveFont(run.font, context.fonts);
  const advanceSum = run.advancesPx.reduce((total, advance) => total + advance, 0);
  const ascentPx = face.fontBoxAscentRatio * run.fontSizePx;
  const descentPx = face.fontBoxDescentRatio * run.fontSizePx;

  const span = context.doc.createElement("span");
  // A browser builds an inline box from the font box, so a line box of exactly
  // that box's height puts its baseline one ascent below its top: subtracting
  // the ascent from the baseline is exact, with no metric guessed here.
  span.style.position = "absolute";
  span.style.left = px(run.xPx - context.originXPx);
  span.style.top = px(run.baselineYPx - ascentPx - context.originYPx);
  span.style.lineHeight = px(ascentPx + descentPx);
  span.style.display = "inline-block";
  span.style.boxSizing = "content-box";
  // The declared extent. The browser shapes and advances the glyphs inside it,
  // so the painted extent can differ from this by the shaper's rounding; that
  // residue is measured rather than assumed (see the module header). A scaled
  // run's box is stated before the transform, which then scales it to the
  // declared width.
  span.style.width = px(advanceSum / (run.adjustments?.horizontalScale ?? 1));
  span.style.whiteSpace = "pre";
  span.style.fontFamily = fontFamilyStack(face);
  span.style.fontSize = px(run.fontSizePx);
  span.style.fontWeight = String(face.weight);
  span.style.fontStyle = face.italic ? "italic" : "normal";
  span.style.color = cssColor(run.color);
  // The run never spans a direction change, so the browser's ordering within
  // it is the producer's ordering: no `unicode-bidi` override is needed.
  span.style.direction = run.direction;
  if (run.stroke !== undefined) {
    // A glyph outline can only be one solid width in CSS: the stroke's pattern
    // is lost here, where PDF can dash it.
    span.style.webkitTextStroke = `${px(run.stroke.thicknessPx)} ${cssColor(run.stroke.color)}`;
  }
  // Stated, never defaulted: a browser asked nothing kerns whenever the face
  // has the table, and these advances may have been measured with it off.
  span.style.fontKerning = run.kerning ? "normal" : "none";
  if (run.smallCaps) {
    span.style.fontVariant = "small-caps";
  }
  applyAdjustments(span, run.adjustments);
  // The producer's own number, readable back out of the DOM by the
  // equivalence harness without measuring anything.
  span.dataset["advanceSum"] = String(advanceSum);

  // One text node, shaped by the browser. The display list places the run's
  // origin; where each glyph lands inside it is the shaper's business, which
  // is the only thing that can form a ligature, apply a kern or choose a
  // cursive positional form. Placing code points independently would override
  // all three to buy a sub-pixel advance nobody sees.
  span.textContent = run.text;

  context.parent.append(span);
};

const grownRect = (rect: DisplayRect, amountPx: number) => ({
  xPx: rect.xPx - amountPx / 2,
  yPx: rect.yPx - amountPx / 2,
  widthPx: rect.widthPx + amountPx,
  heightPx: rect.heightPx + amountPx,
});

const rectEdges = ({ xPx, yPx, widthPx, heightPx }: DisplayRect) => {
  const rightPx = xPx + widthPx;
  const bottomPx = yPx + heightPx;
  return [
    { x1Px: xPx, y1Px: yPx, x2Px: rightPx, y2Px: yPx },
    { x1Px: rightPx, y1Px: yPx, x2Px: rightPx, y2Px: bottomPx },
    { x1Px: xPx, y1Px: bottomPx, x2Px: rightPx, y2Px: bottomPx },
    { x1Px: xPx, y1Px: yPx, x2Px: xPx, y2Px: bottomPx },
  ];
};

const paintRect = ({ rect, fill, stroke }: DisplayRectPrimitive, context: PaintContext) => {
  if (stroke === undefined) {
    context.parent.append(createFilledDiv(context, rect, fill));
    return;
  }

  const painter = RECT_STROKE_PAINTERS[stroke.pattern];
  switch (painter) {
    case "border": {
      // Grown by the thickness so a border-box border straddles the path.
      const element = createFilledDiv(context, grownRect(rect, stroke.thicknessPx), fill);
      element.style.boxSizing = "border-box";
      element.style.border = `${px(stroke.thicknessPx)} solid ${cssColor(stroke.color)}`;
      context.parent.append(element);
      return;
    }
    case "edges": {
      context.parent.append(createFilledDiv(context, rect, fill));
      for (const edge of rectEdges(rect)) {
        paintLineSegment({ kind: "line", ...edge, stroke }, context);
      }
      return;
    }
    default: {
      const unreachable: never = painter;
      panic(`renderDisplayListToDom: unhandled rect stroke painter ${JSON.stringify(unreachable)}`);
    }
  }
};

const paintImage = (
  { image, rect, crop, opacity }: DisplayImagePrimitive,
  context: PaintContext,
) => {
  const source = resolveImage(image, context.images);
  const clip = createAbsoluteDiv(context, rect);
  clip.style.overflow = "hidden";

  const element = context.doc.createElement("img");
  element.style.position = "absolute";
  element.style.opacity = String(opacity);
  element.style.left = "0px";
  element.style.top = "0px";
  element.style.width = px(rect.widthPx);
  element.style.height = px(rect.heightPx);

  if (crop !== undefined) {
    const visibleWidth = 1 - crop.l - crop.r;
    const visibleHeight = 1 - crop.t - crop.b;
    if (visibleWidth <= 0 || visibleHeight <= 0) {
      panic(`renderDisplayListToDom: image crop leaves nothing visible (${JSON.stringify(crop)})`);
    }
    // Oversize the image so the destination box shows exactly the crop.
    const fullWidth = rect.widthPx / visibleWidth;
    const fullHeight = rect.heightPx / visibleHeight;
    element.style.left = px(-crop.l * fullWidth);
    element.style.top = px(-crop.t * fullHeight);
    element.style.width = px(fullWidth);
    element.style.height = px(fullHeight);
  }

  // The display list carries no description for an image.
  element.alt = "";
  element.src = binarySrc(source.bytes, IMAGE_MIME_TYPES[source.format]);
  clip.append(element);
  context.parent.append(clip);
};

const paintClipGroup = ({ rect, children }: DisplayClipGroup, context: PaintContext) => {
  const element = createAbsoluteDiv(context, rect);
  element.style.overflow = "hidden";
  context.parent.append(element);

  const inner = {
    ...context,
    parent: element,
    originXPx: rect.xPx,
    originYPx: rect.yPx,
  };
  for (const child of children) {
    paintPrimitive(child, inner);
  }
};

const paintRotateGroup = (
  { degrees, originXPx, originYPx, children }: DisplayRotateGroup,
  context: PaintContext,
) => {
  const element = createTransparentGroupDiv(context);
  element.style.transformOrigin = `${px(originXPx - context.originXPx)} ${px(originYPx - context.originYPx)}`;
  element.style.transform = `rotate(${degrees}deg)`;
  context.parent.append(element);

  const inner = { ...context, parent: element };
  for (const child of children) {
    paintPrimitive(child, inner);
  }
};

const paintOpacityGroup = ({ opacity, children }: DisplayOpacityGroup, context: PaintContext) => {
  const element = createTransparentGroupDiv(context);
  element.style.opacity = String(opacity);
  context.parent.append(element);

  const inner = { ...context, parent: element };
  for (const child of children) {
    paintPrimitive(child, inner);
  }
};

const paintPrimitive = (primitive: DisplayPrimitive, context: PaintContext): void => {
  switch (primitive.kind) {
    case "glyphRun":
      return paintGlyphRun(primitive, context);
    case "rect":
      return paintRect(primitive, context);
    case "line":
      return paintLineSegment(primitive, context);
    case "image":
      return paintImage(primitive, context);
    case "clipGroup":
      return paintClipGroup(primitive, context);
    case "rotateGroup":
      return paintRotateGroup(primitive, context);
    case "opacityGroup":
      return paintOpacityGroup(primitive, context);
    default: {
      const unreachable: never = primitive;
      panic(`renderDisplayListToDom: unhandled display primitive ${JSON.stringify(unreachable)}`);
    }
  }
};

/**
 * A page target loses its `yPx`: a fragment identifier addresses an element,
 * not an offset within one.
 */
const linkHref = (target: DisplayLinkTarget) => {
  switch (target.kind) {
    case "external":
      return target.href;
    case "page":
      return `#${PAGE_ELEMENT_ID_PREFIX}${target.pageIndex}`;
    default: {
      const unreachable: never = target;
      panic(`renderDisplayListToDom: unhandled link target ${JSON.stringify(unreachable)}`);
    }
  }
};

const paintLink = ({ rect, target, tooltip }: DisplayLink, context: PaintContext) => {
  const element = context.doc.createElement("a");
  element.style.position = "absolute";
  element.style.left = px(rect.xPx - context.originXPx);
  element.style.top = px(rect.yPx - context.originYPx);
  element.style.width = px(rect.widthPx);
  element.style.height = px(rect.heightPx);
  element.href = linkHref(target);
  if (tooltip !== undefined) {
    element.title = tooltip;
  }
  context.parent.append(element);
};

export type RenderDisplayListOptions = {
  /** Document to create elements in. */
  readonly doc: Document;
  /** Page background, painted before any primitive. */
  readonly pageBackground?: DisplayColor;
};

/**
 * A page's glyph runs and images index tables that live on the list, and an
 * internal link addresses a page by its index in the same list, so rendering
 * one page needs all three passed alongside it.
 */
export type RenderDisplayPageOptions = RenderDisplayListOptions & {
  /** `DisplayList.fonts`. */
  readonly fonts: readonly DisplayFontFace[];
  /** `DisplayList.images`. */
  readonly images: readonly DisplayImageSource[];
  /** The page's index in `DisplayList.pages`, which names its anchor. */
  readonly pageIndex: number;
};

type RenderPageOptions = RenderDisplayPageOptions & {
  /** Resolved once per render so every page shares one blob per face. */
  readonly fontFaceCss: string;
};

const renderPage = (page: DisplayPage, options: RenderPageOptions) => {
  const element = options.doc.createElement("div");
  element.className = PAGE_CLASS_NAME;
  element.id = `${PAGE_ELEMENT_ID_PREFIX}${options.pageIndex}`;
  element.style.position = "relative";
  element.style.width = px(page.widthPx);
  element.style.height = px(page.heightPx);
  element.style.overflow = "hidden";
  if (options.pageBackground !== undefined) {
    element.style.backgroundColor = cssColor(options.pageBackground);
  }

  if (options.fontFaceCss.length > 0) {
    const style = options.doc.createElement("style");
    style.textContent = options.fontFaceCss;
    element.append(style);
  }

  const context = {
    doc: options.doc,
    fonts: options.fonts,
    images: options.images,
    parent: element,
    originXPx: 0,
    originYPx: 0,
  };
  // Painted back to front: DOM order is the display list's order, with no
  // z-index to re-sort what the producer already stacked.
  for (const primitive of page.primitives) {
    paintPrimitive(primitive, context);
  }
  // Links sit above the paint by coming last.
  for (const link of page.links) {
    paintLink(link, context);
  }

  return element;
};

export const renderDisplayPageToDom = (
  page: DisplayPage,
  options: RenderDisplayPageOptions,
): HTMLElement => renderPage(page, { ...options, fontFaceCss: embeddedFontFaceCss(options.fonts) });

/** One `div.layout-page` per display page, in page order. */
export const renderDisplayListToDom = (
  list: DisplayList,
  options: RenderDisplayListOptions,
): HTMLElement[] => {
  const fontFaceCss = embeddedFontFaceCss(list.fonts);
  return list.pages.map((page, pageIndex) =>
    renderPage(page, {
      ...options,
      fonts: list.fonts,
      images: list.images,
      pageIndex,
      fontFaceCss,
    }),
  );
};
