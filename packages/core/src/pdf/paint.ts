/**
 * Display list to content stream.
 *
 * Everything here is emitted in display-list coordinates: CSS px, origin at
 * the page's top-left, y down. The page's base CTM (see `pageSpace.ts`)
 * converts once, at the stream's first operator, which is why no primitive
 * below does any arithmetic on the page height.
 */

import { panic, Result, TaggedError } from "better-result";
import {
  DOUBLE_STROKE_GAP_FACTOR,
  glyphCellOffsetsPx,
  STROKE_DASH_FACTORS,
  WAVY_STROKE_AMPLITUDE_FACTOR,
  WAVY_STROKE_PERIOD_FACTOR,
} from "../display-list/primitives";
import type {
  DisplayGlyphRun,
  DisplayImagePrimitive,
  DisplayPage,
  DisplayPrimitive,
  DisplayRectPrimitive,
  DisplayStroke,
} from "../display-list/types";
import {
  createContentStream,
  TEXT_RENDER_MODE,
  type ContentPart,
  type ContentStream,
  type PositionedGlyph,
} from "./contentStream";
import type { PreparedFont } from "./fonts";
import { basePageMatrix, rotationMatrix } from "./pageSpace";

export class PaintError extends TaggedError("PaintError")<{ message: string }> {}

/** Thousandths of an em, the unit of a `TJ` adjustment. */
const TEXT_SPACE_UNITS_PER_EM = 1000;

const OPAQUE = 1;

/** Points sampled along one period of a `wavy` stroke. */
const WAVY_SAMPLES_PER_PERIOD = 8;

/** A `double` stroke is drawn as two lines, one either side of the path. */
const DOUBLE_STROKE_LINES = 2;

type PaintContext = {
  readonly stream: ContentStream;
  readonly fonts: ReadonlyMap<number, PreparedFont>;
  /** Alpha inherited from enclosing opacity groups. */
  readonly alpha: number;
};

const withAlpha = (context: PaintContext, alpha: number): PaintContext => ({
  ...context,
  alpha,
});

/**
 * Opens a graphics-state block for one primitive. Alpha is multiplied down
 * from the enclosing groups rather than replaced, so a half-transparent run
 * inside a half-transparent watermark paints at a quarter and not at a half.
 * PDF's `/ca` is per-object, so a group's alpha reaches its children this way
 * instead of compositing the group as a unit; overlapping children inside one
 * opacity group therefore show their seams where a true transparency group
 * would not.
 */
const beginPrimitive = (context: PaintContext, alpha: number) => {
  context.stream.save();
  if (alpha < OPAQUE) {
    context.stream.setAlpha(alpha);
  }
};

const applyStrokeState = (stream: ContentStream, stroke: DisplayStroke, widthPx: number) => {
  stream.setStrokeColor(stroke.color);
  stream.setLineWidth(widthPx);
  const dash = STROKE_DASH_FACTORS[stroke.pattern];
  if (dash !== null) {
    stream.setDash([dash.dash * stroke.thicknessPx, dash.gap * stroke.thicknessPx], 0);
  }
};

type Segment = {
  readonly x1Px: number;
  readonly y1Px: number;
  readonly x2Px: number;
  readonly y2Px: number;
};

/** Unit normal of a segment, or null when the segment has no length. */
const segmentNormal = (segment: Segment): { readonly nx: number; readonly ny: number } | null => {
  const dx = segment.x2Px - segment.x1Px;
  const dy = segment.y2Px - segment.y1Px;
  const length = Math.hypot(dx, dy);
  if (length === 0) {
    return null;
  }
  return { nx: -dy / length, ny: dx / length };
};

const strokeWavy = (stream: ContentStream, stroke: DisplayStroke, segment: Segment) => {
  const normal = segmentNormal(segment);
  if (normal === null) {
    return;
  }
  const dx = segment.x2Px - segment.x1Px;
  const dy = segment.y2Px - segment.y1Px;
  const length = Math.hypot(dx, dy);
  const period = WAVY_STROKE_PERIOD_FACTOR * stroke.thicknessPx;
  const amplitude = (WAVY_STROKE_AMPLITUDE_FACTOR * stroke.thicknessPx) / 2;
  const steps = Math.max(
    WAVY_SAMPLES_PER_PERIOD,
    Math.ceil((length / period) * WAVY_SAMPLES_PER_PERIOD),
  );
  for (let step = 0; step <= steps; step += 1) {
    const along = (length * step) / steps;
    const offset = amplitude * Math.sin((2 * Math.PI * along) / period);
    const x = segment.x1Px + (dx * step) / steps + normal.nx * offset;
    const y = segment.y1Px + (dy * step) / steps + normal.ny * offset;
    if (step === 0) {
      stream.moveTo(x, y);
    } else {
      stream.lineTo(x, y);
    }
  }
  stream.stroke();
};

const strokeDouble = (stream: ContentStream, stroke: DisplayStroke, segment: Segment) => {
  const normal = segmentNormal(segment);
  if (normal === null) {
    return;
  }
  // Two lines of the authored thickness with one thickness of gap between
  // them, centred on the path: their centres sit one thickness apart.
  const offset = (DOUBLE_STROKE_GAP_FACTOR * stroke.thicknessPx * DOUBLE_STROKE_LINES) / 2;
  for (const side of [-1, 1]) {
    stream.moveTo(
      segment.x1Px + normal.nx * offset * side,
      segment.y1Px + normal.ny * offset * side,
    );
    stream.lineTo(
      segment.x2Px + normal.nx * offset * side,
      segment.y2Px + normal.ny * offset * side,
    );
    stream.stroke();
  }
};

const paintSegment = (context: PaintContext, stroke: DisplayStroke, segment: Segment) => {
  const alpha = context.alpha * stroke.color.a;
  beginPrimitive(context, alpha);
  const { stream } = context;
  switch (stroke.pattern) {
    case "solid":
    case "dashed":
    case "dotted":
      applyStrokeState(stream, stroke, stroke.thicknessPx);
      stream.moveTo(segment.x1Px, segment.y1Px);
      stream.lineTo(segment.x2Px, segment.y2Px);
      stream.stroke();
      break;
    case "double":
      applyStrokeState(stream, stroke, stroke.thicknessPx);
      strokeDouble(stream, stroke, segment);
      break;
    case "wavy":
      applyStrokeState(stream, stroke, stroke.thicknessPx);
      strokeWavy(stream, stroke, segment);
      break;
    default: {
      const unreachable: never = stroke.pattern;
      panic(`unhandled stroke pattern: ${String(unreachable)}`);
    }
  }
  stream.restore();
};

const paintRect = (context: PaintContext, primitive: DisplayRectPrimitive) => {
  const { rect, fill, stroke } = primitive;
  if (fill !== undefined) {
    beginPrimitive(context, context.alpha * fill.a);
    context.stream.setFillColor(fill);
    context.stream.appendRect(rect.xPx, rect.yPx, rect.widthPx, rect.heightPx);
    context.stream.fill();
    context.stream.restore();
  }
  if (stroke === undefined) {
    return;
  }
  if (stroke.pattern === "solid") {
    beginPrimitive(context, context.alpha * stroke.color.a);
    applyStrokeState(context.stream, stroke, stroke.thicknessPx);
    context.stream.appendRect(rect.xPx, rect.yPx, rect.widthPx, rect.heightPx);
    context.stream.stroke();
    context.stream.restore();
    return;
  }
  // Every non-solid border is drawn edge by edge rather than as one
  // rectangle path, because that is what the DOM backend does: a dash array
  // applied to a closed path restarts its phase somewhere else than four
  // separate edges do, and the two would disagree at every corner.
  const { xPx, yPx, widthPx, heightPx } = rect;
  const edges = [
    { x1Px: xPx, y1Px: yPx, x2Px: xPx + widthPx, y2Px: yPx },
    { x1Px: xPx + widthPx, y1Px: yPx, x2Px: xPx + widthPx, y2Px: yPx + heightPx },
    { x1Px: xPx + widthPx, y1Px: yPx + heightPx, x2Px: xPx, y2Px: yPx + heightPx },
    { x1Px: xPx, y1Px: yPx + heightPx, x2Px: xPx, y2Px: yPx },
  ];
  for (const edge of edges) {
    paintSegment(context, stroke, edge);
  }
};

/**
 * Left edge of every code point, from the ordering the display list pins
 * down. The mirroring rule is shared with the DOM backend rather than stated
 * twice: two backends agreeing by coincidence is the divergence the display
 * list exists to prevent.
 */
const glyphOrigins = (run: DisplayGlyphRun): readonly number[] =>
  glyphCellOffsetsPx(run).map((offsetPx) => run.xPx + offsetPx);

/**
 * A maximal consecutive stretch of one run served by a single font resource.
 * A face split into script subsets embeds as several fonts, each with its own
 * glyph space, so a run crossing from one subset to the other has to change
 * font mid-run; the advances are untouched, because every span still carries
 * the display list's own numbers for its own code points.
 */
type GlyphSpan = {
  readonly resourceIndex: number;
  /** Index into the run's code points, for the span's text-matrix origin. */
  readonly startIndex: number;
  readonly glyphs: PositionedGlyph[];
};

const paintGlyphRun = (context: PaintContext, run: DisplayGlyphRun) => {
  const codePoints = [...run.text];
  // A run at zero size has zero extent: there is nothing to paint, and the
  // adjustment arithmetic below divides by the size.
  if (codePoints.length === 0 || run.fontSizePx <= 0) {
    return;
  }
  const font =
    context.fonts.get(run.font) ?? panic(`glyph run refers to unprepared font ${String(run.font)}`);
  const origins = glyphOrigins(run);
  const { stream } = context;

  beginPrimitive(context, context.alpha * run.color.a);
  stream.setFillColor(run.color);
  if (run.stroke !== undefined) {
    stream.setStrokeColor(run.stroke.color);
    stream.setLineWidth(run.stroke.thicknessPx);
  }
  stream.beginText();

  switch (font.kind) {
    case "embedded": {
      const spans: GlyphSpan[] = [];
      for (const [index, codePoint] of codePoints.entries()) {
        const { resourceIndex, glyphId, widthUnits } = font.glyphFor(codePoint.codePointAt(0) ?? 0);
        const origin = origins[index] ?? run.xPx;
        const next = origins[index + 1];
        const advance = run.advancesPx[index] ?? 0;
        // The measurer's advance wins over the font's own, always: line
        // breaking and pagination were decided on these numbers, so a page
        // positioned on the font's metrics is a page the engine never laid
        // out. The correction is against the width *this file* declares for
        // the glyph, which is what a reader will actually advance by.
        const delta = next === undefined ? advance : next - origin;
        const glyph = {
          glyphId,
          adjustment: widthUnits - (delta * TEXT_SPACE_UNITS_PER_EM) / run.fontSizePx,
        };
        const open = spans.at(-1);
        if (open === undefined || open.resourceIndex !== resourceIndex) {
          spans.push({ resourceIndex, startIndex: index, glyphs: [glyph] });
        } else {
          open.glyphs.push(glyph);
        }
      }
      for (const [spanIndex, span] of spans.entries()) {
        stream.setFont(span.resourceIndex, run.fontSizePx);
        // The render mode is text state, not font state: one setting covers
        // every span of the text object.
        if (spanIndex === 0 && run.stroke !== undefined) {
          stream.setTextRenderMode(TEXT_RENDER_MODE.fillThenStroke);
        }
        // The text matrix flips y back: the page CTM already turned the page
        // upside down so that display-list coordinates work, and glyphs must
        // not come along for that ride.
        stream.setTextMatrix([1, 0, 0, -1, origins[span.startIndex] ?? run.xPx, run.baselineYPx]);
        stream.showGlyphs(span.glyphs);
      }
      break;
    }
    case "standard": {
      stream.setFont(font.resourceIndex, run.fontSizePx);
      if (run.stroke !== undefined) {
        stream.setTextRenderMode(TEXT_RENDER_MODE.fillThenStroke);
      }
      // A base-14 stand-in has no width table here, so every code point is
      // positioned outright rather than corrected against an advance nobody
      // in this process knows.
      for (const [index, codePoint] of codePoints.entries()) {
        stream.setTextMatrix([1, 0, 0, -1, origins[index] ?? run.xPx, run.baselineYPx]);
        stream.showBytes([font.byteFor(codePoint.codePointAt(0) ?? 0)]);
      }
      break;
    }
    default: {
      const unreachable: never = font;
      panic(`unhandled prepared font: ${JSON.stringify(unreachable)}`);
    }
  }

  stream.endText();
  stream.restore();
};

const paintImage = (context: PaintContext, primitive: DisplayImagePrimitive) => {
  const { rect, crop, opacity } = primitive;
  const { stream } = context;
  beginPrimitive(context, context.alpha * opacity);
  if (crop === undefined) {
    // The unit square's v axis runs bottom-to-top through the image, so the
    // height is negated to land the image's first row at the box's top edge
    // in this y-down space.
    stream.concat([rect.widthPx, 0, 0, -rect.heightPx, rect.xPx, rect.yPx + rect.heightPx]);
    stream.drawImage(primitive.image);
    stream.restore();
    return;
  }
  const visibleWidth = 1 - crop.l - crop.r;
  const visibleHeight = 1 - crop.t - crop.b;
  if (visibleWidth <= 0 || visibleHeight <= 0) {
    // A crop that keeps nothing paints nothing.
    stream.restore();
    return;
  }
  stream.appendRect(rect.xPx, rect.yPx, rect.widthPx, rect.heightPx);
  stream.clipToCurrentPath();
  const fullWidth = rect.widthPx / visibleWidth;
  const fullHeight = rect.heightPx / visibleHeight;
  const left = rect.xPx - crop.l * fullWidth;
  const top = rect.yPx - crop.t * fullHeight;
  stream.concat([fullWidth, 0, 0, -fullHeight, left, top + fullHeight]);
  stream.drawImage(primitive.image);
  stream.restore();
};

const paintPrimitive = (context: PaintContext, primitive: DisplayPrimitive) => {
  switch (primitive.kind) {
    case "glyphRun":
      paintGlyphRun(context, primitive);
      return;
    case "rect":
      paintRect(context, primitive);
      return;
    case "line":
      paintSegment(context, primitive.stroke, {
        x1Px: primitive.x1Px,
        y1Px: primitive.y1Px,
        x2Px: primitive.x2Px,
        y2Px: primitive.y2Px,
      });
      return;
    case "image":
      paintImage(context, primitive);
      return;
    case "clipGroup": {
      context.stream.save();
      context.stream.appendRect(
        primitive.rect.xPx,
        primitive.rect.yPx,
        primitive.rect.widthPx,
        primitive.rect.heightPx,
      );
      context.stream.clipToCurrentPath();
      for (const child of primitive.children) {
        paintPrimitive(context, child);
      }
      context.stream.restore();
      return;
    }
    case "rotateGroup": {
      context.stream.save();
      context.stream.concat(
        rotationMatrix(primitive.degrees, primitive.originXPx, primitive.originYPx),
      );
      for (const child of primitive.children) {
        paintPrimitive(context, child);
      }
      context.stream.restore();
      return;
    }
    case "opacityGroup": {
      const alpha = context.alpha * primitive.opacity;
      context.stream.save();
      context.stream.setAlpha(alpha);
      const inner = withAlpha(context, alpha);
      for (const child of primitive.children) {
        paintPrimitive(inner, child);
      }
      context.stream.restore();
      return;
    }
    default: {
      const unreachable: never = primitive;
      panic(`unhandled display primitive: ${JSON.stringify(unreachable)}`);
    }
  }
};

export type PaintPageOptions = {
  readonly page: DisplayPage;
  readonly fonts: ReadonlyMap<number, PreparedFont>;
};

export const paintPage = ({ page, fonts }: PaintPageOptions): readonly ContentPart[] => {
  const stream = createContentStream();
  // The one conversion from display-list space to PDF space, at the stream's
  // first operator. Everything after it is written in CSS px, y down.
  stream.concat(basePageMatrix(page.heightPx));
  const context: PaintContext = { stream, fonts, alpha: OPAQUE };
  for (const primitive of page.primitives) {
    paintPrimitive(context, primitive);
  }
  return stream.parts();
};

/** What the file must carry, gathered before anything is written. */
export type DisplayUsage = {
  readonly codePointsByFont: ReadonlyMap<number, ReadonlySet<number>>;
  readonly imageIndices: ReadonlySet<number>;
};

type UsageAccumulator = {
  readonly codePointsByFont: Map<number, Set<number>>;
  readonly imageIndices: Set<number>;
};

const collectPrimitive = (
  primitive: DisplayPrimitive,
  fontCount: number,
  imageCount: number,
  into: UsageAccumulator,
): PaintError | null => {
  switch (primitive.kind) {
    case "glyphRun": {
      if (primitive.font < 0 || primitive.font >= fontCount) {
        return new PaintError({
          message: `glyph run refers to font ${String(primitive.font)}, outside the font table`,
        });
      }
      const codePoints = [...primitive.text];
      if (codePoints.length !== primitive.advancesPx.length) {
        return new PaintError({
          message: `glyph run carries ${String(primitive.advancesPx.length)} advances for ${String(codePoints.length)} code points`,
        });
      }
      const seen = into.codePointsByFont.get(primitive.font) ?? new Set<number>();
      for (const codePoint of codePoints) {
        seen.add(codePoint.codePointAt(0) ?? 0);
      }
      into.codePointsByFont.set(primitive.font, seen);
      return null;
    }
    case "rect":
    case "line":
      return null;
    case "image":
      if (primitive.image < 0 || primitive.image >= imageCount) {
        return new PaintError({
          message: `image primitive refers to image ${String(primitive.image)}, outside the image table`,
        });
      }
      into.imageIndices.add(primitive.image);
      return null;
    case "clipGroup":
    case "rotateGroup":
    case "opacityGroup": {
      for (const child of primitive.children) {
        const error = collectPrimitive(child, fontCount, imageCount, into);
        if (error !== null) {
          return error;
        }
      }
      return null;
    }
    default: {
      const unreachable: never = primitive;
      return panic(`unhandled display primitive: ${JSON.stringify(unreachable)}`);
    }
  }
};

/**
 * Walks the pages once before anything is written, because subsetting needs
 * every glyph the document uses before the first one is emitted.
 */
export const collectUsage = (
  pages: readonly DisplayPage[],
  fontCount: number,
  imageCount: number,
): Result<DisplayUsage, PaintError> => {
  const into: UsageAccumulator = { codePointsByFont: new Map(), imageIndices: new Set() };
  for (const page of pages) {
    for (const primitive of page.primitives) {
      const error = collectPrimitive(primitive, fontCount, imageCount, into);
      if (error !== null) {
        return Result.err(error);
      }
    }
  }
  return Result.ok(into);
};
