/**
 * Paragraph fragments → primitives.
 *
 * This is where the display list earns its keep. The DOM painter never
 * computes a line's start x: it sets `text-align` on the fragment, a
 * `text-indent` on the first line, margins for float offsets, and lets the
 * browser place the glyphs (`renderParagraph.ts:2957` onwards). A PDF has no
 * such engine, so alignment, first-line and hanging indent, list-marker
 * footprint, float offsets and `floatSkipBefore` are all resolved here into one
 * number per line, and justification is resolved into the advances themselves
 * rather than deferred to a backend.
 *
 * Two cursors run side by side. `layoutX` is the measurer's content-area
 * coordinate, which is what tab stops are measured from and what line breaking
 * was decided on; `paintX` is where the glyphs actually land after alignment
 * moved the line. They differ by a constant per line, and conflating them would
 * put a right-aligned line's tabs on the wrong stops.
 */

import {
  getListMarkerInlineWidth,
  getListMarkerVisualOffset,
  resolveListMarkerFont,
} from "../../layout-engine/measure/listMarkerWidth";
import {
  buildRunFontStyle,
  DEFAULT_FONT_FAMILY,
  DEFAULT_FONT_SIZE,
  ptToPx,
} from "../../layout-engine/measure/measureHelpers";
import { getFontMetrics, measureTextWidth } from "../../layout-engine/measure/measureProvider";
import type { FontStyle } from "../../layout-engine/measure/measureTypes";
import {
  calculateTabWidth,
  pixelsToTwips,
  type TabContext,
  type TabStop as TabCalcStop,
} from "../../layout-engine/measure/tabCalculator";
import {
  countCompressibleSpaces,
  toPaintedText,
} from "../../layout-engine/measure/textMeasurementPolicy";
import type {
  BorderStyle,
  FieldRun,
  MathRun,
  MeasuredLine,
  ParagraphBlock,
  ParagraphBorders,
  ParagraphFragment,
  ParagraphMeasure,
  Run,
  TabRun,
  TabStop,
  TextRun,
} from "../../layout-engine/types";
import { resolveParagraphBorderHorizontalOutsets } from "../../layout-painter/borderStroke";
import { getAutomaticTextColorForBackground } from "../../layout-painter/documentColors";
import {
  getLeaderChar,
  getRenderableTextColor,
  sliceRunsForLine,
  splitCollapsibleLineEdgeSpaces,
  splitTextRunsByEastAsia,
  startsAfterSoftWrap,
} from "../../layout-painter/renderParagraph";
import { isFloatingImageRun } from "../../layout-painter/renderUtils";
import { getHorizontalScaleFactor } from "../../utils/horizontalScale";
import { resolvePhysicalParagraphInlineLayout } from "../../utils/paragraphInlineLayout";
import { inlineImageBoundingBox } from "../../utils/rotationBoundingBox";
import { sanitizeExternalUrl } from "../../utils/urlSecurity";
import type {
  DisplayColor,
  DisplayGlyphRun,
  DisplayLine,
  DisplayLinkTarget,
  DisplayPrimitive,
  DisplayRect,
  DisplayStroke,
} from "../types";
import { type BuildContext, trackedChangeColor } from "./buildContext";
import { DOC_CANVAS_TEXT, parseDisplayColor } from "./colors";
import { buildGlyphs, type Glyphs } from "./glyphs";
import { paintImage } from "./imagePrimitives";
import { decorationPatternForStyle, resolveBorderStroke } from "./strokes";
import {
  strikethroughCenterYPx,
  underlineCenterYPx,
  underlineThicknessPx,
} from "./textDecorations";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

/** Word's default hyperlink colour, applied when the run carries no colour of its own. */
const HYPERLINK_COLOR: DisplayColor = { r: 0x05, g: 0x63, b: 0xc1, a: 1 };

/** `renderParagraph.ts:603-616`: a raised run paints 0.4em up, a lowered one 0.2em down. */
const SUPERSCRIPT_RISE_RATIO = 0.4;
const SUBSCRIPT_DROP_RATIO = 0.2;

/** Matches `RIGHT_EDGE_EPSILON_PX` in the painter and `WIDTH_TOLERANCE` in the measurer. */
const WIDTH_EPSILON_PX = 0.5;

/** `w:pBdr` bar borders hang this far left of the text (`renderParagraph.ts:3072`). */
const BAR_BORDER_OFFSET_PX = 8;

const AUTOMATIC_TEXT_COLOR_VALUES = new Set(["auto", "windowtext"]);

/** `w:color w:val="auto"` and its `windowtext` spelling, normalized as the painter normalizes them. */
const isAutomaticTextColor = (color: string): boolean =>
  AUTOMATIC_TEXT_COLOR_VALUES.has(color.trim().toLowerCase().replace(/^#/u, ""));

const isNoteReferenceRun = (run: TextRun | TabRun): boolean =>
  run.footnoteRefId !== undefined || run.endnoteRefId !== undefined;

type PaintableRun = TextRun | FieldRun | MathRun;

/**
 * The colour a run's glyphs are actually painted in.
 *
 * `applyRunStyles` (`renderParagraph.ts:250-258`) sets no inline `color` for a
 * black or `auto` run, on purpose: the editor's `--doc-canvas-text` then adapts
 * to dark mode. A backend with no stylesheet would paint nothing at all, so
 * black is re-materialized here. This is the one place that knows the dropped
 * value was black; drop the fallback and every default-coloured run in the
 * export turns invisible.
 */
const resolveRunColor = (run: TextRun, context: BuildContext): DisplayColor => {
  const explicit = getRenderableTextColor(run);
  let color = explicit === undefined ? undefined : parseDisplayColor(explicit);

  const background = run.highlight ?? run.shading;
  const isTracked = run.isInsertion === true || run.isDeletion === true;
  const hasComments = run.commentIds !== undefined && run.commentIds.length > 0;
  if (background && explicit === undefined && !isTracked && !hasComments) {
    color = parseDisplayColor(getAutomaticTextColorForBackground(background));
  }
  if (isTracked) {
    color = trackedChangeColor(context.authorColors, run.changeAuthor, run.isSuggestion);
  }
  if (run.hyperlink && run.hyperlink.noDefaultStyle !== true) {
    const direct = run.color?.trim();
    const directWins = direct && !isAutomaticTextColor(direct) && run.textColorSource === "direct";
    color = (directWins ? parseDisplayColor(direct) : undefined) ?? color ?? HYPERLINK_COLOR;
  }

  return color ?? DOC_CANVAS_TEXT;
};

const runFontStyle = (run: PaintableRun | TabRun): FontStyle =>
  buildRunFontStyle(run, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE);

/**
 * Field text, resolved without a clock. `evaluateFieldInstruction` needs a
 * `now`, which a pure builder cannot supply, so the cached `w:fldSimple` result
 * Word already wrote wins; only the two pagination fields, whose values the
 * layout itself carries, are recomputed.
 */
const resolveFieldText = (run: FieldRun, context: BuildContext): string => {
  if (run.fieldType === "PAGE") {
    return String(context.pageNumber);
  }
  if (run.fieldType === "NUMPAGES") {
    return String(context.totalPages);
  }
  return run.fallback ?? "";
};

const paintableText = (run: PaintableRun, context: BuildContext): string => {
  switch (run.kind) {
    case "text":
      return toPaintedText(run.text);
    case "field":
      return toPaintedText(resolveFieldText(run, context));
    case "math":
      return toPaintedText(run.plainText || "[equation]");
  }
};

const convertTabStop = (stop: TabStop): TabCalcStop => ({
  val: stop.val,
  pos: stop.pos,
  ...(stop.leader === undefined ? {} : { leader: stop.leader }),
});

type LineSink = {
  readonly backgrounds: DisplayPrimitive[];
  readonly glyphs: DisplayPrimitive[];
  readonly decorations: DisplayPrimitive[];
};

const horizontalLine = (
  xPx: number,
  widthPx: number,
  yPx: number,
  stroke: DisplayStroke,
): DisplayLine => ({
  kind: "line",
  x1Px: xPx,
  y1Px: yPx,
  x2Px: xPx + widthPx,
  y2Px: yPx,
  stroke,
});

export type ParagraphPaintOptions = {
  readonly fragment: ParagraphFragment;
  readonly block: ParagraphBlock;
  readonly measure: ParagraphMeasure;
  readonly context: BuildContext;
  /** Borders of the adjacent paragraphs, for ECMA-376 §17.3.1.24 border grouping. */
  readonly prevBorders?: ParagraphBorders;
  readonly nextBorders?: ParagraphBorders;
};

const bordersEqual = (a?: BorderStyle, b?: BorderStyle): boolean => {
  if (!a && !b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }
  return a.style === b.style && a.width === b.width && a.color === b.color;
};

const bordersFormGroup = (a?: ParagraphBorders, b?: ParagraphBorders): boolean =>
  Boolean(a) &&
  Boolean(b) &&
  bordersEqual(a?.top, b?.top) &&
  bordersEqual(a?.bottom, b?.bottom) &&
  bordersEqual(a?.left, b?.left) &&
  bordersEqual(a?.right, b?.right) &&
  bordersEqual(a?.between, b?.between);

/** Paint the paragraph's own background and border box, before any line. */
const paintParagraphChrome = ({
  fragment,
  block,
  context,
  prevBorders,
  nextBorders,
}: ParagraphPaintOptions): DisplayPrimitive[] => {
  const primitives: DisplayPrimitive[] = [];
  const { indentLeft, indentRight } = resolvePhysicalParagraphInlineLayout(block);

  const shading = block.attrs?.shading;
  if (shading) {
    const fill = parseDisplayColor(shading);
    if (fill) {
      primitives.push({
        kind: "rect",
        rect: {
          xPx: fragment.x,
          yPx: fragment.y,
          widthPx: fragment.width,
          heightPx: fragment.height,
        },
        fill,
      });
    } else {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.unresolvedColor,
        context.pageIndex,
        `paragraph shading ${shading}`,
      );
    }
  }

  const borders = block.attrs?.borders;
  if (!borders) {
    return primitives;
  }

  const topBorder = bordersFormGroup(prevBorders, borders) ? borders.between : borders.top;
  const bottomBorder = bordersFormGroup(borders, nextBorders) ? undefined : borders.bottom;
  const outsets = resolveParagraphBorderHorizontalOutsets(
    borders,
    topBorder !== undefined || bottomBorder !== undefined,
  );

  const boxLeft = fragment.x + indentLeft - outsets.left;
  const boxRight = fragment.x + fragment.width - (indentRight - outsets.right);
  const boxTop = fragment.y - (topBorder?.space ?? 0) - (topBorder?.width ?? 0);
  const boxBottom =
    fragment.y + fragment.height + (bottomBorder?.space ?? 0) + (bottomBorder?.width ?? 0);

  const push = (border: BorderStyle | undefined, line: (stroke: DisplayStroke) => DisplayLine) => {
    if (!border) {
      return;
    }
    const { stroke, unresolvedColor } = resolveBorderStroke(border);
    if (unresolvedColor !== undefined) {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.unresolvedColor,
        context.pageIndex,
        `paragraph border colour ${unresolvedColor}`,
      );
    }
    if (stroke) {
      primitives.push(line(stroke));
    }
  };

  // `box-sizing: border-box` puts the stroke inside the border box, and a
  // display-list stroke is centred on its path, so each edge sits half a
  // thickness inside the box it belongs to.
  push(topBorder, (stroke) => ({
    kind: "line",
    x1Px: boxLeft,
    y1Px: boxTop + stroke.thicknessPx / 2,
    x2Px: boxRight,
    y2Px: boxTop + stroke.thicknessPx / 2,
    stroke,
  }));
  push(bottomBorder, (stroke) => ({
    kind: "line",
    x1Px: boxLeft,
    y1Px: boxBottom - stroke.thicknessPx / 2,
    x2Px: boxRight,
    y2Px: boxBottom - stroke.thicknessPx / 2,
    stroke,
  }));
  push(borders.left, (stroke) => ({
    kind: "line",
    x1Px: boxLeft + stroke.thicknessPx / 2,
    y1Px: boxTop,
    x2Px: boxLeft + stroke.thicknessPx / 2,
    y2Px: boxBottom,
    stroke,
  }));
  push(borders.right, (stroke) => ({
    kind: "line",
    x1Px: boxRight - stroke.thicknessPx / 2,
    y1Px: boxTop,
    x2Px: boxRight - stroke.thicknessPx / 2,
    y2Px: boxBottom,
    stroke,
  }));
  push(borders.bar, (stroke) => ({
    kind: "line",
    x1Px: fragment.x - BAR_BORDER_OFFSET_PX + stroke.thicknessPx / 2,
    y1Px: fragment.y,
    x2Px: fragment.x - BAR_BORDER_OFFSET_PX + stroke.thicknessPx / 2,
    y2Px: fragment.y + fragment.height,
    stroke,
  }));

  return primitives;
};

type LineGeometry = {
  /** Where the first glyph lands, page-absolute, after alignment. */
  readonly paintStartXPx: number;
  /**
   * The same line's start in the measurer's coordinates, whose origin is the
   * fragment's own left edge. Tab stops are twips from there, so a tab must be
   * computed in this space and only then shifted by the alignment.
   */
  readonly layoutStartXPx: number;
  /** Rightmost x inline content may occupy, in the measurer's coordinates. */
  readonly layoutRightEdgeXPx: number;
  readonly baselineYPx: number;
  readonly lineTopYPx: number;
};

type ResolveLineGeometryOptions = {
  readonly fragment: ParagraphFragment;
  readonly line: MeasuredLine;
  readonly lineTopYPx: number;
  readonly alignment: "left" | "center" | "right" | "justify";
  readonly isRtl: boolean;
  readonly indentLeft: number;
  readonly indentRight: number;
  /** `(w:firstLine - w:hanging)`, the measurer's first-line offset. */
  readonly firstLineOffsetPx: number;
  readonly markerInlineWidthPx: number;
  readonly isFirstLine: boolean;
};

const resolveLineGeometry = ({
  fragment,
  line,
  lineTopYPx,
  alignment,
  isRtl,
  indentLeft,
  indentRight,
  firstLineOffsetPx,
  markerInlineWidthPx,
  isFirstLine,
}: ResolveLineGeometryOptions): LineGeometry => {
  const leftOffset = line.leftOffset ?? 0;
  const rightOffset = line.rightOffset ?? 0;
  const firstLineExtra = isFirstLine ? firstLineOffsetPx + markerInlineWidthPx : 0;

  // The measurer accumulates from the fragment's left edge whatever the
  // alignment, so tab stops stay on the document's grid. Both `measureParagraph`
  // and `renderLine` use that origin, not the page's.
  const layoutStartXPx = indentLeft + leftOffset + (isRtl ? 0 : firstLineExtra);
  const layoutRightEdgeXPx = fragment.width - indentRight - rightOffset;

  const leftEdge = fragment.x + layoutStartXPx;
  const rightEdge = fragment.x + layoutRightEdgeXPx - (isRtl ? firstLineExtra : 0);

  let paintStartXPx: number;
  switch (alignment) {
    case "right":
      paintStartXPx = rightEdge - line.width;
      break;
    case "center":
      paintStartXPx = leftEdge + (rightEdge - leftEdge - line.width) / 2;
      break;
    case "left":
    case "justify":
      paintStartXPx = isRtl ? rightEdge - line.width : leftEdge;
      break;
  }

  // CSS centres the inline box in its line box: half the leading sits above the
  // font's ascent, half below its descent.
  const baselineYPx =
    lineTopYPx + (line.lineHeight - (line.ascent + line.descent)) / 2 + line.ascent;

  return { paintStartXPx, layoutStartXPx, layoutRightEdgeXPx, baselineYPx, lineTopYPx };
};

type JustificationPlan = {
  /** Added to every compressible space's advance. Negative contracts the line. */
  readonly spaceDeltaPx: number;
};

type ResolveJustificationOptions = {
  readonly line: MeasuredLine;
  readonly alignment: "left" | "center" | "right" | "justify";
  readonly availableWidthPx: number;
  readonly firstLineOffsetPx: number;
  readonly indentLeft: number;
  readonly hasVisibleMarker: boolean;
  readonly isFirstLine: boolean;
  readonly isLastLine: boolean;
  readonly paragraphEndsWithLineBreak: boolean;
  readonly shrinkableSpaces: number;
};

/**
 * Resolve `renderParagraph.ts:2404-2437` into one number.
 *
 * The painter has three branches (explicit contraction via `word-spacing`, an
 * explicit expansion for tab-bearing lines, and CSS `text-align: justify` for
 * the rest) that all mean the same thing: spread the difference between the
 * measured line and its capacity across the compressible spaces. A display list
 * has no CSS branch, so all three collapse to the delta.
 */
const resolveJustification = ({
  line,
  alignment,
  availableWidthPx,
  firstLineOffsetPx,
  indentLeft,
  hasVisibleMarker,
  isFirstLine,
  isLastLine,
  paragraphEndsWithLineBreak,
  shrinkableSpaces,
}: ResolveJustificationOptions): JustificationPlan => {
  if (alignment !== "justify" || shrinkableSpaces === 0) {
    return { spaceDeltaPx: 0 };
  }

  const firstLineIndentPx = isFirstLine ? firstLineOffsetPx : 0;
  const positiveIndentPx = Math.max(0, firstLineIndentPx);
  const hangingPx = Math.max(0, -firstLineIndentPx);
  // A marker's negative margin cancels the part of the hanging slot outside the
  // paragraph, so only the portion inside the content edge can expand the line.
  const hangingExpansionPx = hasVisibleMarker
    ? Math.min(hangingPx, Math.max(0, indentLeft))
    : hangingPx;
  const capacityPx = availableWidthPx - positiveIndentPx + hangingExpansionPx;
  const overfullPx = line.width - capacityPx;

  const finalContractionPx =
    line.justificationPaint?.type === "space-contraction"
      ? line.justificationPaint.contractionPx
      : undefined;
  const compressFinalLine =
    isLastLine &&
    finalContractionPx !== undefined &&
    finalContractionPx > WIDTH_EPSILON_PX &&
    overfullPx > WIDTH_EPSILON_PX;

  if (compressFinalLine) {
    return { spaceDeltaPx: -(finalContractionPx ?? 0) / shrinkableSpaces };
  }
  if (!isLastLine || paragraphEndsWithLineBreak) {
    if (Math.abs(overfullPx) <= WIDTH_EPSILON_PX) {
      return { spaceDeltaPx: 0 };
    }
    return { spaceDeltaPx: -overfullPx / shrinkableSpaces };
  }
  return { spaceDeltaPx: 0 };
};

/**
 * The editable-model range a run's glyphs came from, or `undefined` when they
 * have no counterpart there. Only body runs qualify: a header, footer or
 * footnote run's positions belong to that story's own document, so handing
 * them out as model positions would address unrelated body content.
 */
const modelRangeOf = (run: TextRun, context: BuildContext): DisplayGlyphRun["pmRange"] =>
  context.story !== "body" || run.pmStart === undefined || run.pmEnd === undefined
    ? undefined
    : { start: run.pmStart, end: run.pmEnd };

type EmitGlyphRunOptions = {
  readonly sink: LineSink;
  readonly context: BuildContext;
  readonly run: TextRun;
  readonly glyphs: Glyphs;
  readonly style: FontStyle;
  readonly paintXPx: number;
  readonly baselineYPx: number;
  readonly lineTopYPx: number;
  readonly lineHeightPx: number;
  readonly isRtl: boolean;
  /**
   * Absent for glyphs the model does not contain: a list marker, a substituted
   * field value, a tab leader. An approximate range is worse than none, since
   * an editing surface maps clicks and selections through it.
   */
  readonly pmRange?: DisplayGlyphRun["pmRange"];
};

/**
 * One `glyphRun` plus everything painted around it: the run's background rect
 * first, then the glyphs, then the decorations whose geometry CSS would have
 * derived from the font.
 */
const emitGlyphRun = ({
  sink,
  context,
  run,
  glyphs,
  style,
  paintXPx,
  baselineYPx,
  lineTopYPx,
  lineHeightPx,
  isRtl,
  pmRange,
}: EmitGlyphRunOptions): void => {
  if (glyphs.text.length === 0) {
    return;
  }

  const fontSizePx = ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE);
  const color = resolveRunColor(run, context);

  let runBaselineYPx = baselineYPx - (run.positionPx ?? 0);
  if (run.superscript) {
    runBaselineYPx -= fontSizePx * SUPERSCRIPT_RISE_RATIO;
  } else if (run.subscript) {
    runBaselineYPx += fontSizePx * SUBSCRIPT_DROP_RATIO;
  }

  const background = run.highlight ?? run.shading;
  if (background && glyphs.widthPx > 0) {
    const fill = parseDisplayColor(background);
    if (fill) {
      // A CSS inline box is the font box, not the ink extent, so a highlight
      // covers the whole em band rather than just where glyphs happen to reach.
      const metrics = getFontMetrics(style);
      sink.backgrounds.push({
        kind: "rect",
        rect: {
          xPx: paintXPx,
          yPx: runBaselineYPx - metrics.fontBoxAscent,
          widthPx: glyphs.widthPx,
          heightPx: metrics.fontBoxAscent + metrics.fontBoxDescent,
        },
        fill,
      });
    } else {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.unresolvedColor,
        context.pageIndex,
        `run background ${background}`,
      );
    }
  }

  const font = context.fonts.intern({
    fontFamily: style.fontFamily ?? DEFAULT_FONT_FAMILY,
    ...(style.alternateFontFamily === undefined
      ? {}
      : { alternateFontFamily: style.alternateFontFamily }),
    ...(style.bold === undefined ? {} : { bold: style.bold }),
    ...(style.italic === undefined ? {} : { italic: style.italic }),
    measureStyle: style,
  });

  // `w:outline` paints the glyph as an outline; the painter emits a 1px
  // `-webkit-text-stroke` in the run's own colour.
  const stroke: DisplayStroke | undefined = run.textOutline
    ? { color, thicknessPx: 1, pattern: "solid" }
    : undefined;

  sink.glyphs.push({
    kind: "glyphRun",
    font,
    fontSizePx,
    color,
    xPx: paintXPx,
    baselineYPx: runBaselineYPx,
    text: glyphs.text,
    advancesPx: glyphs.advancesPx,
    direction: isRtl ? "rtl" : "ltr",
    ...(stroke === undefined ? {} : { stroke }),
    ...(pmRange === undefined ? {} : { pmRange }),
  });

  emitDecorations({
    sink,
    run,
    color,
    context,
    xPx: paintXPx,
    widthPx: glyphs.widthPx,
    baselineYPx: runBaselineYPx,
    fontSizePx,
  });

  if (run.hyperlink) {
    const target = resolveLinkTarget(run.hyperlink.href, context);
    if (target) {
      context.links.push({
        rect: {
          xPx: paintXPx,
          yPx: lineTopYPx,
          widthPx: glyphs.widthPx,
          heightPx: lineHeightPx,
        },
        target,
        ...(run.hyperlink.tooltip === undefined ? {} : { tooltip: run.hyperlink.tooltip }),
      });
    }
  }

  reportRunEffects(run, context);
};

type EmitDecorationsOptions = {
  readonly sink: LineSink;
  readonly run: TextRun;
  readonly color: DisplayColor;
  readonly context: BuildContext;
  readonly xPx: number;
  readonly widthPx: number;
  readonly baselineYPx: number;
  readonly fontSizePx: number;
};

const emitDecorations = ({
  sink,
  run,
  color,
  context,
  xPx,
  widthPx,
  baselineYPx,
  fontSizePx,
}: EmitDecorationsOptions): void => {
  if (widthPx <= 0) {
    return;
  }
  const thicknessPx = underlineThicknessPx(fontSizePx);
  const trackedColor =
    run.isInsertion || run.isDeletion
      ? trackedChangeColor(context.authorColors, run.changeAuthor, run.isSuggestion)
      : undefined;

  const wantsUnderline = Boolean(run.underline) && !isNoteReferenceRun(run);
  if (wantsUnderline || run.isInsertion) {
    const authored = typeof run.underline === "object" ? run.underline : undefined;
    const authoredColor =
      authored?.color === undefined ? undefined : parseDisplayColor(authored.color);
    // A suggested insertion strokes dotted; an author's insertion strokes solid
    // in the author's hue (`renderParagraph.ts:513-545`).
    const pattern = run.isSuggestion ? "dotted" : decorationPatternForStyle(authored?.style);
    sink.decorations.push(
      horizontalLine(xPx, widthPx, underlineCenterYPx(baselineYPx, fontSizePx), {
        color: trackedColor ?? authoredColor ?? color,
        thicknessPx,
        pattern,
      }),
    );
  }

  if (run.strike || run.isDeletion) {
    sink.decorations.push(
      horizontalLine(xPx, widthPx, strikethroughCenterYPx(baselineYPx, fontSizePx), {
        color: trackedColor ?? color,
        thicknessPx,
        pattern: run.isSuggestion ? "dotted" : "solid",
      }),
    );
  }
};

/**
 * Run-level typography the display list cannot express. The glyphs and their
 * advances are still correct; only the treatment is missing, so this reports
 * rather than skips.
 */
const reportRunEffects = (run: TextRun, context: BuildContext): void => {
  const { unsupported, pageIndex } = context;
  if (run.smallCaps) {
    unsupported.report(
      UNSUPPORTED_CONSTRUCT.smallCaps,
      pageIndex,
      "w:smallCaps advances are measured, but the display list carries no small-cap glyph selection",
    );
  }
  if (getHorizontalScaleFactor(run.horizontalScale) !== 1) {
    unsupported.report(
      UNSUPPORTED_CONSTRUCT.horizontalScale,
      pageIndex,
      `w:w ${String(run.horizontalScale)}% is folded into the advances, but the glyphs are not narrowed`,
    );
  }
  if (run.textEffect) {
    unsupported.report(UNSUPPORTED_CONSTRUCT.textEffect, pageIndex, `w:effect ${run.textEffect}`);
  }
  if (run.emphasisMark) {
    unsupported.report(UNSUPPORTED_CONSTRUCT.emphasisMark, pageIndex, `w:em ${run.emphasisMark}`);
  }
  if (run.emboss || run.imprint || run.textShadow) {
    unsupported.report(
      UNSUPPORTED_CONSTRUCT.textShadowEffect,
      pageIndex,
      "w:emboss / w:imprint / w:shadow paint as plain glyphs",
    );
  }
};

const resolveLinkTarget = (href: string, context: BuildContext): DisplayLinkTarget | undefined => {
  if (href.startsWith("#")) {
    return context.bookmarkTargets.get(href.slice(1));
  }
  const safe = sanitizeExternalUrl(href);
  return safe === undefined ? undefined : { kind: "external", href: safe };
};

type PaintLineOptions = {
  readonly block: ParagraphBlock;
  readonly line: MeasuredLine;
  readonly geometry: LineGeometry;
  readonly context: BuildContext;
  readonly alignment: "left" | "center" | "right" | "justify";
  readonly isRtl: boolean;
  readonly isFirstLine: boolean;
  readonly isLastLine: boolean;
  readonly paragraphEndsWithLineBreak: boolean;
  readonly availableWidthPx: number;
  readonly indentLeft: number;
  readonly firstLineOffsetPx: number;
  readonly markerInlineWidthPx: number;
};

const paintLine = ({
  block,
  line,
  geometry,
  context,
  alignment,
  isRtl,
  isFirstLine,
  isLastLine,
  paragraphEndsWithLineBreak,
  availableWidthPx,
  indentLeft,
  firstLineOffsetPx,
  markerInlineWidthPx,
}: PaintLineOptions): DisplayPrimitive[] => {
  const sink: LineSink = { backgrounds: [], glyphs: [], decorations: [] };
  // The painter's own splitter decides which line-edge spaces collapse, and it
  // splits a part-word part-space run so the collapsed span is a whole run.
  // Reusing it is the only way the editor and the export cannot come to
  // disagree about a space.
  const { runs, collapsedLeadingRuns, collapsedTrailingRuns } = splitCollapsibleLineEdgeSpaces(
    splitTextRunsByEastAsia(sliceRunsForLine(block, line)),
    startsAfterSoftWrap(block, line),
  );
  const isCollapsedEdgeRun = (run: TextRun): boolean =>
    collapsedLeadingRuns.has(run) || collapsedTrailingRuns.has(run);

  const hasVisibleMarker =
    Boolean(block.attrs?.listMarker) && block.attrs?.listMarkerHidden !== true;
  let shrinkableSpaces = 0;
  for (const run of runs) {
    if (run.kind === "text") {
      if (!isCollapsedEdgeRun(run)) {
        shrinkableSpaces += countCompressibleSpaces(run.text);
      }
    } else if (run.kind === "field") {
      shrinkableSpaces += countCompressibleSpaces(resolveFieldText(run, context));
    }
  }

  const { spaceDeltaPx } = resolveJustification({
    line,
    alignment,
    availableWidthPx,
    firstLineOffsetPx,
    indentLeft,
    hasVisibleMarker: hasVisibleMarker && isFirstLine,
    isFirstLine,
    isLastLine,
    paragraphEndsWithLineBreak,
    shrinkableSpaces,
  });

  if (isFirstLine && hasVisibleMarker) {
    paintListMarker({
      sink,
      block,
      context,
      // The marker occupies the slot immediately before the body text, which
      // `layoutStartXPx` already sits past.
      paintXPx: geometry.paintStartXPx - markerInlineWidthPx + getListMarkerVisualOffset(block),
      baselineYPx: geometry.baselineYPx,
      isRtl,
    });
  }

  let layoutXPx = geometry.layoutStartXPx;
  const alignShiftPx = geometry.paintStartXPx - geometry.layoutStartXPx;
  const paintXOf = (): number => layoutXPx + alignShiftPx;

  const tabContext: TabContext = {
    ...(block.attrs?.tabs === undefined
      ? {}
      : { explicitStops: block.attrs.tabs.map(convertTabStop) }),
    ...(block.attrs?.defaultTabStopTwips === undefined
      ? {}
      : { defaultTabInterval: block.attrs.defaultTabStopTwips }),
    // The tab grid is measured from the content-area left edge, so the authored
    // (logical) indent is what defines it, not the mirrored physical one.
    leftIndent: pixelsToTwips(block.attrs?.indent?.left ?? 0),
  };

  for (let index = 0; index < runs.length; index += 1) {
    // SAFETY: index < runs.length.
    const run = runs[index]!;
    switch (run.kind) {
      case "text": {
        if (run.hidden) {
          context.unsupported.report(
            UNSUPPORTED_CONSTRUCT.hiddenRun,
            context.pageIndex,
            "w:vanish text is suppressed, as Word's print view suppresses it",
          );
          break;
        }
        const style = runFontStyle(run);
        const glyphs = buildGlyphs({
          text: toPaintedText(run.text),
          style,
          allCaps: run.allCaps === true,
          spaceDeltaPx,
          collapsed: isCollapsedEdgeRun(run),
        });
        const pmRange = modelRangeOf(run, context);
        emitGlyphRun({
          sink,
          context,
          run,
          glyphs,
          style,
          paintXPx: paintXOf(),
          baselineYPx: geometry.baselineYPx,
          lineTopYPx: geometry.lineTopYPx,
          lineHeightPx: line.lineHeight,
          isRtl,
          ...(pmRange === undefined ? {} : { pmRange }),
        });
        layoutXPx += glyphs.widthPx;
        break;
      }
      case "field": {
        const style = runFontStyle(run);
        const glyphs = buildGlyphs({
          text: paintableText(run, context),
          style,
          allCaps: run.allCaps === true,
          spaceDeltaPx,
          collapsed: false,
        });
        // A field result carries the field run's whole `w:rPr`; painting it
        // through the text path is what keeps a footer page number in the
        // footer's font rather than the document default.
        emitGlyphRun({
          sink,
          context,
          run: { ...run, kind: "text", text: glyphs.text },
          glyphs,
          style,
          paintXPx: paintXOf(),
          baselineYPx: geometry.baselineYPx,
          lineTopYPx: geometry.lineTopYPx,
          lineHeightPx: line.lineHeight,
          isRtl,
        });
        layoutXPx += glyphs.widthPx;
        break;
      }
      case "math": {
        const style = runFontStyle(run);
        const glyphs = buildGlyphs({
          text: paintableText(run, context),
          style,
          allCaps: false,
          spaceDeltaPx: 0,
          collapsed: false,
        });
        emitGlyphRun({
          sink,
          context,
          run: { ...run, kind: "text", text: glyphs.text },
          glyphs,
          style,
          paintXPx: paintXOf(),
          baselineYPx: geometry.baselineYPx,
          lineTopYPx: geometry.lineTopYPx,
          lineHeightPx: line.lineHeight,
          isRtl,
        });
        context.unsupported.report(
          UNSUPPORTED_CONSTRUCT.mathRun,
          context.pageIndex,
          "OMML painted as renderMathFallback's plain text, not as typeset math",
        );
        layoutXPx += glyphs.widthPx;
        break;
      }
      case "tab": {
        layoutXPx += paintTab({
          sink,
          context,
          runs,
          index,
          run,
          tabContext,
          layoutXPx,
          paintXPx: paintXOf(),
          baselineYPx: geometry.baselineYPx,
          layoutRightEdgeXPx: geometry.layoutRightEdgeXPx,
        });
        break;
      }
      case "image": {
        if (isFloatingImageRun(run)) {
          break;
        }
        const boundingBox = inlineImageBoundingBox(run);
        // Word seats an inline image as a tall glyph on the text baseline.
        const primitives = paintImage({
          source: run,
          rect: {
            xPx: paintXOf(),
            yPx: geometry.baselineYPx - boundingBox.height,
            widthPx: boundingBox.width,
            heightPx: boundingBox.height,
          },
          context,
          label: "inline image",
        });
        sink.glyphs.push(...primitives);
        if (run.displayMode !== "block" && run.wrapType !== "topAndBottom") {
          layoutXPx += boundingBox.width;
        }
        break;
      }
      case "lineBreak":
      case "renderedPageBreak":
        break;
    }
  }

  return [...sink.backgrounds, ...sink.glyphs, ...sink.decorations];
};

type PaintListMarkerOptions = {
  readonly sink: LineSink;
  readonly block: ParagraphBlock;
  readonly context: BuildContext;
  readonly paintXPx: number;
  readonly baselineYPx: number;
  readonly isRtl: boolean;
};

const paintListMarker = ({
  sink,
  block,
  context,
  paintXPx,
  baselineYPx,
  isRtl,
}: PaintListMarkerOptions): void => {
  const marker = block.attrs?.listMarker;
  if (marker === undefined || marker.length === 0) {
    return;
  }
  const formatting = resolveListMarkerFont(block);
  const style: FontStyle = formatting;
  // A folded LISTNUM marker carries both slots separated by a tab; without the
  // second-slot offset there is nothing to align it to, so paint it as one run.
  const text = toPaintedText(marker.replaceAll("\t", " "));
  const glyphs = buildGlyphs({ text, style, allCaps: false, spaceDeltaPx: 0, collapsed: false });

  const revision = block.attrs?.listMarkerRevision;
  const color = revision
    ? trackedChangeColor(context.authorColors, revision.author, undefined)
    : DOC_CANVAS_TEXT;

  const font = context.fonts.intern({
    fontFamily: formatting.fontFamily,
    ...(formatting.alternateFontFamily === undefined
      ? {}
      : { alternateFontFamily: formatting.alternateFontFamily }),
    ...(formatting.bold === undefined ? {} : { bold: formatting.bold }),
    ...(formatting.italic === undefined ? {} : { italic: formatting.italic }),
    measureStyle: style,
  });

  const fontSizePx = ptToPx(formatting.fontSize);
  sink.glyphs.push({
    kind: "glyphRun",
    font,
    fontSizePx,
    color,
    xPx: paintXPx,
    baselineYPx,
    text: glyphs.text,
    advancesPx: glyphs.advancesPx,
    direction: (formatting.rtl ?? isRtl) ? "rtl" : "ltr",
  });

  if (revision) {
    const stroke: DisplayStroke = {
      color,
      thicknessPx: underlineThicknessPx(fontSizePx),
      pattern: "solid",
    };
    sink.decorations.push(
      revision.kind === "ins"
        ? horizontalLine(
            paintXPx,
            glyphs.widthPx,
            underlineCenterYPx(baselineYPx, fontSizePx),
            stroke,
          )
        : horizontalLine(
            paintXPx,
            glyphs.widthPx,
            strikethroughCenterYPx(baselineYPx, fontSizePx),
            stroke,
          ),
    );
  }
};

type PaintTabOptions = {
  readonly sink: LineSink;
  readonly context: BuildContext;
  readonly runs: readonly Run[];
  readonly index: number;
  readonly run: TabRun;
  readonly tabContext: TabContext;
  readonly layoutXPx: number;
  readonly paintXPx: number;
  readonly baselineYPx: number;
  readonly layoutRightEdgeXPx: number;
};

/**
 * A tab's advance, plus its leader glyphs when the stop sets one.
 *
 * The end-aligned clamp mirrors `measureParagraph.ts:1795-1802` rather than the
 * painter's flex "right anchor": both put the trailing content flush against
 * the line's right edge, but the measurer's version is the one the line width
 * was decided on, and in an explicit-x world it is just arithmetic.
 */
const paintTab = ({
  sink,
  context,
  runs,
  index,
  run,
  tabContext,
  layoutXPx,
  paintXPx,
  baselineYPx,
  layoutRightEdgeXPx,
}: PaintTabOptions): number => {
  let followingWidthPx = 0;
  let followingText = "";
  for (let next = index + 1; next < runs.length; next += 1) {
    // SAFETY: next < runs.length.
    const following = runs[next]!;
    if (following.kind === "tab" || following.kind === "lineBreak") {
      break;
    }
    if (following.kind === "text" || following.kind === "field" || following.kind === "math") {
      const text = paintableText(following, context);
      followingText += text;
      followingWidthPx += measureTextWidth(
        following.kind === "text" && following.allCaps ? text.toLocaleUpperCase() : text,
        runFontStyle(following),
      );
    } else if (following.kind === "image" && !isFloatingImageRun(following)) {
      followingWidthPx += inlineImageBoundingBox(following).width;
    }
  }

  const decimalIndex = followingText.indexOf(".");
  const decimalPrefixWidthPx =
    decimalIndex === -1 ? 0 : measureDecimalPrefixWidth(runs, index, decimalIndex, context);

  const result = calculateTabWidth(layoutXPx, tabContext, {
    followingWidth: followingWidthPx,
    ...(decimalPrefixWidthPx > 0 ? { decimalPrefixWidth: decimalPrefixWidthPx } : {}),
  });

  let widthPx = result.width;
  if (
    result.alignment === "end" &&
    layoutXPx + widthPx + followingWidthPx > layoutRightEdgeXPx + WIDTH_EPSILON_PX
  ) {
    widthPx = Math.max(1, layoutRightEdgeXPx - layoutXPx - followingWidthPx);
  }

  const leaderChar = result.leader === undefined ? null : getLeaderChar(result.leader);
  if (leaderChar !== null && widthPx > 0) {
    const style = runFontStyle(run);
    const unitWidthPx = measureTextWidth(leaderChar, style);
    if (unitWidthPx > 0) {
      const count = Math.floor(widthPx / unitWidthPx);
      if (count > 0) {
        const glyphs = buildGlyphs({
          text: leaderChar.repeat(count),
          style,
          allCaps: false,
          spaceDeltaPx: 0,
          collapsed: false,
        });
        const font = context.fonts.intern({
          fontFamily: style.fontFamily ?? DEFAULT_FONT_FAMILY,
          ...(style.bold === undefined ? {} : { bold: style.bold }),
          ...(style.italic === undefined ? {} : { italic: style.italic }),
          measureStyle: style,
        });
        const clip: DisplayRect = {
          xPx: paintXPx,
          yPx: baselineYPx - ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE),
          widthPx,
          heightPx: ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE) * 2,
        };
        sink.glyphs.push({
          kind: "clipGroup",
          rect: clip,
          children: [
            {
              kind: "glyphRun",
              font,
              fontSizePx: ptToPx(style.fontSize ?? DEFAULT_FONT_SIZE),
              color:
                run.color === undefined
                  ? DOC_CANVAS_TEXT
                  : (parseDisplayColor(run.color) ?? DOC_CANVAS_TEXT),
              xPx: paintXPx,
              baselineYPx,
              text: glyphs.text,
              advancesPx: glyphs.advancesPx,
              direction: "ltr",
            },
          ],
        });
      }
    }
  }

  return widthPx;
};

const measureDecimalPrefixWidth = (
  runs: readonly Run[],
  tabIndex: number,
  decimalIndex: number,
  context: BuildContext,
): number => {
  let width = 0;
  let consumed = 0;
  for (let index = tabIndex + 1; index < runs.length && consumed < decimalIndex; index += 1) {
    // SAFETY: index < runs.length.
    const run = runs[index]!;
    if (run.kind === "tab" || run.kind === "lineBreak") {
      break;
    }
    if (run.kind !== "text" && run.kind !== "field" && run.kind !== "math") {
      continue;
    }
    const text = paintableText(run, context);
    const take = Math.min(text.length, decimalIndex - consumed);
    if (take > 0) {
      width += measureTextWidth(text.slice(0, take), runFontStyle(run));
    }
    consumed += text.length;
  }
  return width;
};

/** Every primitive a paragraph fragment paints, back to front. */
export const paintParagraphFragment = (
  options: ParagraphPaintOptions,
): readonly DisplayPrimitive[] => {
  const { fragment, block, measure, context } = options;
  const primitives: DisplayPrimitive[] = paintParagraphChrome(options);

  const { alignment, indentLeft, indentRight, isRtl } = resolvePhysicalParagraphInlineLayout(block);
  const indent = block.attrs?.indent;
  // The measurer's formula, not the painter's `hanging ? -hanging : firstLine`:
  // line breaking was decided on this one, so paint must agree with it.
  const firstLineOffsetPx = (indent?.firstLine ?? 0) - (indent?.hanging ?? 0);
  const markerInlineWidthPx = getListMarkerInlineWidth(block);
  const paragraphEndsWithLineBreak = block.runs.at(-1)?.kind === "lineBreak";
  const totalLines = measure.lines.length;

  if (isRtl) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.bidiReordering,
      context.pageIndex,
      "runs of an RTL paragraph are placed in logical order; the builder runs no bidi algorithm",
    );
  }

  let cursorYPx = fragment.y;
  for (let offset = 0; offset < fragment.toLine - fragment.fromLine; offset += 1) {
    const lineIndex = fragment.fromLine + offset;
    const line = measure.lines[lineIndex];
    if (!line) {
      continue;
    }
    cursorYPx += line.floatSkipBefore ?? 0;

    const isFirstLine = lineIndex === 0 && fragment.continuesFromPrev !== true;
    const geometry = resolveLineGeometry({
      fragment,
      line,
      lineTopYPx: cursorYPx,
      alignment,
      isRtl,
      indentLeft,
      indentRight,
      firstLineOffsetPx,
      markerInlineWidthPx,
      isFirstLine,
    });

    primitives.push(
      ...paintLine({
        block,
        line,
        geometry,
        context,
        alignment,
        isRtl,
        isFirstLine,
        isLastLine: lineIndex === totalLines - 1,
        paragraphEndsWithLineBreak,
        availableWidthPx:
          fragment.width -
          indentLeft -
          indentRight -
          (line.leftOffset ?? 0) -
          (line.rightOffset ?? 0),
        indentLeft,
        firstLineOffsetPx,
        markerInlineWidthPx,
      }),
    );

    cursorYPx += line.lineHeight;
  }

  return primitives;
};
