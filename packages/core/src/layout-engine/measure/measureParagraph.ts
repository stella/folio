/**
 * Paragraph measurement module
 *
 * Measures paragraph blocks and computes line breaking.
 * Converts runs into measured lines with typography metrics.
 */

import { calculateTabWidth, pixelsToTwips } from "./tabCalculator";
import type { TabContext } from "./tabCalculator";
import { isCjkFont } from "../../utils/fontResolver";
import { inlineImageBoundingBox } from "../../utils/rotationBoundingBox";
import { measuredLineAdvance } from "../lineFlow";
import type { ParagraphBlock, ParagraphMeasure, MeasuredLine, TextRun } from "../types";
import { isFloatingImageRun } from "../types";
import { resolveEffectiveLineBreakPolicy } from "./effectiveLineBreakPolicy";
import {
  getFloatingAvailableWidth,
  getFloatingMargins,
  type FloatingImageZone,
} from "./floatingZones";
import { getListMarkerInlineWidth } from "./listMarkerWidth";
import { buildRunFontStyle, DEFAULT_FONT_FAMILY, ptToPx } from "./measureHelpers";
import { getFontMetrics, measureRun, measureTextWidth } from "./measureProvider";
import type { FontStyle } from "./measureTypes";
import { findHyphenationBreaks, findWordBreaks, isBreakBetween, isBreakChar } from "./lineBreaks";
import {
  DEFAULT_FONT_SIZE,
  DEFAULT_LINE_HEIGHT_MULTIPLIER,
  WIDTH_TOLERANCE,
  runToFontStyle,
  isTextRun,
  isTabRun,
  isImageRun,
  isLineBreakRun,
  isFieldRun,
  fieldMeasureText,
  isMathRun,
  isEmptyTextRun,
} from "./paragraphMeasureShared";
import type { LineTypography, LineState } from "./paragraphMeasureShared";
import {
  cjkLineHeightStyle,
  complexScriptLineHeightStyle,
  calculateTypographyMetrics,
  calculateEmptyParagraphMetrics,
  applyDocumentGrid,
} from "./lineTypography";
import { rotatedBlockImageHeight, estimateMathFootprintPx } from "./inlineObjectMetrics";
import {
  measureInlineWidthAfterTab,
  hasFollowingTabOnLine,
  hasPriorTabOnLine,
  canClampTabToRightEdge,
  measureDecimalPrefixWidthAfterTab,
} from "./paragraphTabs";
import {
  uppercaseLetterRatio,
  resolveJustifyFitStrategy,
  compressibleSpaceWidth,
  justifyFitTolerance,
} from "./justification";
import type { JustificationProfile } from "./justification";
import {
  findMaxFittingLength,
  exceedsHyphenationZone,
  isIgnorableFinalTailRun,
  isFinalTextCandidate,
  resolveTextCandidateFit,
  trimTrailingSpacesAndTabs,
  measureWordWithTrailingWhitespace,
  precedingLetterSpacing,
  trailingHangingPunctuationWidth,
} from "./lineFitting";
import {
  computeTrailingGlueWidths,
  computeFollowingTextLeads,
  computeProtectedCrossRunGlueWidths,
  collectCrossRunWord,
  measureCrossRunPrefix,
} from "./crossRunWords";
import type { CrossRunPrefix } from "./crossRunWords";
import { MIN_WRAP_SEGMENT_WIDTH, findClearLineY } from "./floatingLineClearance";

export { clampFloatingWrapMargins } from "./clampFloatingWrapMargins";
export type { FloatingImageZone } from "./floatingZones";
export { MIN_WRAP_SEGMENT_WIDTH, findClearLineY } from "./floatingLineClearance";

/**
 * Options for paragraph measurement
 */
export type MeasureParagraphOptions = {
  /** Floating image exclusion zones that affect line widths */
  floatingZones?: FloatingImageZone[];
  /** Y offset of this paragraph relative to the exclusion zones (default: 0) */
  paragraphYOffset?: number;
  /** Field run `pmStart` -> resolved display text, so a field measures at its
   *  painted width instead of the cached fallback. */
  fieldValues?: ReadonlyMap<number, string>;
  /** Header/footer tabs may be authored in the page margin, beyond body width. */
  allowEndTabOverflow?: boolean;
};

/**
 * Measure a paragraph block and compute line breaks
 *
 * @param block - The paragraph block to measure
 * @param maxWidth - Maximum available width for the paragraph
 * @param options - Optional measurement options (floating zones, Y offset)
 * @returns ParagraphMeasure with lines and total height
 */
export function measureParagraph(
  block: ParagraphBlock,
  maxWidth: number,
  options?: MeasureParagraphOptions,
): ParagraphMeasure {
  const runs = block.runs;
  const attrs = block.attrs;
  const spacing = attrs?.spacing;
  const isJustifiedParagraph = attrs?.alignment === "justify";
  const justificationProfile: JustificationProfile = {
    hasTabRuns: isJustifiedParagraph && runs.some(isTabRun),
    uppercaseRatio: isJustifiedParagraph
      ? uppercaseLetterRatio(runs.map((run) => (isTextRun(run) ? (run.text ?? "") : "")).join(""))
      : 0,
  };
  const firstLineJustifyFitStrategy = resolveJustifyFitStrategy(block, true, justificationProfile);
  const continuationJustifyFitStrategy = resolveJustifyFitStrategy(
    block,
    false,
    justificationProfile,
  );

  // Floating image support
  const floatingZones = options?.floatingZones;
  const paragraphYOffset = options?.paragraphYOffset ?? 0;

  // Handle indentation
  const indent = attrs?.indent;
  const indentLeft = indent?.left ?? 0;
  const indentRight = indent?.right ?? 0;
  const firstLineOffset = (indent?.firstLine ?? 0) - (indent?.hanging ?? 0);

  // Calculate base available widths (before floating image adjustment)
  const bodyContentWidth = Math.max(1, maxWidth - indentLeft - indentRight);
  // First line offset: positive = first-line indent (less space), negative = hanging (more space)
  // Subtracting gives correct width in both cases
  let baseFirstLineWidth = Math.max(1, bodyContentWidth - firstLineOffset);

  // List marker on the first line: the marker renders as an inline-block
  // span that's *not* in the run list, so the run-based line breaker
  // doesn't see it. The first line's content area spans from the marker's
  // start to the right margin — for a hanging list that's
  // `bodyContentWidth + hanging` (already widened via `firstLineOffset`);
  // for a first-line indent it's `bodyContentWidth − firstLine`. Subtract
  // the marker's actual painted footprint (`getListMarkerInlineWidth`) so
  // the line breaker sees the same text room the painter leaves.
  //
  // The subtraction is unconditional:
  //
  // - Hanging + `w:suff="tab"` (fitting): markerInlineWidth = hanging, so
  //   the subtraction exactly cancels the `+ hanging` widening and the
  //   text budget reduces to bodyContentWidth (matches body wrap).
  // - Hanging + tab overflow: markerInlineWidth > hanging, subtracting
  //   yields bodyContentWidth − overflow (text budget shrinks past the
  //   body wrap column, matching Word's advance to next tab stop).
  // - Hanging + `w:suff="space"|"nothing"`: markerInlineWidth < hanging,
  //   so the budget is bodyContentWidth + (hanging − markerInlineWidth) —
  //   first line is wider than subsequent lines, matching the painter
  //   which starts body before indentLeft.
  // - First-line indent (no hanging): subtract the full marker width.
  const markerInlineWidth = getListMarkerInlineWidth(block);
  if (markerInlineWidth > 0) {
    baseFirstLineWidth = Math.max(1, baseFirstLineWidth - markerInlineWidth);
  }

  // Track cumulative height for floating zone calculations
  let cumulativeHeight = 0;
  // Vertical space queued for the next line to finalize — set when we hop
  // past a float that leaves no usable horizontal width at the current Y.
  // Cleared each time finalizeLine attaches it to a MeasuredLine.
  let pendingFloatSkip = 0;

  /**
   * If floats leave no usable horizontal room at `cumulativeHeight`, advance
   * past them by mutating cumulativeHeight + pendingFloatSkip.
   */
  const skipObstructingFloats = (lineHeight: number, lineMaxWidth: number): void => {
    if (!floatingZones || floatingZones.length === 0) {
      return;
    }
    const absoluteY = paragraphYOffset + cumulativeHeight;
    const clearY = findClearLineY(
      absoluteY,
      lineHeight,
      floatingZones,
      lineMaxWidth,
      MIN_WRAP_SEGMENT_WIDTH,
    );
    const skip = clearY - absoluteY;
    if (skip > 0) {
      cumulativeHeight += skip;
      pendingFloatSkip += skip;
    }
  };

  // Calculate first line width with floating zone adjustment
  // Estimate first line height for floating margin calculation
  const estimatedFirstLineHeight = ptToPx(DEFAULT_FONT_SIZE) * DEFAULT_LINE_HEIGHT_MULTIPLIER;
  skipObstructingFloats(estimatedFirstLineHeight, baseFirstLineWidth);
  const firstLineFloatingMargins = getFloatingMargins(
    cumulativeHeight,
    estimatedFirstLineHeight,
    floatingZones,
    paragraphYOffset,
  );
  const firstLineWidth = Math.max(
    1,
    getFloatingAvailableWidth(firstLineFloatingMargins, baseFirstLineWidth),
  );

  const lines: MeasuredLine[] = [];
  let consecutiveHyphenatedLines = 0;

  if (attrs?.suppressEmptyParagraphHeight) {
    const finalRunIndex = Math.max(0, runs.length - 1);
    const finalRun = runs.at(-1);
    lines.push({
      fromRun: 0,
      fromChar: 0,
      toRun: finalRunIndex,
      toChar: finalRun?.kind === "text" ? finalRun.text.length : 0,
      width: 0,
      ascent: 0,
      descent: 0,
      lineHeight: 0,
    });

    return {
      kind: "paragraph",
      lines,
      totalHeight: 0,
    };
  }

  // Handle empty paragraph
  if (runs.length === 0) {
    const emptyFontSize = attrs?.defaultFontSize ?? DEFAULT_FONT_SIZE;
    const emptyFontFamily = attrs?.defaultFontFamily ?? DEFAULT_FONT_FAMILY;
    const emptyMetrics = calculateEmptyParagraphMetrics(
      emptyFontSize,
      spacing,
      emptyFontFamily,
      attrs,
    );
    // Reference layout reserves a second line box for a story-leading empty top-level
    // outline paragraph. Keep it as one logical/caret line while expanding
    // the line's advance; later and ordinary empty paragraphs stay at the
    // normal single-line height.
    const outlineLineHeight = attrs?.reserveEmptyOutlineHeight
      ? emptyMetrics.lineHeight * 2
      : emptyMetrics.lineHeight;
    lines.push({
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 0,
      width: 0,
      ...emptyMetrics,
      lineHeight: outlineLineHeight,
    });

    let totalHeight = outlineLineHeight;
    if (spacing?.before) {
      totalHeight += spacing.before;
    }
    if (spacing?.after) {
      totalHeight += spacing.after;
    }

    return {
      kind: "paragraph",
      lines,
      totalHeight,
    };
  }

  // Whitespace-only text is layout-empty: use the paragraph-mark formatting
  // carried by attrs rather than formatting from an otherwise invisible run.
  if (runs.length > 0 && runs.every((run) => isTextRun(run) && isEmptyTextRun(run))) {
    const fontSize = attrs?.defaultFontSize ?? DEFAULT_FONT_SIZE;
    const fontFamily = attrs?.defaultFontFamily ?? DEFAULT_FONT_FAMILY;
    const emptyMetrics = calculateEmptyParagraphMetrics(fontSize, spacing, fontFamily, attrs);

    lines.push({
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 0,
      width: 0,
      ...emptyMetrics,
    });

    let totalHeight = emptyMetrics.lineHeight;
    if (spacing?.before) {
      totalHeight += spacing.before;
    }
    if (spacing?.after) {
      totalHeight += spacing.after;
    }

    return {
      kind: "paragraph",
      lines,
      totalHeight,
    };
  }

  const trailingGlueWidths = computeTrailingGlueWidths(block);
  const protectedCrossRunGlueWidths = computeProtectedCrossRunGlueWidths(block);
  const followingTextLeads = computeFollowingTextLeads(block);

  // Initialize line state
  let currentLine: LineState = {
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: 0,
    width: 0,
    trailingWhitespaceWidth: 0,
    regularSpaceWidth: 0,
    maxFontSize: DEFAULT_FONT_SIZE,
    maxFontMetrics: null,
    maxTabFontSize: DEFAULT_FONT_SIZE,
    maxTabFontMetrics: null,
    maxImageHeightPx: 0,
    maxExactImageHeightPx: 0,
    maxMathHeightPx: 0,
    availableWidth: firstLineWidth,
    leftOffset: firstLineFloatingMargins.leftMargin,
    rightOffset: firstLineFloatingMargins.rightMargin,
    ...(firstLineFloatingMargins.segments?.length
      ? { segmentZones: firstLineFloatingMargins.segments }
      : {}),
  };

  const calculateLineTypography = (line: LineState): LineTypography => {
    const paragraphFontSize = attrs?.defaultFontSize ?? DEFAULT_FONT_SIZE;
    const paragraphFontFamily = attrs?.defaultFontFamily ?? DEFAULT_FONT_FAMILY;
    let lineFontMetrics = line.maxFontMetrics;
    let fontSize = line.maxFontSize;
    if (!lineFontMetrics) {
      lineFontMetrics = line.maxTabFontMetrics;
      fontSize = lineFontMetrics ? line.maxTabFontSize : paragraphFontSize;
    }
    const metrics =
      lineFontMetrics ??
      getFontMetrics({
        fontSize: paragraphFontSize,
        fontFamily: paragraphFontFamily,
        ...(attrs?.defaultAlternateFontFamily !== undefined
          ? { alternateFontFamily: attrs.defaultAlternateFontFamily }
          : {}),
      });
    const typography = calculateTypographyMetrics(fontSize, spacing, metrics);

    // If an inline image or stacked equation is taller than the text-based
    // line height, the line grows to fit it. The reference layout seats these inline objects
    // as tall glyphs on the text baseline.
    const finalTypography = { ...typography };
    const inlineObjectHeight = Math.max(line.maxImageHeightPx, line.maxMathHeightPx);
    if (inlineObjectHeight > finalTypography.lineHeight) {
      const objectHeight = inlineObjectHeight;
      const buffer = finalTypography.descent;
      // An image-only line uses the authored image footprint directly. There
      // is no text baseline that needs font descent above or below the image;
      // adding one silently grows large diagrams and can change pagination.
      // Keep this paired with the painter's image-only flex alignment.
      const soleRun = line.fromRun === line.toRun ? runs[line.fromRun] : undefined;
      if (line.maxExactImageHeightPx >= objectHeight) {
        finalTypography.lineHeight = objectHeight;
        finalTypography.ascent = objectHeight;
        finalTypography.descent = 0;
        return applyDocumentGrid(finalTypography, attrs, spacing);
      }
      if (soleRun && isImageRun(soleRun)) {
        finalTypography.lineHeight = objectHeight;
        finalTypography.ascent = objectHeight;
        finalTypography.descent = 0;
      } else {
        // Object flowing with text/tabs (e.g. a logo + label header line):
        // the full object height sits above the baseline and only the text
        // descent is reserved below — no extra leading above it. The painter
        // baseline-aligns the row so the object bottom lands on the text
        // baseline.
        finalTypography.lineHeight = objectHeight + buffer;
        finalTypography.ascent = objectHeight;
      }
    }

    return applyDocumentGrid(finalTypography, attrs, spacing);
  };

  const getPostWrapAvailableWidth = (): number => {
    if (!floatingZones || floatingZones.length === 0) {
      return bodyContentWidth;
    }

    const lineTypography = calculateLineTypography(currentLine);
    let nextCumulativeHeight = cumulativeHeight + lineTypography.lineHeight;
    const estimatedLineHeight = ptToPx(DEFAULT_FONT_SIZE) * DEFAULT_LINE_HEIGHT_MULTIPLIER;
    const absoluteY = paragraphYOffset + nextCumulativeHeight;
    const clearY = findClearLineY(
      absoluteY,
      estimatedLineHeight,
      floatingZones,
      bodyContentWidth,
      MIN_WRAP_SEGMENT_WIDTH,
    );
    const skip = clearY - absoluteY;
    if (skip > 0) {
      nextCumulativeHeight += skip;
    }

    const floatingMargins = getFloatingMargins(
      nextCumulativeHeight,
      estimatedLineHeight,
      floatingZones,
      paragraphYOffset,
    );

    return Math.max(1, getFloatingAvailableWidth(floatingMargins, bodyContentWidth));
  };

  /**
   * Finalize and push the current line to the lines array
   */
  const finalizeLine = (): void => {
    const finalTypography = calculateLineTypography(currentLine);

    const line: MeasuredLine = {
      fromRun: currentLine.fromRun,
      fromChar: currentLine.fromChar,
      toRun: currentLine.toRun,
      toChar: currentLine.toChar,
      width: Math.max(0, currentLine.width - currentLine.trailingWhitespaceWidth),
      ...finalTypography,
      ...(currentLine.justificationPaint !== undefined
        ? { justificationPaint: currentLine.justificationPaint }
        : {}),
      ...(currentLine.discretionaryHyphen
        ? { discretionaryHyphen: currentLine.discretionaryHyphen }
        : {}),
      ...(currentLine.renderedPageBreakBefore ? { renderedPageBreakBefore: true } : {}),
    };

    // Only add offsets if they're non-zero (for floating images)
    if (currentLine.leftOffset > 0) {
      line.leftOffset = currentLine.leftOffset;
    }
    if (currentLine.rightOffset > 0) {
      line.rightOffset = currentLine.rightOffset;
    }

    // Attach any queued float-skip to this line; the painter reserves it
    // via marginTop and totalHeight already grew by this amount above.
    if (pendingFloatSkip > 0) {
      line.floatSkipBefore = pendingFloatSkip;
      pendingFloatSkip = 0;
    }

    lines.push(line);
    consecutiveHyphenatedLines = currentLine.discretionaryHyphen
      ? consecutiveHyphenatedLines + 1
      : 0;

    // Update cumulative height for next line's floating zone calculation
    cumulativeHeight += finalTypography.lineHeight;
  };

  /**
   * Start a new line after the current one
   */
  const startNewLine = (runIndex: number, charIndex: number): void => {
    finalizeLine();

    // Calculate available width for new line based on floating zones
    // Estimate the new line's height for overlap calculation
    const estimatedLineHeight = ptToPx(DEFAULT_FONT_SIZE) * DEFAULT_LINE_HEIGHT_MULTIPLIER;
    skipObstructingFloats(estimatedLineHeight, bodyContentWidth);
    const floatingMargins = getFloatingMargins(
      cumulativeHeight,
      estimatedLineHeight,
      floatingZones,
      paragraphYOffset,
    );

    // Body content width minus floating image margins
    const adjustedWidth = Math.max(1, getFloatingAvailableWidth(floatingMargins, bodyContentWidth));

    currentLine = {
      fromRun: runIndex,
      fromChar: charIndex,
      toRun: runIndex,
      toChar: charIndex,
      width: 0,
      trailingWhitespaceWidth: 0,
      regularSpaceWidth: 0,
      maxFontSize: DEFAULT_FONT_SIZE,
      maxFontMetrics: null,
      maxTabFontSize: DEFAULT_FONT_SIZE,
      maxTabFontMetrics: null,
      maxImageHeightPx: 0,
      maxExactImageHeightPx: 0,
      maxMathHeightPx: 0,
      availableWidth: adjustedWidth,
      leftOffset: floatingMargins.leftMargin,
      rightOffset: floatingMargins.rightMargin,
      ...(floatingMargins.segments?.length ? { segmentZones: floatingMargins.segments } : {}),
    };
  };

  /**
   * Update max font tracking for the current line
   */
  const updateMaxFont = (style: FontStyle): void => {
    const fontSize = style.fontSize ?? DEFAULT_FONT_SIZE;
    // Update when this is the first run on the line (maxFontMetrics not yet set)
    // or when we find a larger font size. Without the !maxFontMetrics check,
    // lines with only <11pt text would use the 11pt default, inflating line height.
    if (!currentLine.maxFontMetrics || fontSize > currentLine.maxFontSize) {
      currentLine.maxFontSize = fontSize;
      currentLine.maxFontMetrics = getFontMetrics(style);
      return;
    }
    // Same-size tie: an East-Asian face carries a taller single-line ratio
    // than a Latin face at the same font size, and Word sizes a mixed line to
    // its tallest run — so a CJK line-height style must be able to raise the
    // metrics even when it does not raise the font size (Latin run first, CJK
    // run after, both at the body size is the common Japanese-document case).
    // Gated on the candidate being a CJK face so Latin-only lines keep the
    // existing first-run-wins tie behaviour bit-for-bit.
    if (fontSize !== currentLine.maxFontSize || !isCjkFont(style.fontFamily ?? "")) {
      return;
    }
    const metrics = getFontMetrics(style);
    if (metrics.singleLineRatio > currentLine.maxFontMetrics.singleLineRatio) {
      currentLine.maxFontMetrics = metrics;
    }
  };

  /**
   * Track the tallest tab on the line separately from content runs: a tab's
   * font does not raise a line that carries other content, so a tab run
   * formatted larger than its neighbouring text (a TOC entry whose tab keeps
   * the paragraph-mark size) leaves the line at the text's height.
   */
  const updateMaxTabFont = (style: FontStyle): void => {
    const fontSize = style.fontSize ?? DEFAULT_FONT_SIZE;
    if (!currentLine.maxTabFontMetrics || fontSize > currentLine.maxTabFontSize) {
      currentLine.maxTabFontSize = fontSize;
      currentLine.maxTabFontMetrics = getFontMetrics(style);
    }
  };

  let crossRunResume: { runIndex: number; charIndex: number } | undefined;

  // Process each run
  for (let runIndex = 0; runIndex < runs.length; runIndex++) {
    if (crossRunResume && runIndex < crossRunResume.runIndex) {
      continue;
    }
    // SAFETY: runIndex is bounded by runs.length
    const run = runs[runIndex]!;

    if (run.kind === "renderedPageBreak") {
      const hasFollowingContent = runs
        .slice(runIndex + 1)
        .some((followingRun) => !isIgnorableFinalTailRun(followingRun));
      if (!hasFollowingContent) {
        currentLine.toRun = runIndex;
        currentLine.toChar = 0;
        continue;
      }
      const lineHasContent =
        currentLine.width > 0 ||
        currentLine.maxImageHeightPx > 0 ||
        currentLine.maxMathHeightPx > 0;
      if (lineHasContent) {
        startNewLine(runIndex + 1, 0);
      } else {
        currentLine.fromRun = runIndex + 1;
        currentLine.toRun = runIndex + 1;
        currentLine.fromChar = 0;
        currentLine.toChar = 0;
      }
      currentLine.renderedPageBreakBefore = true;
      continue;
    }

    if (isLineBreakRun(run)) {
      // Force line break
      currentLine.toRun = runIndex;
      currentLine.toChar = 0;
      startNewLine(runIndex + 1, 0);
      continue;
    }

    if (isTabRun(run)) {
      const style = runToFontStyle(run);
      updateMaxTabFont(style);

      const followingWidth = measureInlineWidthAfterTab(runs, runIndex, options?.fieldValues);
      const decimalPrefixWidth = measureDecimalPrefixWidthAfterTab(
        runs,
        runIndex,
        options?.fieldValues,
      );

      // Tab width comes from the shared tab-stop model (`calculateTabWidth` —
      // computeTabStops + alignment) that the painter also uses, so measurer
      // and painter agree on line widths. `calculateTabWidth` works in
      // content-area coordinates (tab stops are measured from the
      // content-area left edge), so the indent and any first-line offset are
      // folded in here; the line-wrap math further down stays indent-relative.
      const lineX = currentLine.width + currentLine.leftOffset;
      const isFirstLine = lines.length === 0;
      // First-line text body starts past any list marker (the marker occupies
      // the hanging zone for `suff="tab"` lists, or sits inline before the
      // body for other suffixes). Folding `markerInlineWidth` in keeps
      // `contentX` aligned with where the body cursor actually sits in the
      // content area — without it, a hanging + tab-suffix list (e.g. a TOC
      // entry style with a right-aligned dot-leader tab stop) understates
      // the cursor by `hanging`, the resulting tab width overshoots the
      // line's text budget by the same amount, and the line wrongly wraps.
      const contentX = indentLeft + (isFirstLine ? firstLineOffset + markerInlineWidth : 0) + lineX;
      const tabContext: TabContext = {
        ...(attrs?.tabs !== undefined ? { explicitStops: attrs.tabs } : {}),
        ...(attrs?.defaultTabStopTwips !== undefined
          ? { defaultTabInterval: attrs.defaultTabStopTwips }
          : {}),
        leftIndent: pixelsToTwips(indentLeft),
      };
      const tabResult = calculateTabWidth(contentX, tabContext, {
        followingWidth,
        decimalPrefixWidth,
      });
      let tabWidth = tabResult.width;
      const authoredEndpoint = contentX + tabWidth + followingWidth;
      const activeContentRightEdge = maxWidth - currentLine.rightOffset;
      // Explicit `w:tab` stops are authored against the content box, so one
      // past the right indent (`w:ind/@w:right`) is honoured while the content
      // it positions still ends inside the active content frame.
      const preservesAuthoredStop =
        (tabResult.alignment === "end" && options?.allowEndTabOverflow === true) ||
        (tabResult.explicit === true &&
          authoredEndpoint <= activeContentRightEdge + WIDTH_TOLERANCE);
      const landsOnLeftIndent =
        tabResult.alignment === "start" &&
        indentLeft > 0 &&
        Math.abs(contentX + tabWidth - indentLeft) <= WIDTH_TOLERANCE;

      const lineRightEdgeX =
        indentLeft +
        (isFirstLine ? firstLineOffset + markerInlineWidth : 0) +
        currentLine.availableWidth +
        currentLine.leftOffset;
      if (
        !preservesAuthoredStop &&
        !landsOnLeftIndent &&
        !hasFollowingTabOnLine(runs, runIndex) &&
        canClampTabToRightEdge(
          tabResult.alignment,
          currentLine.width,
          hasPriorTabOnLine(runs, runIndex),
          followingWidth,
          currentLine.availableWidth,
        ) &&
        (tabWidth > 0 || followingWidth > 0) &&
        contentX + tabWidth + followingWidth > lineRightEdgeX + WIDTH_TOLERANCE
      ) {
        tabWidth = Math.max(1, lineRightEdgeX - contentX - followingWidth);
      }

      // A preserved explicit stop keeps enough budget for its endpoint even
      // when a paragraph right indent narrows the ordinary line edge; default
      // and out-of-frame tabs retain the clamp above.
      if (preservesAuthoredStop) {
        currentLine.availableWidth = Math.max(
          currentLine.availableWidth,
          currentLine.width + tabWidth + followingWidth,
        );
      }

      if (currentLine.width + tabWidth > currentLine.availableWidth + WIDTH_TOLERANCE) {
        // Tab doesn't fit, start new line
        startNewLine(runIndex, 0);
        updateMaxTabFont(style);
      }

      currentLine.width += tabWidth;
      currentLine.trailingWhitespaceWidth = 0;
      currentLine.toRun = runIndex;
      currentLine.toChar = 1;
      continue;
    }

    if (isImageRun(run)) {
      const wrapType = run.wrapType;
      // Keep measurement aligned with the shared anchored-image predicate.
      // These images paint in a page layer and must not reserve inline width
      // or height at their host run.
      if (isFloatingImageRun(run)) {
        currentLine.toRun = runIndex;
        currentLine.toChar = 1;
        continue;
      }

      // Handle topAndBottom (block) images - they get their own line
      if (wrapType === "topAndBottom" || run.displayMode === "block") {
        // If current line has content, finish it first
        if (currentLine.width > 0) {
          startNewLine(runIndex, 0);
        }

        // The image gets its own line. For rotated images, reserve the
        // axis-aligned bounding-box height so the painter's bbox wrapper
        // (`renderBlockImage`, eigenpal #424) doesn't overflow the line
        // and bleed into the next paragraph. Non-rotated images keep
        // their intrinsic height. Helpers duplicated from the painter
        // until cross-PR dedupe with #518 lands.
        const imageHeight = rotatedBlockImageHeight(run);
        const distTop = run.distTop ?? 6;
        const distBottom = run.distBottom ?? 6;

        // Update line to contain just this image
        currentLine.toRun = runIndex;
        currentLine.toChar = 1;
        // Use image height plus margins as line height (already in pixels)
        currentLine.maxImageHeightPx = imageHeight + distTop + distBottom;

        // Start a new line after the image for subsequent content
        startNewLine(runIndex + 1, 0);
        continue;
      }

      // Handle inline image. Rotated images occupy their axis-aligned bbox,
      // not the raw `run.width × run.height`; the painter wraps them in a
      // bbox-sized span (eigenpal #424). The measurer must reserve the same
      // dims so line-break and line-height match what gets painted.
      const inlineBbox = inlineImageBoundingBox(run);
      const imageWidth = inlineBbox.width;
      const imageHeight = inlineBbox.height;

      if (
        currentLine.width > 0 &&
        currentLine.width + imageWidth > currentLine.availableWidth + WIDTH_TOLERANCE
      ) {
        // Image doesn't fit, start new line. Guarded on a non-empty line:
        // wrapping an image that is already alone on the line can't make it fit
        // and would just insert a blank row above it.
        startNewLine(runIndex, 0);
      }

      // The measurer reserves the image's intrinsic box. The painter fits an
      // over-wide plain inline image down with CSS `max-width: 100%`
      // (eigenpal/docx-editor#760), but that is left as a purely visual cap: a
      // CSS percentage of the line element doesn't correspond to a single
      // computed width once first-line indents, list markers, or floating-image
      // line offsets are in play, so predicting it here would under-reserve
      // height and risk overlap. Reserving the intrinsic box keeps measurement,
      // selection, and click geometry mutually consistent and never short.

      // The image's vertical footprint in the line includes its wp:inline
      // distT/distB wrap distances. These default to 0 for inline images
      // (unlike the block path's synthetic 6px). The painter applies them as
      // top/bottom margins on the <img>, so the run's flex baseline (the
      // margin-box edge) stays consistent with this reserved height. Record it
      // only after the wrap check above: the footprint belongs to the line the
      // image actually lands on, not the line it wrapped away from — otherwise
      // the previous line inflates to image height while the image's own line
      // stays text-height and following content paints over the overflow.
      // (eigenpal/docx-editor#767, fixes #766.)
      const imageFootprintPx = imageHeight + (run.distTop ?? 0) + (run.distBottom ?? 0);
      if (imageFootprintPx > currentLine.maxImageHeightPx) {
        currentLine.maxImageHeightPx = imageFootprintPx;
      }
      if (run.exactLineHeight === true && imageFootprintPx > currentLine.maxExactImageHeightPx) {
        currentLine.maxExactImageHeightPx = imageFootprintPx;
      }

      currentLine.width += imageWidth;
      currentLine.trailingWhitespaceWidth = 0;
      currentLine.toRun = runIndex;
      currentLine.toChar = 1;
      continue;
    }

    if (isFieldRun(run)) {
      // Measure the field at its resolved value when known so the line breaker
      // agrees with the painter; otherwise the cached fallback text.
      const fallback = fieldMeasureText(run, options?.fieldValues);
      // Use the shared builder so the field result measures with the same fields
      // the painter renders — notably eastAsiaFontFamily, so a CJK field result
      // (DATE/TIME/PAGEREF/REF) wraps and tab-positions with its East-Asian font.
      const style = buildRunFontStyle(run, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE);
      updateMaxFont(style);

      const fieldWidth = measureTextWidth(fallback, style);
      const fieldSpaceWidth = compressibleSpaceWidth(fallback, style);
      const fieldTolerance = isJustifiedParagraph
        ? justifyFitTolerance(
            currentLine,
            lines.length === 0 ? firstLineJustifyFitStrategy : continuationJustifyFitStrategy,
            fieldSpaceWidth,
            { width: currentLine.width + fieldWidth, innerSpaceWidth: fieldSpaceWidth },
          )
        : WIDTH_TOLERANCE;
      if (
        currentLine.width > 0 &&
        currentLine.width + fieldWidth > currentLine.availableWidth + fieldTolerance
      ) {
        startNewLine(runIndex, 0);
        updateMaxFont(style);
      }

      currentLine.width += fieldWidth;
      currentLine.regularSpaceWidth += fieldSpaceWidth;
      currentLine.trailingWhitespaceWidth = 0;
      currentLine.toRun = runIndex;
      currentLine.toChar = 1;
      continue;
    }

    if (isMathRun(run)) {
      // Math is opaque to the line breaker: it can't wrap mid-equation.
      // Reserve the rendered MathML's approximate inline width using the
      // plain-text fallback measured in Cambria Math at the run's font
      // size — this matches what the browser actually puts on screen
      // closely enough for line-wrap decisions. The painter then injects
      // the real `<math>` element, which the browser sizes natively.
      const style: FontStyle = {
        fontFamily: run.fontFamily ?? "Cambria Math",
        fontSize: run.fontSize ?? DEFAULT_FONT_SIZE,
        ...(run.bold !== undefined ? { bold: run.bold } : {}),
        ...(run.italic !== undefined ? { italic: run.italic } : {}),
      };
      updateMaxFont(style);
      const mathText = run.plainText || "[equation]";
      const mathWidth = measureTextWidth(mathText, style);
      const mathFootprintPx = estimateMathFootprintPx(run);
      if (
        currentLine.width > 0 &&
        currentLine.width + mathWidth > currentLine.availableWidth + WIDTH_TOLERANCE
      ) {
        startNewLine(runIndex, 0);
        updateMaxFont(style);
      }
      if (mathFootprintPx > currentLine.maxMathHeightPx) {
        currentLine.maxMathHeightPx = mathFootprintPx;
      }
      currentLine.width += mathWidth;
      currentLine.trailingWhitespaceWidth = 0;
      currentLine.toRun = runIndex;
      currentLine.toChar = 1;
      continue;
    }

    if (isTextRun(run)) {
      const textRun = run as TextRun;
      const text = textRun.text;
      const style = runToFontStyle(textRun);
      const effectiveLineBreakPolicy = resolveEffectiveLineBreakPolicy({
        attrs: block.attrs,
        run: textRun,
      });
      const breakPolicy = effectiveLineBreakPolicy.provider;
      // Line height comes from the CJK-aware style; width measurement below
      // keeps `style` so wrapping is unchanged.
      const lineHeightStyle = cjkLineHeightStyle(
        textRun,
        complexScriptLineHeightStyle(textRun, style),
      );

      updateMaxFont(lineHeightStyle);

      if (!text || text.length === 0) {
        // Empty text run, just update position
        currentLine.toRun = runIndex;
        currentLine.toChar = 0;
        continue;
      }

      // Find word break points for wrapping
      const wordBreaks = findWordBreaks(text, breakPolicy);

      // Process text word by word
      let charIndex = crossRunResume?.runIndex === runIndex ? crossRunResume.charIndex : 0;
      if (crossRunResume?.runIndex === runIndex) {
        crossRunResume = undefined;
      }
      let wordBreakIndex = 0;
      let activeHyphenationWord:
        | { start: number; end: number; text: string; breaks: number[] }
        | undefined;

      while (charIndex < text.length) {
        // Find next word boundary
        while ((wordBreaks[wordBreakIndex] ?? Number.POSITIVE_INFINITY) <= charIndex) {
          wordBreakIndex += 1;
        }
        const nextBreak = wordBreaks[wordBreakIndex] ?? text.length;

        // Extract word (includes trailing space if present). Word consumes a
        // break-space without including it in the current line's fit width.
        // Keep the full width while accumulating in case another word follows
        // on this line, then remove any whitespace still trailing at finalize.
        const word = text.slice(charIndex, nextBreak);
        const { measuredWord, wordWidth, fullWordWidth } = measureWordWithTrailingWhitespace(
          word,
          style,
        );
        // Word-break tokens are measured separately, but character
        // spacing (`w:spacing`) follows every character, including the last
        // one of a token or run, so the spacing owed by the previous
        // character on the line is added before this token.
        let leadingLetterSpacing = precedingLetterSpacing({
          block,
          line: currentLine,
          runIndex,
          charIndex,
          runStyle: style,
        });
        const hangingPunctuationWidth = trailingHangingPunctuationWidth(
          measuredWord,
          style,
          breakPolicy,
          effectiveLineBreakPolicy.hangingPunctuation,
        );
        const isFirstLine = lines.length === 0;
        const regularSpaceWidth = compressibleSpaceWidth(measuredWord, style);
        const justifyFitStrategy = isFirstLine
          ? firstLineJustifyFitStrategy
          : continuationJustifyFitStrategy;
        const widthTolerance = isJustifiedParagraph
          ? justifyFitTolerance(currentLine, justifyFitStrategy, regularSpaceWidth, {
              width: currentLine.width + leadingLetterSpacing + wordWidth,
              innerSpaceWidth: 0,
            })
          : WIDTH_TOLERANCE;

        const automaticHyphenation = effectiveLineBreakPolicy.automaticHyphenation;
        const mayHyphenateLine =
          automaticHyphenation.type === "enabled" &&
          (automaticHyphenation.consecutiveLineLimit === 0 ||
            consecutiveHyphenatedLines < automaticHyphenation.consecutiveLineLimit) &&
          exceedsHyphenationZone(currentLine, automaticHyphenation.hyphenationZoneTwips);
        const isRunTail = nextBreak === text.length;
        const crossRunWord =
          mayHyphenateLine && isRunTail && word.length > 0 && !isBreakChar(word.at(-1))
            ? collectCrossRunWord({ block, startRunIndex: runIndex, startChar: charIndex })
            : undefined;
        if (
          mayHyphenateLine &&
          crossRunWord === undefined &&
          measuredWord.length > 0 &&
          (activeHyphenationWord === undefined ||
            charIndex < activeHyphenationWord.start ||
            charIndex >= activeHyphenationWord.end)
        ) {
          const fullWordEnd = charIndex + measuredWord.length;
          const fullWordText = text.slice(charIndex, fullWordEnd);
          activeHyphenationWord = {
            start: charIndex,
            end: fullWordEnd,
            text: fullWordText,
            breaks: findHyphenationBreaks(fullWordText, breakPolicy),
          };
        }

        const overflowsCurrentLine =
          wordWidth > 0 &&
          currentLine.width + leadingLetterSpacing + wordWidth >
            currentLine.availableWidth + widthTolerance;
        if (mayHyphenateLine && overflowsCurrentLine && activeHyphenationWord) {
          const consumed = charIndex - activeHyphenationWord.start;
          const spaceLeft =
            currentLine.availableWidth - currentLine.width - leadingLetterSpacing + widthTolerance;
          const hyphenWidth = measureTextWidth("-", style);
          let fittingBreak: number | undefined;
          for (const breakOffset of activeHyphenationWord.breaks) {
            if (breakOffset <= consumed || breakOffset >= activeHyphenationWord.text.length) {
              continue;
            }
            const prefix = activeHyphenationWord.text.slice(consumed, breakOffset);
            if (measureTextWidth(prefix, style) + hyphenWidth <= spaceLeft) {
              fittingBreak = breakOffset;
            }
          }
          if (fittingBreak !== undefined) {
            const absoluteBreak = activeHyphenationWord.start + fittingBreak;
            const prefix = text.slice(charIndex, absoluteBreak);
            currentLine.width +=
              leadingLetterSpacing + measureTextWidth(prefix, style) + hyphenWidth;
            currentLine.trailingWhitespaceWidth = 0;
            currentLine.toRun = runIndex;
            currentLine.toChar = absoluteBreak;
            currentLine.discretionaryHyphen = { runIndex };
            startNewLine(runIndex, absoluteBreak);
            updateMaxFont(lineHeightStyle);
            charIndex = absoluteBreak;
            continue;
          }
        }

        if (
          crossRunWord &&
          currentLine.width + leadingLetterSpacing + crossRunWord.width >
            currentLine.availableWidth + widthTolerance
        ) {
          const spaceLeft =
            currentLine.availableWidth - currentLine.width - leadingLetterSpacing + widthTolerance;
          let fittingPrefix: CrossRunPrefix | undefined;
          for (let breakIndex = crossRunWord.breaks.length - 1; breakIndex >= 0; breakIndex -= 1) {
            const breakOffset = crossRunWord.breaks[breakIndex];
            if (breakOffset === undefined) {
              continue;
            }
            if (breakOffset <= 0 || breakOffset >= crossRunWord.text.length) {
              continue;
            }
            const prefix = measureCrossRunPrefix(crossRunWord, breakOffset);
            if (prefix && prefix.width + measureTextWidth("-", prefix.hyphenStyle) <= spaceLeft) {
              fittingPrefix = prefix;
              break;
            }
          }
          if (fittingPrefix) {
            for (const segment of fittingPrefix.segments) {
              updateMaxFont(
                cjkLineHeightStyle(
                  segment.run,
                  complexScriptLineHeightStyle(segment.run, segment.style),
                ),
              );
            }
            currentLine.width +=
              leadingLetterSpacing +
              fittingPrefix.width +
              measureTextWidth("-", fittingPrefix.hyphenStyle);
            currentLine.trailingWhitespaceWidth = 0;
            currentLine.toRun = fittingPrefix.endRunIndex;
            currentLine.toChar = fittingPrefix.endChar;
            currentLine.discretionaryHyphen = { runIndex: fittingPrefix.endRunIndex };
            startNewLine(fittingPrefix.endRunIndex, fittingPrefix.endChar);
            if (fittingPrefix.endRunIndex === runIndex) {
              updateMaxFont(lineHeightStyle);
              charIndex = fittingPrefix.endChar;
              continue;
            }
            crossRunResume = {
              runIndex: fittingPrefix.endRunIndex,
              charIndex: fittingPrefix.endChar,
            };
            charIndex = text.length;
            continue;
          }
        }

        // If the word itself is longer than a line, hard-break by grapheme.
        // Use substring measurement (not char-by-char accumulation) to preserve
        // kerning accuracy. Char-by-char accumulation overestimates width by
        // ~1-2px per line due to lost kerning, causing extra wraps in narrow cells.
        if (wordWidth - hangingPunctuationWidth > currentLine.availableWidth + widthTolerance) {
          // Long word that needs hard-breaking. DON'T start a new line first —
          // fill the remaining space on the current line with as many characters
          // as possible. This prevents wasting a full line when a small run
          // (like "{" at 10pt) precedes a long word (like a variable at 5.5pt).
          let chunkStart = 0;
          let chunkLeadingLetterSpacing = leadingLetterSpacing;

          while (chunkStart < measuredWord.length) {
            const spaceLeft =
              currentLine.availableWidth -
              currentLine.width -
              chunkLeadingLetterSpacing +
              WIDTH_TOLERANCE;
            const remaining = measuredWord.slice(chunkStart);
            let bestEnd = findMaxFittingLength(
              remaining,
              style,
              spaceLeft,
              currentLine.width === 0,
              breakPolicy,
            );

            // Nothing fits beside existing content: start a new line and retry.
            // On an empty line findMaxFittingLength forces one whole grapheme,
            // even if that grapheme itself is wider than the line.
            if (bestEnd === 0) {
              startNewLine(runIndex, charIndex + chunkStart);
              chunkLeadingLetterSpacing = 0;
              updateMaxFont(lineHeightStyle);
              continue;
            }

            const chunkEnd = chunkStart + bestEnd;
            const chunk = measuredWord.slice(chunkStart, chunkEnd);
            const chunkWidth = measureTextWidth(chunk, style);

            currentLine.width += chunkLeadingLetterSpacing + chunkWidth;
            chunkLeadingLetterSpacing = 0;
            currentLine.trailingWhitespaceWidth =
              chunkWidth - measureTextWidth(trimTrailingSpacesAndTabs(chunk), style);
            currentLine.regularSpaceWidth += compressibleSpaceWidth(chunk, style);
            currentLine.toRun = runIndex;
            currentLine.toChar = charIndex + chunkEnd;

            chunkStart = chunkEnd;
            if (chunkStart < measuredWord.length) {
              startNewLine(runIndex, charIndex + chunkStart);
              updateMaxFont(lineHeightStyle);
            }
          }

          const trailingWhitespaceWidth = fullWordWidth - wordWidth;
          currentLine.width += trailingWhitespaceWidth;
          currentLine.trailingWhitespaceWidth = trailingWhitespaceWidth;
          currentLine.regularSpaceWidth += compressibleSpaceWidth(word, style);
          currentLine.toChar = nextBreak;

          charIndex = nextBreak;
          continue;
        }

        // Check if word fits on current line. If this is the last word in a
        // run and the next run starts without whitespace, include that glued
        // width in the wrap decision so a note marker or format-only word
        // split is not stranded on the next line (eigenpal/docx-editor#991).
        const wordTail = word.at(-1);
        const protectedGlueWidth = protectedCrossRunGlueWidths[runIndex] ?? 0;
        const rawGlueWidth =
          isRunTail &&
          wordTail !== undefined &&
          (!isBreakBetween(wordTail, followingTextLeads[runIndex]) || protectedGlueWidth > 0)
            ? Math.max(trailingGlueWidths[runIndex] ?? 0, protectedGlueWidth)
            : 0;
        const glueWidth =
          rawGlueWidth > 0 &&
          wordWidth + rawGlueWidth <= getPostWrapAvailableWidth() + widthTolerance
            ? rawGlueWidth
            : 0;
        const hangingAllowance = glueWidth === 0 ? hangingPunctuationWidth : 0;
        const candidateFit = isJustifiedParagraph
          ? resolveTextCandidateFit(
              isFinalTextCandidate(block, runIndex, nextBreak),
              currentLine.width + leadingLetterSpacing + wordWidth + glueWidth - hangingAllowance,
              currentLine.availableWidth,
              widthTolerance,
            )
          : undefined;
        const finalWrapTolerance = candidateFit?.tolerancePx ?? widthTolerance;
        // Let collapsible whitespace remain at the previous line's tail until
        // visible content decides whether to wrap. Starting a line from an
        // overflowed space creates whitespace-only soft-wrap lines.
        if (
          wordWidth > 0 &&
          currentLine.width > 0 &&
          currentLine.width + leadingLetterSpacing + wordWidth + glueWidth >
            currentLine.availableWidth + finalWrapTolerance + hangingAllowance
        ) {
          // Word doesn't fit, start new line
          startNewLine(runIndex, charIndex);
          leadingLetterSpacing = 0;
          // Re-apply font metrics to the new line (startNewLine resets maxFontSize)
          updateMaxFont(lineHeightStyle);
        }

        if (candidateFit?.type === "final-contraction-admitted") {
          currentLine.justificationPaint = candidateFit.paint;
        }

        // Add word to current line
        currentLine.width += leadingLetterSpacing + fullWordWidth;
        const wordTrailingWhitespaceWidth = fullWordWidth - wordWidth;
        // `findWordBreaks` yields each ASCII space separately, so consecutive
        // trailing segments must accumulate until visible content follows.
        currentLine.trailingWhitespaceWidth =
          wordWidth === 0
            ? currentLine.trailingWhitespaceWidth +
              leadingLetterSpacing +
              wordTrailingWhitespaceWidth
            : wordTrailingWhitespaceWidth;
        currentLine.regularSpaceWidth += compressibleSpaceWidth(word, style);
        currentLine.toRun = runIndex;
        currentLine.toChar = nextBreak;

        charIndex = nextBreak;
      }
    }
  }

  const listParagraphMarkFontSize = attrs?.listParagraphMarkFontSize;
  if (
    listParagraphMarkFontSize !== undefined &&
    listParagraphMarkFontSize > currentLine.maxFontSize
  ) {
    const fontFamily = currentLine.maxFontMetrics?.fontFamily ?? attrs?.defaultFontFamily;
    updateMaxFont({
      fontSize: listParagraphMarkFontSize,
      ...(fontFamily === undefined ? {} : { fontFamily }),
      ...(currentLine.maxFontMetrics === null && attrs?.defaultAlternateFontFamily !== undefined
        ? { alternateFontFamily: attrs.defaultAlternateFontFamily }
        : {}),
    });
  }

  // Finalize the last line
  finalizeLine();

  // Calculate total height — include floatSkipBefore from lines bumped past
  // floats so containers stay sized correctly.
  const totalHeight = lines.reduce((sum, line) => sum + measuredLineAdvance(line), 0);

  // Add spacing before/after
  let totalWithSpacing = totalHeight;
  if (spacing?.before) {
    totalWithSpacing += spacing.before;
  }
  if (spacing?.after) {
    totalWithSpacing += spacing.after;
  }

  return {
    kind: "paragraph",
    lines,
    totalHeight: totalWithSpacing,
  };
}

/**
 * Measure multiple paragraph blocks
 *
 * @param blocks - Array of paragraph blocks to measure
 * @param maxWidth - Maximum available width
 * @returns Array of ParagraphMeasure results
 */
export function measureParagraphs(blocks: ParagraphBlock[], maxWidth: number): ParagraphMeasure[] {
  return blocks.map((block) => measureParagraph(block, maxWidth));
}

/**
 * Get per-character widths for a text run (for click positioning)
 *
 * @param run - The text run to measure
 * @returns Array of character widths
 */
export function getRunCharWidths(run: TextRun): number[] {
  const style = runToFontStyle(run);
  const result = measureRun(run.text, style);
  return result.charWidths;
}
