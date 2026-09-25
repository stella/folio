/**
 * Places a measured paragraph's lines onto pages: keep rules, widow/orphan
 * control, rendered-break reconciliation, and per-line footnote reservation.
 */

import { hasKeepLines } from "./keep-together";
import { measuredLineAdvance, measuredLineRangeHeight } from "./lineFlow";
import { FOOTNOTE_SEPARATOR_HEIGHT, createPaginator } from "./paginator";
import { getParagraphFragmentPmRange } from "./paragraphFragmentRange";
import {
  collapseParagraphSpacing,
  getParagraphSpacingAfter,
  getParagraphSpacingBefore,
} from "./paragraphSpacing";
import type { ParagraphBlock, ParagraphMeasure, ParagraphFragment } from "./types";
import {
  RENDERED_BREAK_REFLOW_TOLERANCE_LINES,
  projectedFootnoteReserveGrowth,
} from "./layoutFlowShared";

function hasWidowControl(block: ParagraphBlock): boolean {
  return block.attrs?.widowControl !== false;
}

/**
 * Footnote refs whose run sits in line `[fromRun..toRun]`, with their
 * pre-measured content heights. Empty when the engine isn't tracking
 * dynamic fn demand. Used by `layoutParagraph` to (a) reserve fn
 * height per line and (b) record the IDs on the host page so the
 * painter renders the fn on the page where the ref-bearing line
 * actually landed — even when the paragraph splits across pages
 * (fragment pmStart/pmEnd is paragraph-wide and cannot disambiguate
 * between split halves).
 */
function getLineFootnoteRefs(
  block: ParagraphBlock,
  fromRun: number,
  toRun: number,
  fnHeights: Map<number, number> | undefined,
): { ids: number[]; height: number } {
  if (!fnHeights) {
    return { ids: [], height: 0 };
  }
  const ids: number[] = [];
  let height = 0;
  for (let r = fromRun; r <= toRun; r++) {
    const run = block.runs[r];
    if (!run || run.kind !== "text") {
      continue;
    }
    const id = (run as { footnoteRefId?: number }).footnoteRefId;
    if (id === undefined) {
      continue;
    }
    const h = fnHeights.get(id);
    if (h !== undefined) {
      ids.push(id);
      height += h;
    }
  }
  return { ids, height };
}

/**
 * Layout a paragraph block onto pages.
 *
 * When `footnoteHeightById` is provided, each line carrying a footnote
 * ref additionally reserves space on its host page for the fn's
 * content. Pagination decisions include only the reservation growth
 * beyond any retry floor, so a line cannot land on a page that lacks
 * room for both body and note content.
 */
type LayoutParagraphOptions = {
  block: ParagraphBlock;
  measure: ParagraphMeasure;
  paginator: ReturnType<typeof createPaginator>;
  contentWidth: number;
  footnoteHeightById: Map<number, number> | undefined;
  suppressSpaceBefore: boolean;
};

export function layoutParagraph({
  block,
  measure,
  paginator,
  contentWidth,
  footnoteHeightById,
  suppressSpaceBefore,
}: LayoutParagraphOptions): void {
  const lines = measure.lines;
  if (lines.length === 0) {
    // Empty paragraph - still takes up space based on spacing
    const spaceBefore = suppressSpaceBefore ? 0 : getParagraphSpacingBefore(block);
    const spaceAfter = getParagraphSpacingAfter(block);
    const state = paginator.getCurrentState();

    // Create minimal fragment
    const fragment: ParagraphFragment = {
      kind: "paragraph",
      ...(block.attrs?.suppressEmptyParagraphHeight === true
        ? { paginationRole: "empty-carrier" as const }
        : {}),
      blockId: block.id,
      x: paginator.getColumnX(state.columnIndex),
      y: state.cursorY + spaceBefore,
      width: contentWidth,
      height: 0,
      fromLine: 0,
      toLine: 0,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
      ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
    };

    paginator.addFragment(fragment, 0, spaceBefore, spaceAfter);
    return;
  }

  const spaceBefore = suppressSpaceBefore ? 0 : getParagraphSpacingBefore(block);
  const spaceAfter = getParagraphSpacingAfter(block);

  // `w:keepLines` (§17.3.1.14) needs the whole paragraph's line and footnote
  // demand. Sum it once so the column retries below stay constant-time.
  let keepLinesDemand: { linesHeight: number; footnoteHeight: number } | undefined;
  if (lines.length > 1 && hasKeepLines(block)) {
    let footnoteHeight = 0;
    for (const line of lines) {
      footnoteHeight += getLineFootnoteRefs(
        block,
        line.fromRun,
        line.toRun,
        footnoteHeightById,
      ).height;
    }
    keepLinesDemand = {
      linesHeight: measuredLineRangeHeight(lines, 0, lines.length),
      footnoteHeight,
    };
  }
  let keepLinesMoveAttempted = false;

  // Try to fit all lines on current page/column
  let currentLineIndex = 0;

  while (currentLineIndex < lines.length) {
    const state = paginator.getCurrentState();
    const availableHeight = paginator.getAvailableHeight();
    const currentLine = lines.at(currentLineIndex);
    const markerAlreadySatisfied =
      state.columnIndex === 0 &&
      state.cursorY === state.topMargin &&
      state.page.fragments.length === 0;
    if (
      currentLine?.renderedPageBreakBefore === true &&
      !markerAlreadySatisfied &&
      availableHeight <= measuredLineAdvance(currentLine) * RENDERED_BREAK_REFLOW_TOLERANCE_LINES
    ) {
      paginator.forcePageBreak();
      continue;
    }

    // If the paragraph cannot begin on this page solely because its leading
    // spacing collapses with the previous block's trailing spacing — spacing
    // the next page sheds — advance first so the whole paragraph re-fits there.
    // Otherwise the `fittingLines === 0` fallback below strands its first line
    // here and `addFragment` carries it to the next page alone, splitting a
    // paragraph that fits whole into an artificial intra-page continuation
    // (eigenpal/docx-editor#782 follow-up).
    //
    // Only when the page already holds content (`cursorY !== topMargin`):
    // advancing then lands on a fresh page that resets `trailingSpacing` to 0,
    // so the re-entry takes the normal path. At a fresh page top, `ensureFits`
    // deliberately does not advance oversized content, so continuing there
    // would loop forever — fall through and let the fallback place the line.
    if (currentLineIndex === 0 && state.trailingSpacing > 0 && state.cursorY !== state.topMargin) {
      const firstLine = lines[0]!; // SAFETY: lines.length > 0 in while guard
      const firstLineRefs = getLineFootnoteRefs(
        block,
        firstLine.fromRun,
        firstLine.toRun,
        footnoteHeightById,
      );
      const firstLineHeight =
        measuredLineAdvance(firstLine) +
        projectedFootnoteReserveGrowth(state, firstLineRefs.height);
      const collapsedLead = collapseParagraphSpacing({
        before: spaceBefore,
        after: state.trailingSpacing,
      });
      const columnCapacity = state.contentBottom - state.topMargin;
      if (
        collapsedLead + firstLineHeight > availableHeight &&
        spaceBefore + firstLineHeight <= columnCapacity
      ) {
        paginator.ensureFits(collapsedLead + firstLineHeight);
        continue;
      }
    }

    // `w:keepLines` (§17.3.1.14): a paragraph that does not fit whole in the
    // rest of this column starts on the next one instead of splitting. One
    // taller than a full column cannot be kept together and splits normally.
    // The move is tried once; a paragraph still too tall for the column it
    // lands in splits there rather than retrying column by column.
    if (
      keepLinesDemand !== undefined &&
      !keepLinesMoveAttempted &&
      currentLineIndex === 0 &&
      state.cursorY !== state.topMargin
    ) {
      const wholeHeight =
        keepLinesDemand.linesHeight +
        projectedFootnoteReserveGrowth(state, keepLinesDemand.footnoteHeight);
      const collapsedLead = collapseParagraphSpacing({
        before: spaceBefore,
        after: state.trailingSpacing,
      });
      const columnCapacity = state.contentBottom - state.topMargin;
      if (
        collapsedLead + wholeHeight > availableHeight &&
        spaceBefore + wholeHeight <= columnCapacity
      ) {
        keepLinesMoveAttempted = true;
        paginator.ensureFits(collapsedLead + wholeHeight);
        continue;
      }
    }

    // Calculate how many lines fit
    let linesHeight = 0;
    let linesFnHeight = 0;
    const linesFnIds: number[] = [];
    let fittingLines = 0;
    let forcePageBreakAfterFragment = false;

    // The first fragment of a paragraph eats `spaceBefore` from the
    // available height for *every* line check, not only the first one.
    // Pre-fix the loop checked `linesHeight + lineHeight + spaceBefore`
    // only when `j === currentLineIndex`; subsequent lines compared bare
    // line totals against the full available height. That let the loop
    // claim more lines than would actually fit, then `addFragment` (which
    // correctly reserves collapsed spacing + line height) refused the placement
    // and bumped the *whole* fragment to the next page. Result: page-end
    // paragraphs with multi-line content didn't split — they jumped the
    // page boundary, leaving a chunk of empty space above.
    //
    // `addFragment` collapses `spaceBefore` with the previous block's trailing
    // spacing, so the fit loop must reserve the same amount; otherwise a large
    // preceding `spaceAfter` repeats the same over-count (eigenpal/docx-editor#782).
    const firstFragmentSpaceBefore =
      currentLineIndex === 0
        ? collapseParagraphSpacing({ before: spaceBefore, after: state.trailingSpacing })
        : 0;

    for (let j = currentLineIndex; j < lines.length; j++) {
      const line = lines[j]!; // SAFETY: j < lines.length
      const lineAdvance = measuredLineAdvance(line);
      if (
        j > currentLineIndex &&
        line.renderedPageBreakBefore === true &&
        availableHeight - (linesHeight + firstFragmentSpaceBefore + linesFnHeight) <=
          lineAdvance * RENDERED_BREAK_REFLOW_TOLERANCE_LINES
      ) {
        forcePageBreakAfterFragment = true;
        break;
      }
      const lineRefs = getLineFootnoteRefs(block, line.fromRun, line.toRun, footnoteHeightById);
      const totalWithLine = linesHeight + lineAdvance;
      const footnoteGrowth = projectedFootnoteReserveGrowth(state, linesFnHeight + lineRefs.height);
      const withSpacing = totalWithLine + firstFragmentSpaceBefore + footnoteGrowth;

      if (withSpacing <= availableHeight || fittingLines === 0) {
        linesHeight = totalWithLine;
        linesFnHeight += lineRefs.height;
        for (const id of lineRefs.ids) {
          linesFnIds.push(id);
        }
        fittingLines++;
      } else {
        break;
      }
    }

    let forceBreakAfterFragment = false;
    if (hasWidowControl(block)) {
      const remainingAfter = lines.length - (currentLineIndex + fittingLines);
      if (fittingLines > 1 && remainingAfter === 1) {
        if (currentLineIndex === 0 && fittingLines === 2 && state.cursorY !== state.topMargin) {
          paginator.forceColumnBreak();
          continue;
        }
        fittingLines -= 1;
        forceBreakAfterFragment = true;
        linesHeight = 0;
        linesFnHeight = 0;
        linesFnIds.length = 0;
        for (let j = currentLineIndex; j < currentLineIndex + fittingLines; j++) {
          const line = lines[j]!; // SAFETY: j is within adjusted fitting range
          linesHeight += measuredLineAdvance(line);
          const lineRefs = getLineFootnoteRefs(block, line.fromRun, line.toRun, footnoteHeightById);
          linesFnHeight += lineRefs.height;
          for (const id of lineRefs.ids) {
            linesFnIds.push(id);
          }
        }
      }
    }

    // Create fragment for these lines
    const isFirstFragment = currentLineIndex === 0;
    const isLastFragment = currentLineIndex + fittingLines >= lines.length;
    const effectiveSpaceBefore = isFirstFragment ? spaceBefore : 0;
    const effectiveSpaceAfter = isLastFragment ? spaceAfter : 0;

    const pmRange = getParagraphFragmentPmRange(
      block,
      measure,
      currentLineIndex,
      currentLineIndex + fittingLines,
    );

    const fragment: ParagraphFragment = {
      kind: "paragraph",
      ...(block.attrs?.suppressEmptyParagraphHeight === true
        ? { paginationRole: "empty-carrier" as const }
        : {}),
      blockId: block.id,
      x: paginator.getColumnX(state.columnIndex),
      y: 0, // Will be set by addFragment
      width: contentWidth,
      height: linesHeight,
      fromLine: currentLineIndex,
      toLine: currentLineIndex + fittingLines,
      ...(pmRange.pmStart !== undefined ? { pmStart: pmRange.pmStart } : {}),
      ...(pmRange.pmEnd !== undefined ? { pmEnd: pmRange.pmEnd } : {}),
      ...(!isFirstFragment ? { continuesFromPrev: true } : {}),
      ...(!isLastFragment ? { continuesOnNext: true } : {}),
      ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
    };

    // Ensure the page can accommodate body lines + footnote demand
    // *together* before placing. Without this, the `fittingLines === 0`
    // fallback in the loop above may force a line carrying a fn ref
    // onto a page that has only a hair of body space left — body fits
    // by itself but `addFootnoteHeight` afterwards drops contentBottom
    // below cursorY, producing an overlap with the fn area.
    //
    // Two-phase check (Codex PR #258 reviews — both edges):
    //
    // 1. Try without the separator overhead. The current page may
    //    already host a footnote — in that case `addFootnoteHeight`
    //    will *not* reserve another separator, so adding 13 px here
    //    would force an unnecessary page advance and split the
    //    paragraph between pages even though the line + fn still fit.
    //
    // 2. After the first `ensureFits`, re-read the page state. If it has
    //    no dynamic footnote demand, the next `addFootnoteHeight` call
    //    includes a separator. Verify that remaining growth still fits;
    //    a retry floor may already cover some or all of it.
    if (linesFnHeight > 0) {
      const stateBefore = paginator.getCurrentState();
      paginator.ensureFits(
        effectiveSpaceBefore +
          linesHeight +
          projectedFootnoteReserveGrowth(stateBefore, linesFnHeight),
      );
      const stateAfter = paginator.getCurrentState();
      if (stateAfter.footnoteDemandHeight === 0) {
        paginator.ensureFits(
          effectiveSpaceBefore +
            linesHeight +
            projectedFootnoteReserveGrowth(stateAfter, linesFnHeight + FOOTNOTE_SEPARATOR_HEIGHT),
        );
      }
    }

    const result = paginator.addFragment(
      fragment,
      linesHeight,
      effectiveSpaceBefore,
      effectiveSpaceAfter,
    );
    fragment.y = result.y;

    // Now that the lines have committed to this page, grow the page's
    // footnote reservation for any fn refs they carry, and record the
    // IDs on the host page directly. Page → fn-ID mapping is driven by
    // line-level placement here (not by post-layout pmRange mapping)
    // so a fn ref that lives in a continuation fragment of a split
    // paragraph is correctly attributed to the page where the
    // ref-bearing line landed (Codex PR #258 review).
    if (linesFnHeight > 0) {
      paginator.addFootnoteHeight(linesFnHeight, linesFnIds);
    }

    currentLineIndex += fittingLines;

    // If more lines remain, advance to next column/page
    if (currentLineIndex < lines.length) {
      if (forcePageBreakAfterFragment) {
        paginator.forcePageBreak();
      } else if (forceBreakAfterFragment) {
        paginator.forceColumnBreak();
      } else {
        paginator.ensureFits(measuredLineAdvance(lines[currentLineIndex]!)); // SAFETY: guarded by length check
      }
    }
  }
}
