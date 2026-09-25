/**
 * Keep Together Logic - Handle keepNext and keepLines paragraph properties
 *
 * DOCX paragraphs can have keepNext (keep with next paragraph) and keepLines
 * (keep all lines together) properties that affect pagination.
 */

import { measuredLineRangeHeight } from "./lineFlow";
import {
  collapseParagraphSpacing,
  isAuthoredEmptyParagraph,
  isEmptyParagraph,
} from "./paragraphSpacing";
import type {
  FlowBlock,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
  TableBlock,
  TableMeasure,
} from "./types";

/** Lines §17.3.1.44 widow control keeps on each side of a paragraph split. */
const MIN_WIDOW_CONTROL_SPLIT_LINES = 2;

/**
 * A chain of paragraphs that Word keeps with following content. This includes
 * explicit keepNext links and a trailing table separator that would otherwise
 * be stranded at the bottom of a page.
 */
export type KeepNextChain = {
  /** Index of the first paragraph in the chain. */
  startIndex: number;
  /** Index of the last keepNext or pass-through empty member. */
  endIndex: number;
  /** All keepNext and pass-through empty paragraph indices in the chain. */
  memberIndices: number[];
  /** Index of the anchor paragraph (first non-keepNext after chain), or -1 if none. */
  anchorIndex: number;
};

/**
 * Whether an empty paragraph is only a structural separator through which a
 * keep-with-next chain may pass. An explicitly styled or directly formatted
 * blank is authored layout content: it anchors the preceding keepNext rather
 * than implicitly linking that heading to later content.
 */
function isKeepNextPassThroughParagraph(block: ParagraphBlock): boolean {
  if (!isEmptyParagraph(block)) {
    return false;
  }
  if (block.attrs?.suppressEmptyParagraphHeight === true) {
    return true;
  }
  return !isAuthoredEmptyParagraph(block);
}

function startsTrailingTableSeparatorChain(blocks: FlowBlock[], index: number): boolean {
  const block = blocks[index];
  if (
    block?.kind !== "paragraph" ||
    !isKeepNextPassThroughParagraph(block) ||
    blocks[index - 1]?.kind !== "table"
  ) {
    return false;
  }

  for (let nextIndex = index + 1; nextIndex < blocks.length; nextIndex++) {
    const nextBlock = blocks[nextIndex];
    if (nextBlock?.kind !== "paragraph") {
      return false;
    }
    if (!isKeepNextPassThroughParagraph(nextBlock)) {
      return true;
    }
  }

  return false;
}

/**
 * Pre-scan blocks to find all keepNext chains.
 *
 * A chain starts with a paragraph whose keepNext=true or with an empty
 * separator immediately following a table. It continues through further
 * keepNext paragraphs and structural empty separators. The first visible
 * non-keepNext paragraph is its anchor.
 *
 * Returns a map from chain start index to chain info.
 */
export function computeKeepNextChains(blocks: FlowBlock[]): Map<number, KeepNextChain> {
  const chains = new Map<number, KeepNextChain>();
  const processed = new Set<number>();

  for (let i = 0; i < blocks.length; i++) {
    // Skip already-processed blocks (mid-chain members)
    if (processed.has(i)) {
      continue;
    }

    // SAFETY: i is bounded by blocks.length
    const block = blocks[i]!;
    // Only paragraphs can have keepNext
    if (block.kind !== "paragraph") {
      continue;
    }

    const para = block;
    // Word carries a trailing table separator to the next visible paragraph
    // instead of leaving the blank at the bottom of the preceding page.
    if (!para.attrs?.keepNext && !startsTrailingTableSeparatorChain(blocks, i)) {
      continue;
    }

    // Found a keepNext paragraph - scan forward to find full chain
    const memberIndices: number[] = [i];
    let endIndex = i;

    for (let j = i + 1; j < blocks.length; j++) {
      // SAFETY: j is bounded by blocks.length
      const nextBlock = blocks[j]!;

      // Breaks terminate the chain
      if (
        nextBlock.kind === "sectionBreak" ||
        nextBlock.kind === "pageBreak" ||
        nextBlock.kind === "columnBreak"
      ) {
        break;
      }

      // Non-paragraphs terminate the chain
      if (nextBlock.kind !== "paragraph") {
        break;
      }

      const nextPara = nextBlock;
      if (nextPara.attrs?.keepNext || isKeepNextPassThroughParagraph(nextPara)) {
        // Word carries keepNext across structural empty paragraphs. Treat
        // them as pass-through members so a blank separator cannot strand a
        // heading at the bottom of the preceding page.
        memberIndices.push(j);
        endIndex = j;
        processed.add(j);
      } else {
        // Found the anchor - stop here
        break;
      }
    }

    // Find the anchor (first paragraph after the chain)
    const potentialAnchor = endIndex + 1;
    let anchorIndex = -1;

    if (potentialAnchor < blocks.length) {
      // SAFETY: potentialAnchor < blocks.length
      const anchorBlock = blocks[potentialAnchor]!;
      // Anchor must not be a break
      if (
        anchorBlock.kind !== "sectionBreak" &&
        anchorBlock.kind !== "pageBreak" &&
        anchorBlock.kind !== "columnBreak"
      ) {
        anchorIndex = potentialAnchor;
      }
    }

    // Record the chain
    chains.set(i, {
      startIndex: i,
      endIndex,
      memberIndices,
      anchorIndex,
    });
  }

  return chains;
}

/**
 * Lines of a paragraph that must share a page with the paragraph before it for
 * a `w:keepNext` link (§17.3.1.15) to hold.
 *
 * - `w:keepLines` (§17.3.1.14): the paragraph never splits, so all of it.
 * - `w:widowControl` (§17.3.1.44): a split leaves at least two lines on each
 *   side, so a paragraph shorter than four lines cannot split at all and a
 *   longer one needs its first two lines.
 * - Otherwise a single line.
 */
function minimumOpeningLineCount(block: ParagraphBlock, measure: ParagraphMeasure): number {
  const lineCount = measure.lines.length;
  if (block.attrs?.keepLines === true) {
    return lineCount;
  }
  if (block.attrs?.widowControl !== false) {
    return lineCount < MIN_WIDOW_CONTROL_SPLIT_LINES * 2
      ? lineCount
      : MIN_WIDOW_CONTROL_SPLIT_LINES;
  }
  return Math.min(1, lineCount);
}

/**
 * Calculate the height needed to keep consecutive paragraph boundaries from
 * breaking across pages.
 *
 * Each successor (later member or the anchor) must bring its minimum
 * unbreakable opening onto the page: see {@link minimumOpeningLineCount}. When
 * that opening is the whole paragraph, the successor is indivisible and the
 * reservation continues through it; otherwise the chain stops after the
 * opening, because a later split cannot separate it from its predecessor.
 *
 * A table anchor contributes `tableOpeningHeight`: the part of the table
 * that must start on the same page as the paragraph before it.
 *
 * A paragraph measure's `totalHeight` already includes its own spacing before
 * and after. The chain accounts for spacing separately (collapsing each gap to
 * the larger side, as the paginator does), so members contribute only their
 * line advances; otherwise every gap would be reserved twice.
 */
export function calculateChainHeight(
  chain: KeepNextChain,
  blocks: FlowBlock[],
  measures: Measure[],
  incomingSpacing = 0,
  tableOpeningHeight?: (block: TableBlock, measure: TableMeasure) => number,
): number {
  const firstMemberIndex = chain.memberIndices.at(0);
  if (firstMemberIndex === undefined) {
    return 0;
  }
  const firstBlock = blocks[firstMemberIndex];
  const firstMeasure = measures[firstMemberIndex];
  if (firstBlock?.kind !== "paragraph" || firstMeasure?.kind !== "paragraph") {
    return 0;
  }

  let totalHeight =
    collapseParagraphSpacing({
      before: firstBlock.attrs?.spacing?.before ?? 0,
      after: incomingSpacing,
    }) + measuredLineRangeHeight(firstMeasure.lines, 0, firstMeasure.lines.length);
  let trailingSpacing = firstBlock.attrs?.spacing?.after ?? 0;

  const successorIndices = [...chain.memberIndices.slice(1)];
  if (chain.anchorIndex !== -1) {
    successorIndices.push(chain.anchorIndex);
  }

  for (let index = 0; index < successorIndices.length; index++) {
    const successorIndex = successorIndices[index];
    if (successorIndex === undefined) {
      continue;
    }
    const successorBlock = blocks[successorIndex];
    const successorMeasure = measures[successorIndex];
    if (
      tableOpeningHeight !== undefined &&
      successorBlock?.kind === "table" &&
      successorMeasure?.kind === "table" &&
      successorBlock.floating === undefined
    ) {
      // A table has no paragraph spacing of its own: the gap before it is the
      // previous paragraph's space after.
      return totalHeight + trailingSpacing + tableOpeningHeight(successorBlock, successorMeasure);
    }
    if (successorBlock?.kind !== "paragraph" || successorMeasure?.kind !== "paragraph") {
      return totalHeight;
    }

    totalHeight += collapseParagraphSpacing({
      before: successorBlock.attrs?.spacing?.before ?? 0,
      after: trailingSpacing,
    });
    const lineCount = successorMeasure.lines.length;
    if (lineCount === 0) {
      return totalHeight;
    }

    const isAnchor = index === successorIndices.length - 1 && chain.anchorIndex !== -1;
    const openingLines = minimumOpeningLineCount(successorBlock, successorMeasure);
    totalHeight += measuredLineRangeHeight(successorMeasure.lines, 0, openingLines);
    if (isAnchor || openingLines < lineCount) {
      return totalHeight;
    }

    trailingSpacing = successorBlock.attrs?.spacing?.after ?? 0;
  }

  return totalHeight;
}

/**
 * Get the set of indices that are mid-chain (not chain starters).
 * These should skip the keepNext check since their chain starter already decided.
 */
export function getMidChainIndices(chains: Map<number, KeepNextChain>): Set<number> {
  const midChain = new Set<number>();

  for (const chain of chains.values()) {
    // All members except the first are mid-chain
    for (let i = 1; i < chain.memberIndices.length; i++) {
      // SAFETY: i is bounded by chain.memberIndices.length
      midChain.add(chain.memberIndices[i]!);
    }
  }

  return midChain;
}

/**
 * Check if a paragraph has keepLines property (all lines must stay together).
 */
export function hasKeepLines(block: FlowBlock): boolean {
  if (block.kind !== "paragraph") {
    return false;
  }
  const para = block;
  return para.attrs?.keepLines === true;
}

/**
 * Check if a paragraph should start on a new page (pageBreakBefore).
 */
export function hasPageBreakBefore(block: FlowBlock): boolean {
  if (block.kind !== "paragraph") {
    return false;
  }
  const para = block;
  return para.attrs?.pageBreakBefore === true;
}
