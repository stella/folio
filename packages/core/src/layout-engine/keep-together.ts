/**
 * Keep Together Logic - Handle keepNext and keepLines paragraph properties
 *
 * DOCX paragraphs can have keepNext (keep with next paragraph) and keepLines
 * (keep all lines together) properties that affect pagination.
 */

import type { FlowBlock, ParagraphBlock, Measure } from "./types";

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

function isVisuallyEmptyParagraph(block: ParagraphBlock): boolean {
  if (block.runs.length === 0) {
    return true;
  }
  if (block.runs.length !== 1) {
    return false;
  }
  const run = block.runs.at(0);
  return run?.kind === "text" && run.text === "";
}

function startsTrailingTableSeparatorChain(blocks: FlowBlock[], index: number): boolean {
  const block = blocks[index];
  if (
    block?.kind !== "paragraph" ||
    !isVisuallyEmptyParagraph(block) ||
    blocks[index - 1]?.kind !== "table"
  ) {
    return false;
  }

  for (let nextIndex = index + 1; nextIndex < blocks.length; nextIndex++) {
    const nextBlock = blocks[nextIndex];
    if (nextBlock?.kind !== "paragraph") {
      return false;
    }
    if (!isVisuallyEmptyParagraph(nextBlock)) {
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
      if (nextPara.attrs?.keepNext || isVisuallyEmptyParagraph(nextPara)) {
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
 * Calculate the height needed to keep consecutive paragraph boundaries from
 * breaking across pages.
 *
 * Single-line and `keepLines` members are indivisible, so the reservation must
 * continue through them to the anchor's first line. A multi-line member without
 * `keepLines` may itself split; only its first line is needed to satisfy the
 * preceding member's `keepNext`, and the chain can stop there.
 */
export function calculateChainHeight(
  chain: KeepNextChain,
  blocks: FlowBlock[],
  measures: Measure[],
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

  let totalHeight = (firstBlock.attrs?.spacing?.before ?? 0) + firstMeasure.totalHeight;
  let trailingSpacing = firstBlock.attrs?.spacing?.after ?? 0;
  const startsWithTrailingTableSeparator =
    blocks[firstMemberIndex - 1]?.kind === "table" &&
    isVisuallyEmptyParagraph(firstBlock) &&
    !firstBlock.attrs?.keepNext;

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
    if (successorBlock?.kind !== "paragraph" || successorMeasure?.kind !== "paragraph") {
      return totalHeight;
    }

    totalHeight += Math.max(trailingSpacing, successorBlock.attrs?.spacing?.before ?? 0);
    const firstLine = successorMeasure.lines.at(0);
    if (!firstLine) {
      return totalHeight;
    }

    const isAnchor = index === successorIndices.length - 1 && chain.anchorIndex !== -1;
    const isSplittable = successorMeasure.lines.length > 1 && !successorBlock.attrs?.keepLines;
    if (
      isAnchor &&
      startsWithTrailingTableSeparator &&
      successorBlock.attrs?.widowControl !== false &&
      successorMeasure.lines.length > 1
    ) {
      const secondLine = successorMeasure.lines.at(1);
      return totalHeight + firstLine.lineHeight + (secondLine?.lineHeight ?? 0);
    }
    if (isAnchor || isSplittable) {
      return totalHeight + firstLine.lineHeight;
    }

    totalHeight += successorMeasure.totalHeight;
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
