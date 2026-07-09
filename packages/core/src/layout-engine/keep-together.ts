/**
 * Keep Together Logic - Handle keepNext and keepLines paragraph properties
 *
 * DOCX paragraphs can have keepNext (keep with next paragraph) and keepLines
 * (keep all lines together) properties that affect pagination.
 */

import type { FlowBlock, ParagraphBlock, Measure } from "./types";

/**
 * A chain of consecutive keepNext paragraphs.
 */
export type KeepNextChain = {
  /** Index of the first paragraph in the chain. */
  startIndex: number;
  /** Index of the last paragraph in the chain. */
  endIndex: number;
  /** All paragraph indices in the chain. */
  memberIndices: number[];
  /** Index of the anchor paragraph (first non-keepNext after chain), or -1 if none. */
  anchorIndex: number;
};

/**
 * Pre-scan blocks to find all keepNext chains.
 *
 * A keepNext chain is a sequence of consecutive paragraphs with keepNext=true,
 * followed by an anchor paragraph (the first non-keepNext paragraph).
 * The entire chain must stay on the same page as the anchor's first line.
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

    const para = block as ParagraphBlock;
    // Skip paragraphs without keepNext
    if (!para.attrs?.keepNext) {
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

      const nextPara = nextBlock as ParagraphBlock;
      if (nextPara.attrs?.keepNext) {
        // Continue the chain
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
 * Calculate the height needed to keep the chain's first paragraph with the
 * first line of its successor.
 *
 * `keepNext` prevents a break between paragraphs; it does not imply
 * `keepLines`. A later chain member may therefore split across pages. Reserving
 * every line of every member moves otherwise splittable legal clauses as one
 * oversized unit.
 */
export function calculateChainHeight(
  chain: KeepNextChain,
  blocks: FlowBlock[],
  measures: Measure[],
): number {
  let totalHeight = 0;

  const firstMemberIndex = chain.memberIndices.at(0);
  if (firstMemberIndex === undefined) {
    return 0;
  }
  const firstBlock = blocks[firstMemberIndex];
  const firstMeasure = measures[firstMemberIndex];
  if (firstBlock?.kind !== "paragraph" || firstMeasure?.kind !== "paragraph") {
    return 0;
  }

  totalHeight += firstBlock.attrs?.spacing?.before ?? 0;
  totalHeight += firstMeasure.totalHeight;
  totalHeight += firstBlock.attrs?.spacing?.after ?? 0;

  const successorIndex = chain.memberIndices.at(1) ?? chain.anchorIndex;
  const successorMeasure = successorIndex === -1 ? undefined : measures[successorIndex];
  if (successorMeasure?.kind === "paragraph") {
    const firstSuccessorLine = successorMeasure.lines.at(0);
    if (firstSuccessorLine) {
      totalHeight += firstSuccessorLine.lineHeight;
    }
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
  const para = block as ParagraphBlock;
  return para.attrs?.keepLines === true;
}

/**
 * Check if a paragraph should start on a new page (pageBreakBefore).
 */
export function hasPageBreakBefore(block: FlowBlock): boolean {
  if (block.kind !== "paragraph") {
    return false;
  }
  const para = block as ParagraphBlock;
  return para.attrs?.pageBreakBefore === true;
}
