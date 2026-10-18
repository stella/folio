/**
 * Run Consolidator - Merge consecutive runs with identical formatting
 *
 * DOCX files often contain many small runs with the same formatting,
 * created by Word for various reasons (spell checking, revision tracking,
 * cursor positioning, etc.). This causes:
 * - 252+ tiny <span> elements instead of a few
 * - Poor editing UX (cursor jumps between spans)
 * - Performance issues
 *
 * This module provides utilities to consolidate runs with identical
 * formatting into single runs, reducing fragmentation.
 */

import type { Run, ParagraphContent, Paragraph, Hyperlink } from "../types/document";
import { normalizeRunSegmentation } from "@stll/docx-core/ops";
import { cloneParagraphWithPropertySource } from "./paragraphPropertySource";

export { formattingEquals, isTextOnlyRun, runsMergeable } from "@stll/docx-core/ops";

/**
 * Consolidate an array of runs by merging consecutive runs with identical formatting
 *
 * @param runs - Array of runs to consolidate
 * @returns Consolidated array with fewer, larger runs
 */
export function consolidateRuns(runs: Run[]): Run[] {
  return normalizeRunSegmentation(runs) as Run[];
}

/**
 * Consolidate runs within a paragraph content array
 *
 * This handles the full paragraph structure, consolidating runs while
 * preserving hyperlinks, bookmarks, and fields as merge boundaries.
 */
export function consolidateParagraphContent(content: Hyperlink["children"]): Hyperlink["children"];
export function consolidateParagraphContent(content: ParagraphContent[]): ParagraphContent[];
export function consolidateParagraphContent(content: ParagraphContent[]): ParagraphContent[] {
  return normalizeRunSegmentation(content);
}

/**
 * Consolidate all runs within a paragraph
 *
 * @param paragraph - Paragraph to consolidate
 * @returns New paragraph with consolidated runs
 */
export function consolidateParagraph(paragraph: Paragraph): Paragraph {
  if (paragraph.content.length === 0) {
    return paragraph;
  }

  return cloneParagraphWithPropertySource(paragraph, {
    content: consolidateParagraphContent(paragraph.content),
  });
}

/**
 * Get the number of runs in a paragraph (for debugging/metrics)
 */
export function countRuns(paragraph: Paragraph): number {
  let count = 0;

  function countInContent(content: ParagraphContent[]): void {
    for (const item of content) {
      if (item.type === "run") {
        count++;
      } else if (item.type === "hyperlink") {
        countInContent(item.children);
      }
    }
  }

  countInContent(paragraph.content);

  return count;
}

/**
 * Calculate the consolidation ratio (reduction in number of runs)
 * Useful for debugging and metrics
 */
export function getConsolidationStats(
  originalParagraphs: Paragraph[],
  consolidatedParagraphs: Paragraph[],
): {
  originalRunCount: number;
  consolidatedRunCount: number;
  reductionPercentage: number;
} {
  const originalCount = originalParagraphs.reduce((sum, p) => sum + countRuns(p), 0);
  const consolidatedCount = consolidatedParagraphs.reduce((sum, p) => sum + countRuns(p), 0);

  const reduction =
    originalCount > 0 ? ((originalCount - consolidatedCount) / originalCount) * 100 : 0;

  return {
    originalRunCount: originalCount,
    consolidatedRunCount: consolidatedCount,
    reductionPercentage: Math.round(reduction * 10) / 10,
  };
}
