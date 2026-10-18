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
import { mergeRunContent, runsMergeable } from "@stll/docx-core/ops";
import { cloneParagraphWithPropertySource } from "./paragraphPropertySource";
import { runHoldsPayload } from "./runPayload";

export { formattingEquals, isTextOnlyRun, runsMergeable } from "@stll/docx-core/ops";

/**
 * Consolidate an array of runs by merging consecutive runs with identical formatting
 *
 * @param runs - Array of runs to consolidate
 * @returns Consolidated array with fewer, larger runs
 */
export function consolidateRuns(runs: Run[]): Run[] {
  if (runs.length <= 1) {
    return runs;
  }

  const result: Run[] = [];
  let current: Run | null = null;

  for (const run of runs) {
    // A run holding no payload is a merge boundary, and it is kept: the keep
    // rule has already decided that this run exists, and it read the source
    // element to decide it. Such a run reaches here because a later pass will
    // supply its payload — `enrichParagraphTextBoxes` matches the text box it
    // lifted to the empty run that carried it — so dropping it here lost the
    // carrier's own `w:rPr` and could move the box off its position. Merging
    // it away is the same loss by another route, hence the flush.
    if (!runHoldsPayload(run)) {
      if (current !== null) {
        result.push(current);
        current = null;
      }
      result.push(run);
      continue;
    }

    // If no current run, start with this one
    if (current === null) {
      current = { ...run, content: [...run.content] };
      continue;
    }

    if (runsMergeable(current, run)) {
      // Spread the survivor rather than rebuilding it from a field list: the
      // predicate has already established that every record a merged run can
      // hold only one of is equal on both sides, so the one it keeps is the
      // one both wrote, and no field can go missing by omission here.
      current = {
        ...current,
        content: mergeRunContent(current.content, run.content),
      };
    } else {
      // Can't merge - save current and start new
      result.push(current);
      current = { ...run, content: [...run.content] };
    }
  }

  // Don't forget the last run
  if (current !== null) {
    result.push(current);
  }

  return result;
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
  const result: ParagraphContent[] = [];
  const pendingRuns: Run[] = [];

  function flushRuns(): void {
    if (pendingRuns.length > 0) {
      const consolidated = consolidateRuns(pendingRuns);
      result.push(...consolidated);
      pendingRuns.length = 0;
    }
  }

  for (const item of content) {
    if (item.type === "run") {
      pendingRuns.push(item);
    } else {
      // Non-run content acts as a merge boundary
      flushRuns();

      // Handle hyperlinks - consolidate their internal runs
      if (item.type === "hyperlink") {
        const hyperlink: Hyperlink = {
          ...item,
          children: consolidateParagraphContent(item.children),
        };
        result.push(hyperlink);
      } else {
        result.push(item);
      }
    }
  }

  // Flush any remaining runs
  flushRuns();

  return result;
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
