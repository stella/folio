/**
 * Section handling during pagination: per-section layout configuration,
 * section-break scheduling (in place, next page, odd/even filler pages), and
 * balanced column height for continuous sections.
 */

import { panic } from "better-result";
import type { SectionStart } from "@stll/docx-core/model";
import { measuredLineAdvance } from "./lineFlow";
import { SECTION_START_PLACEMENT, createPaginator } from "./paginator";
import type { PageState } from "./paginator";
import {
  collapseParagraphSpacing,
  getParagraphSpacingAfter,
  getParagraphSpacingBefore,
} from "./paragraphSpacing";
import { physicalColumnRegionIsShared } from "./section-breaks";
import type {
  FlowBlock,
  Measure,
  PageMargins,
  ColumnLayout,
  SectionBreakBlock,
  SectionPageNumbering,
} from "./types";

export type SectionLayoutConfig = {
  pageSize: { w: number; h: number };
  margins: PageMargins;
  pageNumbering: SectionPageNumbering;
  columns?: ColumnLayout;
};

export const DEFAULT_COLUMNS: ColumnLayout = { count: 1, gap: 0 };
export const CONTINUE_PAGE_NUMBERING: SectionPageNumbering = { type: "continue" };

export function collectSectionConfigs(
  blocks: FlowBlock[],
  initialConfig: SectionLayoutConfig,
  finalConfig: SectionLayoutConfig,
): {
  configs: SectionLayoutConfig[];
  breakIndices: number[];
} {
  const configs: SectionLayoutConfig[] = [];
  const breakIndices: number[] = [];
  let previousConfig = initialConfig;

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    if (block?.kind !== "sectionBreak") {
      continue;
    }

    const sectionBreak = block as SectionBreakBlock;
    const config: SectionLayoutConfig = {
      pageSize: sectionBreak.pageSize ?? previousConfig.pageSize,
      margins: sectionBreak.margins ?? previousConfig.margins,
      pageNumbering: sectionBreak.pageNumbering ?? CONTINUE_PAGE_NUMBERING,
    };
    if (sectionBreak.columns !== undefined) {
      config.columns = sectionBreak.columns;
    }
    configs.push(config);
    breakIndices.push(index);
    previousConfig = configs.at(-1) ?? previousConfig;
  }

  configs.push(finalConfig);
  return { configs, breakIndices };
}

/**
 * Estimate the height of a short paragraph-only multi-column section so its
 * final page can use Word-style balanced columns. Sections containing tables,
 * floating blocks, or authored breaks keep the normal bottom-up pagination;
 * those need structure-aware balancing rather than a height estimate.
 */
type BalancedParagraphSectionHeightOptions = {
  blocks: FlowBlock[];
  measures: Measure[];
  startIndex: number;
  endIndex: number;
  incomingSpacing: number;
  columnCount: number;
  availableHeight: number;
};

export function balancedParagraphSectionHeight({
  blocks,
  measures,
  startIndex,
  endIndex,
  incomingSpacing,
  columnCount,
  availableHeight,
}: BalancedParagraphSectionHeightOptions): number | undefined {
  if (columnCount <= 1 || startIndex >= endIndex || availableHeight <= 0) {
    return undefined;
  }

  const lineUnits: number[] = [];
  let totalHeight = 0;
  let tallestUnit = 0;
  let trailingSpacing = incomingSpacing;
  for (let index = startIndex; index < endIndex; index++) {
    const block = blocks[index];
    const measure = measures[index];
    if (block?.kind !== "paragraph" || measure?.kind !== "paragraph") {
      return undefined;
    }
    if (
      block.attrs?.keepNext === true ||
      block.attrs?.keepLines === true ||
      block.runs.some((run) => run.kind === "text" && run.footnoteRefId !== undefined)
    ) {
      return undefined;
    }

    const leadingSpacing = collapseParagraphSpacing({
      before: getParagraphSpacingBefore(block),
      after: trailingSpacing,
    });
    if (measure.lines.length === 0) {
      if (leadingSpacing > 0) {
        lineUnits.push(leadingSpacing);
        totalHeight += leadingSpacing;
        tallestUnit = Math.max(tallestUnit, leadingSpacing);
      }
    }
    for (let lineIndex = 0; lineIndex < measure.lines.length; lineIndex++) {
      const line = measure.lines[lineIndex];
      if (!line) {
        continue;
      }
      const lineHeight = measuredLineAdvance(line);
      const unitHeight = lineHeight + (lineIndex === 0 ? leadingSpacing : 0);
      lineUnits.push(unitHeight);
      totalHeight += unitHeight;
      tallestUnit = Math.max(tallestUnit, unitHeight);
    }
    if (tallestUnit > availableHeight || totalHeight > availableHeight * columnCount) {
      return undefined;
    }
    trailingSpacing = getParagraphSpacingAfter(block);
  }

  if (totalHeight <= 0) {
    return undefined;
  }

  const columnsNeeded = (targetHeight: number): number => {
    let usedHeight = 0;
    let usedColumns = 1;
    for (const unitHeight of lineUnits) {
      if (usedHeight > 0 && usedHeight + unitHeight > targetHeight) {
        usedColumns += 1;
        usedHeight = unitHeight;
      } else {
        usedHeight += unitHeight;
      }
    }
    return usedColumns;
  };

  let lower = Math.max(tallestUnit, totalHeight / columnCount);
  let upper = Math.min(totalHeight, availableHeight);
  if (columnsNeeded(upper) > columnCount) {
    return undefined;
  }
  for (let iteration = 0; iteration < 32; iteration++) {
    const middle = (lower + upper) / 2;
    if (columnsNeeded(middle) <= columnCount) {
      upper = middle;
    } else {
      lower = middle;
    }
  }
  return Math.ceil(upper * 1000) / 1000;
}

/**
 * Start the next section in the current page region, below the outgoing
 * content.
 *
 * ECMA-376 §17.6.22: a `continuous` break normally keeps the current page
 * geometry and defers the new size/margins to the next natural page break. But
 * a break that changes page size or orientation cannot share a physical sheet
 * with the preceding section, so Word and LibreOffice promote it to a page
 * break (eigenpal/docx-editor#841).
 *
 * Compare against the last laid-out page without materializing one: a break
 * before any content has no sheet to share, so it defers (the first content
 * then opens a page with the new geometry) rather than stranding a blank
 * leading page.
 */
function startSectionInPlace(
  paginator: ReturnType<typeof createPaginator>,
  nextSectionConfig: SectionLayoutConfig,
  nextSectionIndex: number | undefined,
): void {
  const currentPage = paginator.states.at(-1)?.page;
  const nextSize = nextSectionConfig.pageSize;
  const pageSizeChanges =
    currentPage != null &&
    (Math.round(nextSize.w) !== Math.round(currentPage.size.w) ||
      Math.round(nextSize.h) !== Math.round(currentPage.size.h));
  if (nextSectionIndex !== undefined) {
    paginator.startSection({
      sectionIndex: nextSectionIndex,
      pageNumbering: nextSectionConfig.pageNumbering,
      placement: pageSizeChanges
        ? SECTION_START_PLACEMENT.NEXT_PAGE
        : SECTION_START_PLACEMENT.CONTINUOUS,
    });
  }
  if (pageSizeChanges) {
    // Promote to a page break, but reuse an already blank current page as
    // the next section's first page instead of leaving it stranded.
    paginator.updatePageLayout(nextSize, nextSectionConfig.margins);
    if (!paginator.retargetCurrentBlankPage()) {
      paginator.forcePageBreak({ coalesceBlankPage: true });
    }
    return;
  }
  paginator.updatePageLayout(nextSize, nextSectionConfig.margins, false);
  paginator.retargetCurrentBlankPage();
}

/**
 * Whether a section opening on `target` would give two consecutive sheets page
 * numbers that are both odd or both even under `w:evenAndOddHeaders` (§17.10.1). Odd and
 * even page numbers alternate between right- and left-hand sheets, so a
 * `w:pgNumType w:start` restart that repeats whether the previous sheet's number is odd or even
 * needs a blank sheet between them.
 */
function sectionStartRepeatsOddEven(
  paginator: ReturnType<typeof createPaginator>,
  target: PageState,
): boolean {
  if (target.page.headerFooterRefs?.evenAndOddHeaders !== true) {
    return false;
  }
  const previous = paginator.states.at(-2)?.page;
  if (previous === undefined || previous.number !== target.page.number - 1) {
    return false;
  }
  return previous.logicalNumber % 2 === target.page.logicalNumber % 2;
}

/**
 * Keep the blank sheet the section break opened as a filler with no page
 * furniture, and open the section again on the following sheet so its first
 * page keeps the restarted page number.
 */
function insertOddEvenFillerPage(
  paginator: ReturnType<typeof createPaginator>,
  filler: PageState,
  sectionIndex: number,
  sectionConfig: SectionLayoutConfig,
): void {
  // A page that selects no header or footer part paints none.
  filler.page.headerFooterRefs = {};
  paginator.startSection({ sectionIndex, pageNumbering: sectionConfig.pageNumbering });
  paginator.forcePageBreak();
}

/**
 * Handle a section break block.
 * @param block - The section break block (current section's properties)
 * @param paginator - The paginator instance
 * @param nextSectionConfig - Page layout for the NEXT section
 * @param nextSectionType - Break type of the NEXT section (how it starts relative to current)
 */
export function handleSectionBreak(
  _block: SectionBreakBlock,
  paginator: ReturnType<typeof createPaginator>,
  nextSectionConfig: SectionLayoutConfig,
  nextSectionType: SectionStart,
  nextSectionIndex?: number,
): void {
  switch (nextSectionType) {
    case "nextPage": {
      paginator.updatePageLayout(nextSectionConfig.pageSize, nextSectionConfig.margins);
      if (nextSectionIndex !== undefined) {
        paginator.startSection({
          sectionIndex: nextSectionIndex,
          pageNumbering: nextSectionConfig.pageNumbering,
        });
      }
      const target = paginator.forcePageBreak({ coalesceBlankPage: true });
      if (nextSectionIndex !== undefined && sectionStartRepeatsOddEven(paginator, target)) {
        insertOddEvenFillerPage(paginator, target, nextSectionIndex, nextSectionConfig);
      }
      break;
    }

    case "evenPage": {
      const target = paginator.forcePageBreak({ coalesceBlankPage: true });
      if (target.page.number % 2 !== 0) {
        paginator.forcePageBreak();
      }
      paginator.updatePageLayout(nextSectionConfig.pageSize, nextSectionConfig.margins);
      if (nextSectionIndex !== undefined) {
        paginator.startSection({
          sectionIndex: nextSectionIndex,
          pageNumbering: nextSectionConfig.pageNumbering,
        });
      }
      if (!paginator.retargetCurrentBlankPage()) {
        panic("Even-page section target must be blank");
      }
      break;
    }

    case "oddPage": {
      const target = paginator.forcePageBreak({ coalesceBlankPage: true });
      if (target.page.number % 2 === 0) {
        paginator.forcePageBreak();
      }
      paginator.updatePageLayout(nextSectionConfig.pageSize, nextSectionConfig.margins);
      if (nextSectionIndex !== undefined) {
        paginator.startSection({
          sectionIndex: nextSectionIndex,
          pageNumbering: nextSectionConfig.pageNumbering,
        });
      }
      if (!paginator.retargetCurrentBlankPage()) {
        panic("Odd-page section target must be blank");
      }
      break;
    }

    case "nextColumn": {
      // ECMA-376 Part 1 §17.18.77: the section begins in the next column.
      // Only a multi-column region has one, and only a section that repeats
      // that geometry can continue into it; otherwise there is no column
      // boundary to honour and the break degrades to `continuous`, which is
      // what Word does with a `nextColumn` break in single-column text.
      const currentPage = paginator.states.at(-1)?.page;
      if (
        currentPage === undefined ||
        !physicalColumnRegionIsShared(
          {
            pageSize: currentPage.size,
            margins: currentPage.margins,
            columns: paginator.columns,
          },
          nextSectionConfig,
        )
      ) {
        startSectionInPlace(paginator, nextSectionConfig, nextSectionIndex);
        break;
      }
      if (nextSectionIndex !== undefined) {
        paginator.startSection({
          sectionIndex: nextSectionIndex,
          pageNumbering: nextSectionConfig.pageNumbering,
          placement: SECTION_START_PLACEMENT.CONTINUOUS,
        });
      }
      paginator.updatePageLayout(nextSectionConfig.pageSize, nextSectionConfig.margins, false);
      paginator.forceColumnBreak();
      // The section carries on in the region it just advanced into, so it must
      // not be restarted below that region the way `updateColumns` restarts one.
      return;
    }

    case "continuous":
      startSectionInPlace(paginator, nextSectionConfig, nextSectionIndex);
      break;

    default:
      nextSectionType satisfies never;
  }

  // Update column layout for the next section
  paginator.updateColumns(nextSectionConfig.columns ?? DEFAULT_COLUMNS);
}
