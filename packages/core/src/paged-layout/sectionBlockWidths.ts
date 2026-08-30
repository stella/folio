import { collectSectionConfigs } from "../layout-engine";
import type { SectionLayoutConfig } from "../layout-engine";
import { calculateColumnLefts, calculateColumnWidths } from "../layout-engine/paginator";
import { normalizeSectionBreakType } from "../layout-engine/section-breaks";
import type { ColumnLayout, FlowBlock } from "../layout-engine/types";

type ComputePerBlockMeasureInput = {
  blocks: FlowBlock[];
  bodyConfig: SectionLayoutConfig;
  finalConfig: SectionLayoutConfig;
};

type PerBlockMeasureInputs = {
  widths: number[];
  marginTops: number[];
  /** Physical page content origin, which can differ after a continuous section. */
  contentTops: number[];
  // Page geometry per block, used to resolve page/margin-pinned topAndBottom
  // bands (`bandTopContentY`) in the measure pass. Sections can vary page size
  // and margins, so these are per-block like `widths`/`marginTops`.
  pageHeights: number[];
  pageWidths: number[];
  marginLefts: number[];
  marginRights: number[];
  marginBottoms: number[];
  /** Absolute page X of the active column's content origin. */
  contentLefts: number[];
  columnIndices: number[];
  columnCounts: number[];
};

const SINGLE_COLUMN_LAYOUT: ColumnLayout = { count: 1, gap: 0 };

/**
 * Compute per-block measurement widths by scanning for section breaks.
 * Blocks in multi-column sections must be measured at column width, not full content width.
 *
 * OOXML note: Each section break carries the CURRENT section's properties.
 * Section N's blocks use config from sectionBreak[N].
 * The final section (after all breaks) uses body-level config.
 */
export function computePerBlockWidths({
  blocks,
  bodyConfig,
  finalConfig,
}: ComputePerBlockMeasureInput): number[] {
  return computePerBlockMeasureInputs({
    blocks,
    bodyConfig,
    finalConfig,
  }).widths;
}

export function computePerBlockMeasureInputs({
  blocks,
  bodyConfig,
  finalConfig,
}: ComputePerBlockMeasureInput): PerBlockMeasureInputs {
  const { configs: sectionConfigs, breakIndices } = collectSectionConfigs(
    blocks,
    bodyConfig,
    finalConfig,
  );

  let sectionIdx = 0;
  let columnIndex = 0;
  let measuredSectionIdx = -1;
  let activeColumns = SINGLE_COLUMN_LAYOUT;
  let activeColumnWidths: number[] = [];
  let activeContentLefts: number[] = [];
  const widths: number[] = [];
  const marginTops: number[] = [];
  const contentTops: number[] = [];
  const pageHeights: number[] = [];
  const pageWidths: number[] = [];
  const marginLefts: number[] = [];
  const marginRights: number[] = [];
  const marginBottoms: number[] = [];
  const contentLefts: number[] = [];
  const columnIndices: number[] = [];
  const columnCounts: number[] = [];
  const initialConfig = sectionConfigs[0] ?? finalConfig;
  let physicalContentTop = initialConfig.margins.top;
  let physicalPageSize = initialConfig.pageSize;
  let hasPhysicalPageContent = false;

  for (let i = 0; i < blocks.length; i++) {
    const config = sectionConfigs[sectionIdx] ?? finalConfig;
    if (measuredSectionIdx !== sectionIdx) {
      activeColumns = config.columns ?? SINGLE_COLUMN_LAYOUT;
      activeColumnWidths = calculateColumnWidths(
        config.pageSize.w,
        config.margins.left,
        config.margins.right,
        activeColumns,
      );
      activeContentLefts = calculateColumnLefts(
        config.margins.left,
        activeColumnWidths,
        activeColumns,
      );
      measuredSectionIdx = sectionIdx;
    }
    widths.push(activeColumnWidths[columnIndex] ?? activeColumnWidths[0] ?? 0);
    contentLefts.push(activeContentLefts[columnIndex] ?? config.margins.left);
    columnIndices.push(columnIndex);
    columnCounts.push(activeColumns.count);
    marginTops.push(config.margins.top);
    contentTops.push(physicalContentTop);
    pageHeights.push(config.pageSize.h);
    pageWidths.push(config.pageSize.w);
    marginLefts.push(config.margins.left);
    marginRights.push(config.margins.right);
    marginBottoms.push(config.margins.bottom);

    if (sectionIdx < breakIndices.length && i === breakIndices[sectionIdx]) {
      const sectionBreak = blocks[i];
      const nextConfig = sectionConfigs[sectionIdx + 1] ?? finalConfig;
      const sharesPhysicalPage =
        sectionBreak?.kind === "sectionBreak" &&
        normalizeSectionBreakType(sectionBreak.type) === "continuous" &&
        hasPhysicalPageContent &&
        Math.round(nextConfig.pageSize.w) === Math.round(physicalPageSize.w) &&
        Math.round(nextConfig.pageSize.h) === Math.round(physicalPageSize.h);
      if (!sharesPhysicalPage) {
        physicalContentTop = nextConfig.margins.top;
        physicalPageSize = nextConfig.pageSize;
        hasPhysicalPageContent = false;
      }
      sectionIdx++;
      columnIndex = 0;
    } else if (blocks[i]?.kind === "pageBreak") {
      physicalContentTop = config.margins.top;
      physicalPageSize = config.pageSize;
      hasPhysicalPageContent = false;
      columnIndex = 0;
    } else if (blocks[i]?.kind === "columnBreak") {
      if (columnIndex + 1 >= activeColumns.count) {
        physicalContentTop = config.margins.top;
        physicalPageSize = config.pageSize;
        hasPhysicalPageContent = false;
      }
      columnIndex = (columnIndex + 1) % activeColumns.count;
    } else {
      hasPhysicalPageContent = true;
    }
  }

  return {
    widths,
    marginTops,
    contentTops,
    pageHeights,
    pageWidths,
    marginLefts,
    marginRights,
    marginBottoms,
    contentLefts,
    columnIndices,
    columnCounts,
  };
}
