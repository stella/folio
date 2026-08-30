import { collectSectionConfigs } from "../layout-engine";
import type { SectionLayoutConfig } from "../layout-engine";
import { calculateColumnLefts, calculateColumnWidths } from "../layout-engine/paginator";
import type { ColumnLayout, FlowBlock } from "../layout-engine/types";

type ComputePerBlockMeasureInput = {
  blocks: FlowBlock[];
  bodyConfig: SectionLayoutConfig;
  finalConfig: SectionLayoutConfig;
};

type PerBlockMeasureInputs = {
  widths: number[];
  marginTops: number[];
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
  const pageHeights: number[] = [];
  const pageWidths: number[] = [];
  const marginLefts: number[] = [];
  const marginRights: number[] = [];
  const marginBottoms: number[] = [];
  const contentLefts: number[] = [];

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
    marginTops.push(config.margins.top);
    pageHeights.push(config.pageSize.h);
    pageWidths.push(config.pageSize.w);
    marginLefts.push(config.margins.left);
    marginRights.push(config.margins.right);
    marginBottoms.push(config.margins.bottom);

    if (sectionIdx < breakIndices.length && i === breakIndices[sectionIdx]) {
      sectionIdx++;
      columnIndex = 0;
    } else if (blocks[i]?.kind === "pageBreak") {
      columnIndex = 0;
    } else if (blocks[i]?.kind === "columnBreak") {
      columnIndex = (columnIndex + 1) % activeColumns.count;
    }
  }

  return {
    widths,
    marginTops,
    pageHeights,
    pageWidths,
    marginLefts,
    marginRights,
    marginBottoms,
    contentLefts,
  };
}
