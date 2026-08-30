import { collectSectionConfigs } from "../layout-engine";
import type { SectionLayoutConfig } from "../layout-engine";
import { hasPageBreakBefore } from "../layout-engine/keep-together";
import { calculateColumnLefts, calculateColumnWidths } from "../layout-engine/paginator";
import { normalizeSectionBreakType } from "../layout-engine/section-breaks";
import type { ColumnLayout, FlowBlock } from "../layout-engine/types";

type ComputePerBlockMeasureInput = {
  blocks: FlowBlock[];
  bodyConfig: SectionLayoutConfig;
  finalConfig: SectionLayoutConfig;
};

type PerBlockPhysicalPageGeometry = {
  pageHeights: number[];
  pageWidths: number[];
  marginTops: number[];
  marginLefts: number[];
  marginRights: number[];
  marginBottoms: number[];
};

type PerBlockMeasureInputs = {
  widths: number[];
  /** Authored section top, retained for section-relative text-box semantics. */
  marginTops: number[];
  pageHeights: number[];
  pageWidths: number[];
  marginLefts: number[];
  marginRights: number[];
  marginBottoms: number[];
  /** Absolute page X of the active column's content origin. */
  contentLefts: number[];
  columnIndices: number[];
  columnCounts: number[];
  physicalPageGeometry: PerBlockPhysicalPageGeometry;
};

type PhysicalPageGeometry = Pick<SectionLayoutConfig, "pageSize" | "margins">;

type PreBlockPageTransition =
  | { type: "none" }
  | { type: "physicalPage"; geometry: PhysicalPageGeometry };

function resolvePreBlockPageTransition(
  block: FlowBlock,
  config: SectionLayoutConfig,
): PreBlockPageTransition {
  if (!hasPageBreakBefore(block)) {
    return { type: "none" };
  }
  return {
    type: "physicalPage",
    geometry: { pageSize: config.pageSize, margins: config.margins },
  };
}

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
  let columnGeometryDirty = true;
  const widths: number[] = [];
  const marginTops: number[] = [];
  const physicalMarginTops: number[] = [];
  const pageHeights: number[] = [];
  const pageWidths: number[] = [];
  const marginLefts: number[] = [];
  const marginRights: number[] = [];
  const marginBottoms: number[] = [];
  const physicalPageHeights: number[] = [];
  const physicalPageWidths: number[] = [];
  const physicalMarginLefts: number[] = [];
  const physicalMarginRights: number[] = [];
  const physicalMarginBottoms: number[] = [];
  const contentLefts: number[] = [];
  const columnIndices: number[] = [];
  const columnCounts: number[] = [];
  const initialConfig = sectionConfigs[0] ?? finalConfig;
  let physicalPageGeometry: PhysicalPageGeometry = {
    pageSize: initialConfig.pageSize,
    margins: initialConfig.margins,
  };
  let hasPhysicalPageContent = false;

  for (let i = 0; i < blocks.length; i++) {
    const config = sectionConfigs[sectionIdx] ?? finalConfig;
    const block = blocks[i]!; // SAFETY: i < blocks.length
    const preBlockTransition = resolvePreBlockPageTransition(block, config);
    switch (preBlockTransition.type) {
      case "physicalPage":
        physicalPageGeometry = preBlockTransition.geometry;
        hasPhysicalPageContent = false;
        columnGeometryDirty = true;
        columnIndex = 0;
        break;
      case "none":
        break;
      default:
        preBlockTransition satisfies never;
    }
    if (measuredSectionIdx !== sectionIdx || columnGeometryDirty) {
      activeColumns = config.columns ?? SINGLE_COLUMN_LAYOUT;
      activeColumnWidths = calculateColumnWidths(
        physicalPageGeometry.pageSize.w,
        physicalPageGeometry.margins.left,
        physicalPageGeometry.margins.right,
        activeColumns,
      );
      activeContentLefts = calculateColumnLefts(
        physicalPageGeometry.margins.left,
        activeColumnWidths,
        activeColumns,
      );
      measuredSectionIdx = sectionIdx;
      columnGeometryDirty = false;
    }
    widths.push(activeColumnWidths[columnIndex] ?? activeColumnWidths[0] ?? 0);
    contentLefts.push(activeContentLefts[columnIndex] ?? physicalPageGeometry.margins.left);
    columnIndices.push(columnIndex);
    columnCounts.push(activeColumns.count);
    marginTops.push(config.margins.top);
    physicalMarginTops.push(physicalPageGeometry.margins.top);
    pageHeights.push(config.pageSize.h);
    pageWidths.push(config.pageSize.w);
    marginLefts.push(config.margins.left);
    marginRights.push(config.margins.right);
    marginBottoms.push(config.margins.bottom);
    physicalPageHeights.push(physicalPageGeometry.pageSize.h);
    physicalPageWidths.push(physicalPageGeometry.pageSize.w);
    physicalMarginLefts.push(physicalPageGeometry.margins.left);
    physicalMarginRights.push(physicalPageGeometry.margins.right);
    physicalMarginBottoms.push(physicalPageGeometry.margins.bottom);

    if (sectionIdx < breakIndices.length && i === breakIndices[sectionIdx]) {
      const sectionBreak = block;
      const nextConfig = sectionConfigs[sectionIdx + 1] ?? finalConfig;
      const sharesPhysicalPage =
        sectionBreak?.kind === "sectionBreak" &&
        normalizeSectionBreakType(sectionBreak.type) === "continuous" &&
        hasPhysicalPageContent &&
        Math.round(nextConfig.pageSize.w) === Math.round(physicalPageGeometry.pageSize.w) &&
        Math.round(nextConfig.pageSize.h) === Math.round(physicalPageGeometry.pageSize.h);
      if (!sharesPhysicalPage) {
        physicalPageGeometry = {
          pageSize: nextConfig.pageSize,
          margins: nextConfig.margins,
        };
        hasPhysicalPageContent = false;
      }
      columnGeometryDirty = true;
      sectionIdx++;
      columnIndex = 0;
    } else if (block.kind === "pageBreak") {
      physicalPageGeometry = { pageSize: config.pageSize, margins: config.margins };
      hasPhysicalPageContent = false;
      columnGeometryDirty = true;
      columnIndex = 0;
    } else if (block.kind === "columnBreak") {
      if (columnIndex + 1 >= activeColumns.count) {
        physicalPageGeometry = { pageSize: config.pageSize, margins: config.margins };
        hasPhysicalPageContent = false;
        columnGeometryDirty = true;
      }
      columnIndex = (columnIndex + 1) % activeColumns.count;
    } else {
      hasPhysicalPageContent = true;
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
    columnIndices,
    columnCounts,
    physicalPageGeometry: {
      pageHeights: physicalPageHeights,
      pageWidths: physicalPageWidths,
      marginTops: physicalMarginTops,
      marginLefts: physicalMarginLefts,
      marginRights: physicalMarginRights,
      marginBottoms: physicalMarginBottoms,
    },
  };
}
