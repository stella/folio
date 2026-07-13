/**
 * Section-properties → column-layout helper.
 *
 * Sibling to `paged-layout/sectionGeometry.ts` (which owns page size + margin
 * translation). Column geometry lives here so any rendering adapter (React,
 * Vue, etc.) can share the same `SectionProperties` → `ColumnLayout` mapping
 * without re-deriving the twip math. Reuses `twipsToPixels` from the geometry
 * helper so the DPI conversion stays in one place.
 */

import type { ColumnLayout } from "../layout-engine/types";
import type { SectionProperties } from "../types/document";
import { twipsToPixels } from "../paged-layout/sectionGeometry";

/** Default column spacing per OOXML spec: 720 twips (0.5 inch). */
const DEFAULT_COLUMN_SPACE_TWIPS = 720;

/**
 * Extract column layout from section properties.
 * Returns undefined for single-column (default) to avoid unnecessary paginator overhead.
 */
export function getColumns(
  sectionProps: SectionProperties | null | undefined,
): ColumnLayout | undefined {
  const authoredColumns = sectionProps?.columns;
  const count = sectionProps?.columnCount ?? authoredColumns?.length ?? 1;
  if (count <= 1) return undefined;
  const gap = twipsToPixels(sectionProps?.columnSpace ?? DEFAULT_COLUMN_SPACE_TWIPS);
  const columns: ColumnLayout = {
    count,
    gap,
    equalWidth: sectionProps?.equalWidth ?? true,
  };
  if (
    sectionProps?.equalWidth === false &&
    authoredColumns?.length === count &&
    authoredColumns.every(({ width }) => width !== undefined)
  ) {
    columns.widths = authoredColumns.map(({ width }) => twipsToPixels(width ?? 0));
    columns.gaps = authoredColumns
      .slice(0, -1)
      .map(({ space }) =>
        twipsToPixels(space ?? sectionProps.columnSpace ?? DEFAULT_COLUMN_SPACE_TWIPS),
      );
  }
  // exactOptionalPropertyTypes: only attach `separator` when explicitly set.
  if (sectionProps?.separator !== undefined) columns.separator = sectionProps.separator;
  return columns;
}
