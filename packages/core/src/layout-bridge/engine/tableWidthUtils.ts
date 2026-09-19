export function normalizeTableColumnWidths(
  widths: number[],
  columnCount: number,
  targetWidth: number,
): number[] {
  if (columnCount <= 0) {
    return [];
  }
  if (widths.length === 0) {
    return Array.from({ length: columnCount }, () => targetWidth / columnCount);
  }

  const normalized = widths.slice(0, columnCount);
  const positives = normalized.filter((width) => width > 0);
  const fallback =
    positives.length > 0
      ? positives.reduce((sum, width) => sum + width, 0) / positives.length
      : targetWidth / columnCount;

  while (normalized.length < columnCount) {
    normalized.push(fallback);
  }

  let fixedTotal = 0;
  for (const width of normalized) {
    if (width > 0) {
      fixedTotal += width;
    }
  }
  const missingCount = normalized.filter((width) => width <= 0).length;
  const missingWidth = missingCount > 0 ? Math.max(0, targetWidth - fixedTotal) / missingCount : 0;

  return normalized.map((width) => (width > 0 ? width : missingWidth));
}
