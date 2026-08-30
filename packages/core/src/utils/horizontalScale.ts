/**
 * Normalize OOXML `w:w` horizontal text scale percentages.
 *
 * The OOXML value is bounded from 0% through 600%. Missing or malformed
 * values use the document default (100%) and stay absent in the model.
 */
export function normalizeHorizontalScalePercent(
  horizontalScale: number | null | undefined,
): number | undefined {
  if (
    horizontalScale === null ||
    horizontalScale === undefined ||
    !Number.isFinite(horizontalScale) ||
    horizontalScale < 0 ||
    horizontalScale > 600
  ) {
    return undefined;
  }
  return horizontalScale;
}

export function getHorizontalScaleFactor(horizontalScale: number | null | undefined): number {
  return (normalizeHorizontalScalePercent(horizontalScale) ?? 100) / 100;
}
