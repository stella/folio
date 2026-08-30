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
    !Number.isInteger(horizontalScale) ||
    horizontalScale < 0 ||
    horizontalScale > 600
  ) {
    return undefined;
  }
  return horizontalScale;
}

/**
 * Parse the lexical forms accepted by OOXML `ST_TextScale`.
 *
 * Decimal values use the XML Schema integer grammar; percent values use the
 * unsigned decimal grammar defined by `ST_TextScalePercent`. Surrounding
 * whitespace is ignored, but an empty value or partial numeric prefix is not.
 */
export function parseHorizontalScalePercent(
  horizontalScale: string | null | undefined,
): number | undefined {
  if (horizontalScale === null || horizontalScale === undefined) {
    return undefined;
  }

  const lexicalValue = horizontalScale.trim();
  if (!/^(?:[+-]?\d+|\d+%)$/u.test(lexicalValue)) {
    return undefined;
  }

  const numericValue = lexicalValue.endsWith("%") ? lexicalValue.slice(0, -1) : lexicalValue;
  return normalizeHorizontalScalePercent(Number(numericValue));
}

export function getHorizontalScaleFactor(horizontalScale: number | null | undefined): number {
  return (normalizeHorizontalScalePercent(horizontalScale) ?? 100) / 100;
}
