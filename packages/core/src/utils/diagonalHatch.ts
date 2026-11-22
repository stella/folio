import type { DiagonalHatchPattern } from "../display-list/types";

// Paint rising diagonals inside the rectangular band.
export function diagonalHatchGradient(
  color: string,
  { pitchPx, lineWidthPx, slots }: DiagonalHatchPattern,
): string {
  const perpendicularPitchPx = pitchPx / Math.SQRT2;
  const stops = [];
  for (const [index, painted] of slots.entries()) {
    const startPx = index * perpendicularPitchPx;
    const endPx = startPx + perpendicularPitchPx;
    if (painted) {
      const inkEndPx = startPx + lineWidthPx;
      stops.push(`${color} ${startPx}px ${inkEndPx}px`);
      stops.push(`transparent ${inkEndPx}px ${endPx}px`);
    } else {
      stops.push(`transparent ${startPx}px ${endPx}px`);
    }
  }
  return `repeating-linear-gradient(135deg, ${stops.join(", ")})`;
}
