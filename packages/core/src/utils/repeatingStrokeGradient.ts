import type { StrokeDashCycle } from "../display-list/types";

type RepeatingStrokeGradientOptions = {
  color: string;
  direction: string;
  dashArrayPx: StrokeDashCycle;
};

// Paint one cycle with hard stops between ink and gaps.
export function repeatingStrokeGradient({
  color,
  direction,
  dashArrayPx,
}: RepeatingStrokeGradientOptions): string {
  const stops = [];
  let start = 0;
  for (const [index, length] of dashArrayPx.entries()) {
    const ink = index % 2 === 0 ? color : "transparent";
    const end = start + length;
    stops.push(`${ink} ${start}px ${end}px`);
    start = end;
  }
  return `repeating-linear-gradient(${direction}, ${stops.join(", ")})`;
}
