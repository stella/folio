/**
 * Shared text-measurement policy.
 *
 * Keep authored typography decisions here so layout measurement, worker
 * requests, and painting cannot interpret the same run differently.
 */

import type { RunFormatting } from "../types";
import type { FontStyle } from "./measureTypes";

export const FONT_KERNING_MODE = {
  enabled: "normal",
  disabled: "none",
} as const;

export type FontKerningMode = (typeof FONT_KERNING_MODE)[keyof typeof FONT_KERNING_MODE];

type RunKerningInput = Pick<RunFormatting, "fontSize" | "kerningMinPt">;

/** Resolve an authored kerning threshold against the effective run size. */
export function getRunFontKerningMode(
  run: RunKerningInput,
  fallbackFontSize: number,
): FontKerningMode {
  const threshold = run.kerningMinPt;
  const fontSize = run.fontSize ?? fallbackFontSize;
  if (threshold === undefined || fontSize < threshold) {
    return FONT_KERNING_MODE.disabled;
  }
  return FONT_KERNING_MODE.enabled;
}

/** Convert normalized measurement style into the canvas/CSS kerning mode. */
export function getFontKerningMode(style: Pick<FontStyle, "kerning">): FontKerningMode {
  return style.kerning ? FONT_KERNING_MODE.enabled : FONT_KERNING_MODE.disabled;
}

/** Count ordinary ASCII spaces that justification may compress. */
export function countCompressibleSpaces(text: string): number {
  let count = 0;
  for (const char of text) {
    if (char === " ") {
      count += 1;
    }
  }
  return count;
}
