import type { ParagraphAttrs } from "./types";

const LEGACY_JUSTIFICATION_MAX_COMPATIBILITY_MODE = 14;

export function resolveJustificationCompatibility(
  compatibilityMode: number | undefined,
): NonNullable<ParagraphAttrs["justificationCompatibility"]> | undefined {
  if (
    compatibilityMode === undefined ||
    compatibilityMode > LEGACY_JUSTIFICATION_MAX_COMPATIBILITY_MODE
  ) {
    return undefined;
  }
  return { type: "legacy" };
}
