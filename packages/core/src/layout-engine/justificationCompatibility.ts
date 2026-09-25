import type { ParagraphAttrs } from "./types";

/**
 * First `compatibilityMode` that fits justified lines with the current rules.
 *
 * Earlier modes keep strict justified fitting. A document that declares no
 * `compatSetting` at all is read with the oldest semantics, so it takes the
 * legacy rule too.
 */
const FIRST_MODERN_JUSTIFICATION_COMPATIBILITY_MODE = 15;

export const resolveJustificationCompatibility = (
  compatibilityMode: number | undefined,
): NonNullable<ParagraphAttrs["justificationCompatibility"]> | undefined =>
  compatibilityMode !== undefined &&
  compatibilityMode >= FIRST_MODERN_JUSTIFICATION_COMPATIBILITY_MODE
    ? undefined
    : { type: "legacy" };
