import type { ColorValue, TextFormatting } from "../types/document";
import { mergeFontFamily } from "./fontFamilyMerge";

// Double strike is deliberately absent: it uses ordinary last-defined inheritance.
export const STYLE_TOGGLE_KEYS = [
  "bold",
  "boldCs",
  "italic",
  "italicCs",
  "allCaps",
  "emboss",
  "imprint",
  "outline",
  "shadow",
  "smallCaps",
  "strike",
  "hidden",
] as const satisfies readonly (keyof TextFormatting)[];

/**
 * Merge the properties exposed by one resolved style definition.
 *
 * A false toggle at a child style level does not cancel the inherited state. The
 * ordinary merge remains unchanged for direct formatting and non-toggle properties.
 */
export function mergeStyleTextFormatting(
  target: TextFormatting | undefined,
  source: TextFormatting | undefined,
): TextFormatting | undefined {
  const result = mergeTextFormatting(target, source);
  if (!result || !source) {
    return result;
  }

  for (const key of STYLE_TOGGLE_KEYS) {
    if (source[key] !== false) {
      continue;
    }
    const inherited = target?.[key];
    if (inherited === undefined) {
      Reflect.deleteProperty(result, key);
    } else {
      result[key] = inherited;
    }
  }
  return result;
}

export function mergeTextFormatting(
  target: TextFormatting | undefined,
  source: TextFormatting | undefined,
): TextFormatting | undefined {
  if (!source && !target) {
    return undefined;
  }
  if (!source) {
    return target;
  }
  if (!target) {
    return { ...source };
  }

  const result: Record<string, unknown> = { ...target };

  for (const key of Object.keys(source) as (keyof TextFormatting)[]) {
    const value = source[key];
    if (value === undefined) {
      continue;
    }

    if (key === "fontFamily" && typeof value === "object") {
      result["fontFamily"] = mergeFontFamily(
        target.fontFamily,
        value as NonNullable<TextFormatting["fontFamily"]>,
      );
      continue;
    }

    if (key === "color" && typeof value === "object") {
      result["color"] = value as ColorValue;
      continue;
    }

    if (typeof value === "object" && !Array.isArray(value)) {
      result[key] = {
        ...(target[key] as Record<string, unknown> | undefined),
        ...(value as Record<string, unknown>),
      };
      continue;
    }

    result[key] = value;
  }

  return result as TextFormatting;
}
