/**
 * Normalizer for the `fontFamilies` prop (the font-family picker list). Pure,
 * so it is unit-testable; shared by DocxEditor's toolbar wiring.
 */

import type { FontOption } from "./FontPicker";

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * Structural guard for one `fontFamilies` entry. The prop is typed, but a
 * plain-JS host can slip in null / a bare value / an object missing its name
 * or family, so validate at the boundary rather than trusting the type.
 */
function isFontOption(value: unknown): value is FontOption {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    isNonEmptyString(Reflect.get(value, "name")) &&
    isNonEmptyString(Reflect.get(value, "fontFamily"))
  );
}

/**
 * Normalize the `fontFamilies` prop (a mix of plain family-name strings and
 * {@link FontOption} objects) into a uniform `FontOption[]`. Returns
 * `undefined` for `undefined` input so the picker falls back to its built-in
 * defaults; an empty array stays empty (an empty, enabled dropdown). Strings
 * expand into the `"other"` category with the name used as the CSS family.
 * Null / undefined / malformed entries are skipped, never crashing the picker.
 */
export function normalizeFontFamilies(
  fontFamilies: ReadonlyArray<string | FontOption> | undefined,
): FontOption[] | undefined {
  if (fontFamilies === undefined) {
    return undefined;
  }
  const normalized: FontOption[] = [];
  for (const font of fontFamilies) {
    if (isNonEmptyString(font)) {
      normalized.push({ name: font, fontFamily: font, category: "other" });
      continue;
    }
    if (isFontOption(font)) {
      normalized.push(font);
    }
  }
  return normalized;
}
