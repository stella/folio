/**
 * Normalizer for the `fontFamilies` prop (the font-family picker list). Pure,
 * so it is unit-testable; shared by DocxEditor's toolbar wiring.
 */

import type { FontOption } from "./FontPicker";

/**
 * Normalize the `fontFamilies` prop (a mix of plain family-name strings and
 * {@link FontOption} objects) into a uniform `FontOption[]`. Returns
 * `undefined` for `undefined` input so the picker falls back to its built-in
 * defaults; an empty array stays empty (an empty, enabled dropdown). Strings
 * expand into the `"other"` category with the name used as the CSS family.
 */
export function normalizeFontFamilies(
  fontFamilies: ReadonlyArray<string | FontOption> | undefined,
): FontOption[] | undefined {
  if (fontFamilies === undefined) {
    return undefined;
  }
  return fontFamilies.map(
    (font): FontOption =>
      typeof font === "string" ? { name: font, fontFamily: font, category: "other" } : font,
  );
}
