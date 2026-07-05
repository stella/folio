/**
 * Shared FontOption shape + normaliser used by FontPicker components.
 * @packageDocumentation
 * @public
 */

export type FontOption = {
  name: string;
  fontFamily: string;
  category?: "sans-serif" | "serif" | "monospace" | "other";
}

/**
 * Normalize a `fontFamilies` prop (mix of strings and FontOption
 * objects) into a uniform `FontOption[]`. Returns `undefined` for
 * `undefined` input so callers fall back to their built-in defaults.
 * Strings expand into the `'other'` group with no CSS fallback chain.
 */
export function normalizeFontFamilies(
  fontFamilies: ReadonlyArray<string | FontOption> | undefined,
): FontOption[] | undefined {
  if (fontFamilies === undefined) return undefined;
  const normalized = fontFamilies.map(
    (f): FontOption => (typeof f === "string" ? { name: f, fontFamily: f, category: "other" } : f),
  );
  if (isDev()) {
    const warned = new Set<string>();
    const seen = new Set<string>();
    for (const f of normalized) {
      if (seen.has(f.name) && !warned.has(f.name)) {
        console.warn(`[DocxEditor] Duplicate font name in fontFamilies: "${f.name}"`);
        warned.add(f.name);
      }
      seen.add(f.name);
    }
  }
  return normalized;
}

function isDev(): boolean {
  return typeof process !== "undefined" && process.env?.["NODE_ENV"] !== "production";
}

/**
 * Drop fonts whose names already appear in `existingNames` (case-insensitive),
 * also deduping the input. Used by the toolbar's font picker to render the
 * "Document fonts" group without repeating a font the built-in list covers.
 *
 * Ported from the upstream core util `utils/documentPickerFonts.ts`; folio-core
 * has no equivalent, so it lives with the Vue adapter's font helpers.
 */
export function excludeFontsByName(
  fonts: readonly FontOption[] | undefined,
  existingNames: Iterable<string>,
): FontOption[] {
  if (!fonts || fonts.length === 0) return [];
  const existing = new Set<string>();
  for (const name of existingNames) existing.add(name.trim().toLowerCase());
  const seen = new Set<string>();
  const out: FontOption[] = [];
  for (const f of fonts) {
    const key = f.name.trim().toLowerCase();
    if (existing.has(key) || seen.has(key)) continue;
    seen.add(key);
    out.push(f);
  }
  return out;
}
