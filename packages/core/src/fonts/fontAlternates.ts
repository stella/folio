import type { FontTable } from "../types/document";

export type FontAlternates = ReadonlyMap<string, string>;

const normalizeFontName = (name: string): string => name.trim().toLowerCase();

/** Build a document-scoped primary-name → OOXML alternate-name lookup. */
export const buildFontAlternates = (fontTable: FontTable | null | undefined): FontAlternates => {
  const alternates = new Map<string, string>();
  for (const font of fontTable?.fonts ?? []) {
    const primary = normalizeFontName(font.name);
    const alternate = font.altName?.trim();
    if (!primary || !alternate || normalizeFontName(alternate) === primary) {
      continue;
    }
    alternates.set(primary, alternate);
  }
  return alternates;
};

export const getFontAlternate = (
  family: string | undefined,
  alternates: FontAlternates | undefined,
): string | undefined =>
  family === undefined ? undefined : alternates?.get(normalizeFontName(family));
