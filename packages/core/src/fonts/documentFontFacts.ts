import type { FontInfo, FontTable } from "@stll/docx-core/model";

import { canonicalJson } from "../utils/canonicalJson";

type DocumentFontInfo = Readonly<
  Pick<FontInfo, "family" | "pitch"> & {
    charset?: Readonly<NonNullable<FontInfo["charset"]>>;
  }
>;

export type DocumentFontFacts = Readonly<{
  fingerprint: string;
  fontInfoOf(name: string): DocumentFontInfo | undefined;
}>;

const normalizeFontName = (name: string): string => name.trim().toLowerCase();

class OwnedDocumentFontFacts implements DocumentFontFacts {
  readonly #fontsByName: ReadonlyMap<string, DocumentFontInfo>;
  readonly fingerprint: string;

  constructor(fontsByName: Map<string, DocumentFontInfo>) {
    this.#fontsByName = fontsByName;
    this.fingerprint = fingerprintEntries(fontsByName);
    Object.freeze(this);
  }

  fontInfoOf(name: string): DocumentFontInfo | undefined {
    return this.#fontsByName.get(normalizeFontName(name));
  }
}

export const buildDocumentFontFacts = (
  fontTable: FontTable | null | undefined,
): DocumentFontFacts | undefined => {
  if (!fontTable?.fonts.length) {
    return undefined;
  }
  const fontsByName = new Map<string, DocumentFontInfo>();
  for (const font of fontTable.fonts) {
    for (const name of [font.name, font.altName ?? ""]) {
      const normalized = normalizeFontName(name);
      if (normalized && !fontsByName.has(normalized)) {
        fontsByName.set(
          normalized,
          Object.freeze({
            ...(font.family === undefined ? {} : { family: font.family }),
            ...(font.pitch === undefined ? {} : { pitch: font.pitch }),
            ...(font.charset === undefined
              ? {}
              : { charset: Object.freeze({ ...font.charset }) }),
          }),
        );
      }
    }
  }
  return new OwnedDocumentFontFacts(fontsByName);
};

const fingerprintEntries = (
  facts: ReadonlyMap<string, DocumentFontInfo>,
): string => {
  const entries = [...facts].sort(([left], [right]) => {
    if (left < right) return -1;
    if (left > right) return 1;
    return 0;
  });
  return canonicalJson(entries);
};

export const documentFontFactsFingerprint = (
  facts: DocumentFontFacts | undefined,
): string | undefined => facts?.fingerprint;

export const documentFontInfoOf = (
  fontName: string,
  facts: DocumentFontFacts | undefined,
): DocumentFontInfo | undefined => facts?.fontInfoOf(fontName);
