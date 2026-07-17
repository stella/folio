/** Register fonts embedded in a document with the browser font set. */

import {
  extractEmbeddedFonts,
  buildEmbeddedFontFamilyMap,
  type EmbeddedFont,
} from "@stll/folio-core/fonts/embeddedFonts";

import { getDocumentFontSet, registerFontFace } from "./fontFaces";

/** Result of registering a document's embedded fonts with the browser. */
export type LoadedEmbeddedFonts = {
  /** The `FontFace` handles that loaded successfully, for later cleanup. */
  faces: FontFace[];
  /**
   * Original DOCX font name (`w:font w:name`) → the per-document scoped
   * family it was registered under (never the raw name — see
   * `@stll/folio-core/fonts/embeddedFonts`). Callers activate this with
   * `setEmbeddedFontFamilyMap` (`@stll/folio-core/utils/fontResolver`) so the
   * document's own runs keep resolving to their embedded face.
   */
  familyMap: ReadonlyMap<string, string>;
};

const EMPTY_LOADED_EMBEDDED_FONTS: LoadedEmbeddedFonts = { faces: [], familyMap: new Map() };

/**
 * Extract and register embedded faces from a raw document buffer. Extraction
 * and individual face failures are best-effort: rendering falls back to the
 * normal CSS font stack instead of failing the document load.
 */
export async function loadEmbeddedFontFaces(buffer: ArrayBuffer): Promise<LoadedEmbeddedFonts> {
  const fontSet = getDocumentFontSet();
  if (!fontSet) {
    return EMPTY_LOADED_EMBEDDED_FONTS;
  }

  let fonts: EmbeddedFont[];
  try {
    fonts = await extractEmbeddedFonts(buffer);
  } catch {
    return EMPTY_LOADED_EMBEDDED_FONTS;
  }

  const faces = await Promise.all(
    fonts.map((font) =>
      // `font.family` is already the per-document scoped name, never the raw
      // DOCX name (see `@stll/folio-core/fonts/embeddedFonts`).
      registerFontFace(fontSet, font.family, font.bytes, {
        weight: String(font.weight),
        style: font.style,
      }),
    ),
  );
  return {
    faces: faces.filter((face): face is FontFace => face !== null),
    familyMap: buildEmbeddedFontFamilyMap(fonts),
  };
}
