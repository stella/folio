/** Register fonts embedded in a document with the browser font set. */

import { extractEmbeddedFonts, type EmbeddedFont } from "@stll/folio-core/fonts/embeddedFonts";

import { getDocumentFontSet, registerFontFace } from "./fontFaces";

/**
 * Extract and register embedded faces from a raw document buffer. Extraction
 * and individual face failures are best-effort: rendering falls back to the
 * normal CSS font stack instead of failing the document load.
 */
export async function loadEmbeddedFontFaces(buffer: ArrayBuffer): Promise<FontFace[]> {
  const fontSet = getDocumentFontSet();
  if (!fontSet) {
    return [];
  }

  let fonts: EmbeddedFont[];
  try {
    fonts = await extractEmbeddedFonts(buffer);
  } catch {
    return [];
  }

  const faces = await Promise.all(
    fonts.map((font) =>
      registerFontFace(fontSet, font.family, font.bytes, {
        weight: String(font.weight),
        style: font.style,
      }),
    ),
  );
  return faces.filter((face): face is FontFace => face !== null);
}
