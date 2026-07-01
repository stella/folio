/**
 * Registers a document's embedded fonts with the browser via the `FontFace`
 * API, so runs render in their authored faces instead of the fallback stack.
 *
 * The extraction + de-obfuscation is done headlessly in `@stll/folio-core`;
 * this module only performs the DOM registration and cleanup. Faces are kept as
 * `FontFace` handles so the caller can remove them on document change/unmount,
 * preventing one document's embedded family from bleeding into the next.
 */

import { extractEmbeddedFonts, type EmbeddedFont } from "@stll/folio-core/fonts/embeddedFonts";

function getDocumentFontSet(): FontFaceSet | null {
  if (typeof document === "undefined" || !("fonts" in document)) {
    return null;
  }
  return document.fonts;
}

/**
 * Extract, register, and load the embedded font faces from a raw `.docx`
 * buffer. No-op (returns `[]`) outside a DOM or when the document has no
 * embedded fonts. Returns the `FontFace` handles that loaded successfully;
 * faces the browser rejects (malformed/subsetted binaries) are dropped so the
 * run falls back to the CSS font stack.
 */
export async function loadEmbeddedFontFaces(buffer: ArrayBuffer): Promise<FontFace[]> {
  const fontSet = getDocumentFontSet();
  if (!fontSet) {
    return [];
  }

  // Embedded fonts are best-effort: a corrupt/oversized package makes unzip
  // throw. Never let that break rendering — fall back to the CSS font stack.
  let fonts: EmbeddedFont[];
  try {
    fonts = await extractEmbeddedFonts(buffer);
  } catch {
    return [];
  }
  if (fonts.length === 0) {
    return [];
  }

  const loaded: FontFace[] = [];
  await Promise.all(
    fonts.map(async (font) => {
      let face: FontFace;
      try {
        // `new FontFace` throws synchronously (SyntaxError) on an invalid family
        // or descriptor; `fontSet.add` can reject a face too. Skip on failure.
        face = new FontFace(font.family, font.bytes, {
          weight: String(font.weight),
          style: font.style,
        });
        fontSet.add(face);
      } catch {
        return;
      }
      const ready = await face.load().then(
        () => true,
        () => false,
      );
      if (ready) {
        loaded.push(face);
        return;
      }
      fontSet.delete(face);
    }),
  );
  return loaded;
}

/** Remove previously registered embedded faces from the document font set. */
export function removeEmbeddedFontFaces(faces: readonly FontFace[]): void {
  const fontSet = getDocumentFontSet();
  if (!fontSet) {
    return;
  }
  for (const face of faces) {
    fontSet.delete(face);
  }
}
