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

export function getDocumentFontSet(): FontFaceSet | null {
  if (typeof document === "undefined" || !("fonts" in document)) {
    return null;
  }
  return document.fonts;
}

/**
 * Register one font face on the document font set, best-effort. Returns the
 * loaded `FontFace`, or `null` when construction throws (invalid family or
 * descriptor) or the binary/URL fails to load (malformed / subsetted). Shared
 * by the embedded-document-font and host-provided-font (`fonts` prop) paths so
 * both skip a bad face identically instead of throwing.
 */
export async function registerFontFace(
  fontSet: FontFaceSet,
  family: string,
  source: string | BufferSource,
  descriptors: FontFaceDescriptors,
): Promise<FontFace | null> {
  let face: FontFace;
  try {
    // `new FontFace` throws synchronously (SyntaxError) on an invalid family
    // or descriptor; `fontSet.add` can reject a face too. Skip on failure.
    face = new FontFace(family, source, descriptors);
    fontSet.add(face);
  } catch {
    return null;
  }
  const ready = await face.load().then(
    () => true,
    () => false,
  );
  if (ready) {
    return face;
  }
  fontSet.delete(face);
  return null;
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
      const face = await registerFontFace(fontSet, font.family, font.bytes, {
        weight: String(font.weight),
        style: font.style,
      });
      if (face) {
        loaded.push(face);
      }
    }),
  );
  return loaded;
}

/** Remove previously registered faces from the document font set. */
export function removeFontFaces(faces: readonly FontFace[]): void {
  const fontSet = getDocumentFontSet();
  if (!fontSet) {
    return;
  }
  for (const face of faces) {
    fontSet.delete(face);
  }
}
