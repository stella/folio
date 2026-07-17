/** Browser-only `FontFace` registration shared by Vue's font sources. */

/** The document's `FontFaceSet`, or `null` outside a DOM. */
export function getDocumentFontSet(): FontFaceSet | null {
  if (typeof document === "undefined" || !("fonts" in document)) {
    return null;
  }
  return document.fonts;
}

/**
 * Register one font face best-effort. Invalid descriptors and rejected font
 * binaries are skipped so one bad face cannot prevent the document rendering.
 */
export async function registerFontFace(
  fontSet: FontFaceSet,
  family: string,
  source: string | BufferSource,
  descriptors: FontFaceDescriptors,
): Promise<FontFace | null> {
  let face: FontFace;
  try {
    face = new FontFace(family, source, descriptors);
    fontSet.add(face);
  } catch {
    return null;
  }

  const loaded = await face.load().then(
    () => true,
    () => false,
  );
  if (loaded) {
    return face;
  }

  fontSet.delete(face);
  return null;
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
