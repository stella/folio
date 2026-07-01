/**
 * Host-app custom fonts — the `fonts` prop on `DocxEditor`.
 *
 * The host passes its own brand/web font faces; they register with the browser
 * through the same best-effort `FontFace` path as embedded document fonts
 * ({@link registerFontFace}), so a bad entry is skipped rather than throwing.
 * Once registered, the family name is available for rendering and can be listed
 * in the font-family picker via the `fontFamilies` prop.
 *
 * The prop normalizer ({@link toFontFaceInputs}) is pure (no DOM), so it is
 * unit-tested directly; the DOM registration is visual-only.
 */

import { getDocumentFontSet, registerFontFace } from "./embeddedFonts";

/**
 * Declarative description of one custom font face the host registers with the
 * editor. Multiple entries can share `family` to register distinct weights.
 */
export type FontDefinition = {
  /** CSS `font-family` name to expose; match the family your documents reference. */
  family: string;
  /** URL to the font file (woff2, woff, ttf, or otf). */
  src: string;
  /** CSS `font-weight` for this face (a number like `700`, or a keyword). Defaults to `normal`. */
  weight?: number | string;
};

/** `FontFace` constructor inputs derived from a {@link FontDefinition}. */
export type FontFaceInput = {
  family: string;
  source: string;
  descriptors: FontFaceDescriptors;
};

/**
 * Normalize the `fonts` prop into `FontFace` constructor inputs. Pure: no DOM
 * work, so it is unit-testable. Entries missing a family or src are skipped
 * (best-effort); `src` becomes a CSS `url(...)` source string.
 */
export function toFontFaceInputs(
  fonts: ReadonlyArray<FontDefinition> | undefined,
): FontFaceInput[] {
  if (!fonts) {
    return [];
  }
  const inputs: FontFaceInput[] = [];
  for (const font of fonts) {
    const family = font.family.trim();
    const src = font.src.trim();
    if (!family || !src) {
      continue;
    }
    const descriptors: FontFaceDescriptors = { display: "swap" };
    if (font.weight !== undefined) {
      descriptors.weight = String(font.weight);
    }
    inputs.push({ family, source: `url(${JSON.stringify(src)})`, descriptors });
  }
  return inputs;
}

/**
 * Register the host's custom font faces with the browser via the `FontFace`
 * API. No-op (returns `[]`) outside a DOM or when no valid fonts are given.
 * Returns the faces that loaded successfully; each bad entry is skipped, never
 * thrown. Pair with {@link removeFontFaces} to unregister on change/unmount.
 */
export async function loadHostFontFaces(
  fonts: ReadonlyArray<FontDefinition> | undefined,
): Promise<FontFace[]> {
  const fontSet = getDocumentFontSet();
  if (!fontSet) {
    return [];
  }
  const inputs = toFontFaceInputs(fonts);
  if (inputs.length === 0) {
    return [];
  }
  const loaded: FontFace[] = [];
  await Promise.all(
    inputs.map(async ({ family, source, descriptors }) => {
      const face = await registerFontFace(fontSet, family, source, descriptors);
      if (face) {
        loaded.push(face);
      }
    }),
  );
  return loaded;
}
