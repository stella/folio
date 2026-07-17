/**
 * Host-app custom fonts — the `fonts` prop on the Vue `DocxEditor`.
 *
 * Vue port of `packages/react/src/paged-editor/hostFonts.ts`. The host passes
 * its own brand/web font faces; each registers with the browser through a
 * best-effort `FontFace` path, so a bad entry is skipped rather than throwing.
 * Once registered, the family name is available for rendering and can be
 * listed in the font-family picker via the `fontFamilies` prop.
 *
 * Framework-neutral: DOM `FontFace` APIs only, no Vue or React. The prop
 * normalizer ({@link toFontFaceInputs}) is pure (no DOM), so it is unit-testable
 * directly; the DOM registration is visual-only.
 */

import type { FontDefinition } from "../components/DocxEditor/types";
import { getDocumentFontSet, registerFontFace } from "./fontFaces";

/** `FontFace` constructor inputs derived from a {@link FontDefinition}. */
export type FontFaceInput = {
  family: string;
  source: string;
  descriptors: FontFaceDescriptors;
};

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

/**
 * Structural guard for one `fonts` entry. The prop is typed, but a plain-JS host
 * can pass anything (null, a bare value, wrong field types), so validate at the
 * boundary instead of trusting the declared type.
 */
function isValidFontDefinition(font: unknown): font is FontDefinition {
  if (typeof font !== "object" || font === null) {
    return false;
  }
  return (
    isNonEmptyString(Reflect.get(font, "family")) && isNonEmptyString(Reflect.get(font, "src"))
  );
}

/**
 * Normalize the `fonts` prop into `FontFace` constructor inputs. Pure: no DOM
 * work, so it is unit-testable. Entries that are not objects with non-empty
 * string family/src are skipped (best-effort — never throw on host input); `src`
 * becomes a CSS `url(...)` source string.
 */
export function toFontFaceInputs(
  fonts: ReadonlyArray<FontDefinition> | undefined,
): FontFaceInput[] {
  if (!fonts) {
    return [];
  }
  const inputs: FontFaceInput[] = [];
  for (const font of fonts) {
    if (!isValidFontDefinition(font)) {
      continue;
    }
    const descriptors: FontFaceDescriptors = { display: "swap" };
    // Only a number/string weight is representable; anything else is dropped to
    // the default.
    if (typeof font.weight === "number" || typeof font.weight === "string") {
      descriptors.weight = String(font.weight);
    }
    inputs.push({
      family: font.family.trim(),
      source: `url(${JSON.stringify(font.src.trim())})`,
      descriptors,
    });
  }
  return inputs;
}

/**
 * Register the host's custom font faces with the browser via the `FontFace` API.
 * No-op (returns `[]`) outside a DOM or when no valid fonts are given. Returns
 * the faces that loaded successfully; each bad entry is skipped, never thrown.
 * The font lifecycle unregisters returned faces on change/unmount.
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
