/**
 * The display list's font table: faces interned in first-use order.
 *
 * Order matters beyond tidiness. `DisplayFontRef` is an index, so two builds of
 * the same layout must intern in the same sequence or the two lists compare
 * unequal for reasons that have nothing to do with what is painted. First use
 * in page order is the only ordering the producer can derive without a clock.
 */

import type { EmbeddedFont } from "../../fonts/embeddedFonts";
import { ptToPx } from "../../layout-engine/measure/measureHelpers";
import { getFontMetrics } from "../../layout-engine/measure/measureProvider";
import type { FontStyle } from "../../layout-engine/measure/measureTypes";
import { resolveFontFamily } from "../../utils/fontResolver";
import type { DisplayEmbeddedFont, DisplayFontFace, DisplayFontRef } from "../types";

const CSS_NORMAL_WEIGHT = 400;
const CSS_BOLD_WEIGHT = 700;

type GenericFamily = DisplayFontFace["generic"];

const GENERIC_FAMILIES = {
  serif: "serif",
  "sans-serif": "sans-serif",
  monospace: "monospace",
  cursive: "cursive",
  fantasy: "fantasy",
} as const satisfies Record<GenericFamily, GenericFamily>;

const DEFAULT_GENERIC: GenericFamily = "sans-serif";

const splitFallbackStack = (cssFallback: string): string[] =>
  cssFallback.split(",").map((entry) => entry.trim().replace(/^["']|["']$/gu, ""));

/**
 * The category `resolveFontFamily` chose. It is the last CSS generic in the
 * stack it built, which is where `detectFontCategory` puts it; reading it back
 * is cheaper and less brittle than re-deriving the category from the name.
 */
const genericOf = (stack: readonly string[]): GenericFamily => {
  for (let index = stack.length - 1; index >= 0; index -= 1) {
    // SAFETY: index is walked down from length - 1.
    const generic = GENERIC_FAMILIES[stack[index]!.toLowerCase() as GenericFamily];
    if (generic !== undefined) {
      return generic;
    }
  }
  return DEFAULT_GENERIC;
};

const faceKey = (family: string, weight: number, italic: boolean): string =>
  `${family.toLowerCase()} ${weight} ${italic ? "i" : "r"}`;

/**
 * Package faces keyed by every family name a run can resolve to.
 *
 * `resolveFontFamily` substitutes the scoped `folio-embedded-{nonce}-{name}`
 * family only while the document's embedded-family map is installed; a build
 * that never installed it resolves the same run to the authored name. Both
 * names therefore key the same face, and the bytes are interned by id so one
 * face's binary appears once however many names reach it.
 */
const indexEmbeddedFonts = (
  faces: readonly EmbeddedFont[],
): ReadonlyMap<string, DisplayEmbeddedFont> => {
  const byId = new Map<string, DisplayEmbeddedFont>();
  const byKey = new Map<string, DisplayEmbeddedFont>();
  for (const face of faces) {
    const id = `${face.family}:${face.weight}:${face.style}`;
    const embedded = byId.get(id) ?? { id, bytes: face.bytes };
    byId.set(id, embedded);
    const italic = face.style === "italic";
    byKey.set(faceKey(face.family, face.weight, italic), embedded);
    byKey.set(faceKey(face.originalFamily, face.weight, italic), embedded);
  }
  return byKey;
};

export type FontFaceRequest = {
  readonly fontFamily: string;
  readonly alternateFontFamily?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /**
   * The style the run measures with. The face's font-box ratios come from the
   * measurer, not from a table of our own, so a backend places a baseline
   * where the engine placed it.
   */
  readonly measureStyle: FontStyle;
};

export class FontTable {
  private readonly faces: DisplayFontFace[] = [];
  private readonly indexByKey = new Map<string, DisplayFontRef>();
  private readonly embeddedByKey: ReadonlyMap<string, DisplayEmbeddedFont>;

  /**
   * `embeddedFonts` are the package's own faces (`fonts/embeddedFonts.ts`).
   * Without them a face the measurer used but no host can supply travels as a
   * family name a backend has no way to resolve.
   */
  constructor(embeddedFonts: readonly EmbeddedFont[] = []) {
    this.embeddedByKey = indexEmbeddedFonts(embeddedFonts);
  }

  intern({
    fontFamily,
    alternateFontFamily,
    bold,
    italic,
    measureStyle,
  }: FontFaceRequest): DisplayFontRef {
    const stack = splitFallbackStack(
      resolveFontFamily(fontFamily, alternateFontFamily).cssFallback,
    );
    // The leading entry is the concrete face the canvas measured with, and the
    // only one a backend can embed; the trailing generic is the fallback.
    const family = stack.at(0) ?? fontFamily;
    const weight = bold ? CSS_BOLD_WEIGHT : CSS_NORMAL_WEIGHT;
    const isItalic = italic === true;
    const key = faceKey(family, weight, isItalic);

    const existing = this.indexByKey.get(key);
    if (existing !== undefined) {
      return existing;
    }

    const metrics = getFontMetrics(measureStyle);
    const fontSizePx = ptToPx(measureStyle.fontSize ?? metrics.fontSize);
    const scale = fontSizePx > 0 ? 1 / fontSizePx : 0;

    const index = this.faces.length;
    const embedded = this.embeddedByKey.get(key);
    this.faces.push({
      family,
      weight,
      italic: isItalic,
      generic: genericOf(stack),
      // Everything the measurer would fall through before reaching the
      // generic. Dropping it is what makes a backend paint a different face
      // from the one measured when the first family is not installed.
      fallbacks: stack
        .slice(1)
        .filter((entry) => GENERIC_FAMILIES[entry.toLowerCase() as GenericFamily] === undefined),
      fontBoxAscentRatio: metrics.fontBoxAscent * scale,
      fontBoxDescentRatio: metrics.fontBoxDescent * scale,
      ...(embedded === undefined ? {} : { embedded }),
    });
    this.indexByKey.set(key, index);
    return index;
  }

  snapshot(): readonly DisplayFontFace[] {
    return this.faces;
  }
}
