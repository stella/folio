/**
 * The font the editing surface draws a run's text with, resolved by the owners
 * the painter uses: the run's slots come from `extractRunFormatting` and each
 * slot's stack from `resolveFontFamily`, so text the browser draws natively,
 * such as an IME composition, has the metrics of the painted text.
 */

import type { Mark } from "prosemirror-model";

import {
  buildDocumentFontFacts,
  type DocumentFontFacts,
} from "../../fonts/documentFontFacts";
import {
  buildFontAlternates,
  type FontAlternates,
} from "../../fonts/fontAlternates";
import type { Theme } from "../../types/document";
import {
  type DeepReadonly,
  readModel,
  type ReadonlyDocument,
} from "../../types/documentModel";
import { resolveFontFamily } from "../../utils/fontResolver";
import { extractRunFormatting } from "./runMarkFormatting";

export type EditingFontContext = {
  theme?: DeepReadonly<Theme> | null | undefined;
  fontAlternates?: FontAlternates | undefined;
  embeddedFontFamilyMap?: ReadonlyMap<string, string> | undefined;
  documentFontFacts?: DocumentFontFacts | undefined;
};

/** The font context the painter resolves a document's runs with. */
export const editingFontContextOf = (
  document: ReadonlyDocument | null,
  embeddedFontFamilyMap: ReadonlyMap<string, string> | undefined,
): EditingFontContext => {
  // The font builders only read the table.
  const fontTable = document && readModel(document).package.fontTable;
  return {
    theme: document?.package.theme,
    fontAlternates: buildFontAlternates(fontTable),
    embeddedFontFamilyMap,
    documentFontFacts: buildDocumentFontFacts(fontTable),
  };
};

const GENERIC_FAMILIES: ReadonlySet<string> = new Set([
  "sans-serif",
  "serif",
  "monospace",
  "cursive",
]);

/**
 * One CSS `font-family` for a font mark's text: the Western slot's stack, then
 * the East Asian slot's, with the generic family last. The painter draws each
 * script with its own slot; the browser reaches the East Asian faces for the
 * characters the Western faces lack. `null` when the mark names no font.
 */
export const editingFontFamily = (
  mark: Mark,
  context: EditingFontContext,
): string | null => {
  const run = extractRunFormatting(
    [mark],
    context.theme,
    context.fontAlternates,
  );
  const stack = (
    family: string | undefined,
    alternateFontName: string | undefined,
  ): string[] =>
    family === undefined
      ? []
      : resolveFontFamily(family, {
          alternateFontName,
          embeddedFontFamilyMap: context.embeddedFontFamilyMap,
          documentFontFacts: context.documentFontFacts,
        })
          .cssFallback.split(",")
          .map((entry) => entry.trim())
          .filter((entry) => entry.length > 0);
  const families = [
    ...stack(run.fontFamily, run.alternateFontFamily),
    ...stack(run.eastAsiaFontFamily, run.eastAsiaAlternateFontFamily),
  ];
  if (families.length === 0) return null;
  const named = [
    ...new Set(families.filter((family) => !GENERIC_FAMILIES.has(family))),
  ];
  const generic = families.find((family) => GENERIC_FAMILIES.has(family));
  return [...named, ...(generic === undefined ? [] : [generic])].join(", ");
};
