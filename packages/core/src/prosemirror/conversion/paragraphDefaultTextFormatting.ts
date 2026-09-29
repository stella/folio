/**
 * The run formatting a paragraph's style chain lends its text, read from the
 * styles alone: no document conversion, so editing code can reach it.
 */

import type { StyleEngine } from "../../style-engine";
import type { Paragraph, TextFormatting } from "../../types/document";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
import { stripParagraphMarkOnlyFormatting } from "../runStyleFormatting";
import { cascadeStyleTextFormatting } from "../styles/styleToggleCascade";

/**
 * Resolve an embedded character-style reference without importing
 * `docDefaults`. The caller already has the paragraph cascade, including
 * document defaults, and will layer these own properties over it.
 */
export type ParagraphDefaultFormattingResolver = Pick<
  StyleEngine,
  | "getStyle"
  | "getDocDefaults"
  | "getDefaultParagraphStyle"
  | "getDefaultCharacterStyle"
  | "getRunStyleOwnProperties"
>;

export function resolveRunFormattingWithoutDefaults(
  formatting: TextFormatting | undefined,
  styleResolver: ParagraphDefaultFormattingResolver | null,
): TextFormatting | undefined {
  if (!formatting || !styleResolver) {
    return formatting;
  }

  const characterStyleFormatting = formatting.styleId
    ? styleResolver.getRunStyleOwnProperties(formatting.styleId)
    : undefined;
  return cascadeStyleTextFormatting([
    { formatting: characterStyleFormatting, type: "style" },
    { formatting, type: "direct" },
  ]).formatting;
}

/** @internal Recompute a paragraph's inherited run defaults from authored package state. */
export function resolveParagraphDefaultTextFormatting(
  styleId: string | undefined,
  formatting: Paragraph["formatting"] | undefined,
  styleResolver: ParagraphDefaultFormattingResolver,
  options: { includeParagraphMarkRunProperties?: boolean } = {},
): TextFormatting | undefined {
  const style = styleId
    ? (styleResolver.getStyle(styleId) ?? styleResolver.getDefaultParagraphStyle())
    : styleResolver.getDefaultParagraphStyle();
  const paragraphStyleRpr = style?.type === "paragraph" ? style.rPr : undefined;
  // The pPr/rPr block describes the paragraph mark only — see the comment on
  // `stripParagraphMarkOnlyFormatting`. We must NOT route this through
  // `resolveTextFormatting` here, because that folds docDefaults back into
  // the run properties and then overwrites the paragraph style's font
  // (e.g. FootnoteText's Times New Roman) with the docDefault Calibri when
  // merged into the cascade below.
  const rawParagraphMarkRpr =
    options.includeParagraphMarkRunProperties === false ? undefined : formatting?.runProperties;
  const paragraphRunProperties = rawParagraphMarkRpr
    ? stripParagraphMarkOnlyFormatting(
        resolveRunFormattingWithoutDefaults(rawParagraphMarkRpr, styleResolver) ?? {},
      )
    : undefined;

  const orderedBodyToggleFormatting = cascadeStyleTextFormatting(
    [
      { formatting: styleResolver.getDocDefaults()?.rPr, type: "defaults" },
      { formatting: paragraphStyleRpr, type: "style" },
      { formatting: styleResolver.getDefaultCharacterStyle()?.rPr, type: "style" },
    ],
    {
      ordinaryFormatting: mergeTextFormatting(
        mergeTextFormatting(
          styleResolver.getDocDefaults()?.rPr,
          styleResolver.getDefaultCharacterStyle()?.rPr,
        ),
        paragraphStyleRpr,
      ),
    },
  );
  const bodyRunDefaults = orderedBodyToggleFormatting.formatting;
  return cascadeStyleTextFormatting(
    [
      { cascade: orderedBodyToggleFormatting, type: "carried" },
      { formatting: paragraphRunProperties, type: "direct" },
    ],
    {
      ordinaryFormatting: mergeTextFormatting(bodyRunDefaults, paragraphRunProperties),
    },
  ).formatting;
}
