/**
 * Run properties of a paragraph mark (ECMA-376 §17.3.1.29 `w:pPr/w:rPr`).
 *
 * The mark resolves like any run of its paragraph: document defaults, the
 * table style, the paragraph style chain, then a character style the mark
 * names in `w:rStyle` (whose toggles combine per §17.7.3), then the mark's
 * direct properties. The default character style is not applied to the mark.
 */
import {
  resolveParagraphBodyRunFormatting,
  type RunStyleResolver,
} from "../../prosemirror/runStyleFormatting";
import type { ParagraphAttrs as PMParagraphAttrs } from "../../prosemirror/schema/nodes";
import { cascadeStyleTextFormatting } from "../../prosemirror/styles/styleToggleCascade";
import type { TextFormatting } from "../../types/document";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";

/**
 * The properties of the character style a mark names in `w:rStyle`. An id that
 * is missing or names a style of another type contributes nothing.
 */
const markCharacterStyleFormatting = (
  mark: TextFormatting | undefined,
  styleResolver: RunStyleResolver,
): TextFormatting | undefined => {
  const styleId = mark?.styleId;
  if (styleId === undefined || styleResolver.getStyle(styleId)?.type !== "character") {
    return undefined;
  }
  return styleResolver.getRunStyleOwnProperties(styleId);
};

/** The mark's own properties: its `w:rStyle` character style under its direct `w:rPr`. */
export const resolveParagraphMarkOwnFormatting = (
  mark: TextFormatting | undefined,
  styleResolver: RunStyleResolver,
): TextFormatting | undefined => {
  const characterStyle = markCharacterStyleFormatting(mark, styleResolver);
  if (characterStyle === undefined) {
    return mark;
  }
  return cascadeStyleTextFormatting([
    { formatting: characterStyle, type: "style" },
    { formatting: mark, type: "direct" },
  ]).formatting;
};

/** The paragraph mark's fully resolved run properties. */
export const resolveParagraphMarkFormatting = (
  pmAttrs: PMParagraphAttrs,
  styleResolver: RunStyleResolver,
): TextFormatting | undefined => {
  const mark = pmAttrs._originalFormatting?.runProperties;
  const { baseFormatting, baseToggleCascade } = resolveParagraphBodyRunFormatting({
    styleId: pmAttrs.styleId ?? undefined,
    styleResolver,
    tableRunFormatting: pmAttrs._tableRunFormatting,
  });
  const characterStyle = markCharacterStyleFormatting(mark, styleResolver);
  return cascadeStyleTextFormatting(
    [
      { cascade: baseToggleCascade, type: "carried" },
      { formatting: characterStyle, type: "style" },
      { formatting: mark, type: "direct" },
    ],
    {
      ordinaryFormatting: mergeTextFormatting(
        mergeTextFormatting(baseFormatting, characterStyle),
        mark,
      ),
    },
  ).formatting;
};
