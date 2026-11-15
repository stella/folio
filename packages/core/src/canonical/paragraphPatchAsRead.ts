/**
 * The paragraph properties an operation states, as the parser reads back the
 * `w:pPr` the save writes for them.
 *
 * An operation's patch is the one place an edited paragraph's formatting is
 * produced, so it states these fields the way the saved file reopens to; a
 * paragraph no operation names keeps its record as it is.
 */

import {
  paragraphSpacingExplicit,
  type SpacingExplicit,
} from "@stll/docx-core/model";

import type {
  FolioAIParagraphIndentation,
  FolioAIParagraphSpacing,
} from "../ai-edits/types";

/** The spacing sides `spacing` states, marked explicit by the parser's rule. */
export const spacingSidesAsRead = (
  spacing: FolioAIParagraphSpacing,
): SpacingExplicit | undefined => paragraphSpacingExplicit(spacing);

/**
 * `indentation` as the parser reads back the `w:ind` it saves to: the element
 * writes a first line or a hanging magnitude, and a hanging one reads as a
 * first line pulled out to the left, flagged hanging. A flag without a first
 * line writes nothing, and an unset flag is how the parser reads a first line.
 */
export const indentationAsRead = (
  indentation: FolioAIParagraphIndentation,
): FolioAIParagraphIndentation => {
  const { indentFirstLine, hangingIndent, ...sides } = indentation;
  if (indentFirstLine === undefined) return sides;
  if (hangingIndent !== true) return { ...sides, indentFirstLine };
  return {
    ...sides,
    indentFirstLine: indentFirstLine === 0 ? 0 : -Math.abs(indentFirstLine),
    hangingIndent: true,
  };
};
