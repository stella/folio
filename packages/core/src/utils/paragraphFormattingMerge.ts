import type { ParagraphFormatting, TabStop } from "../types/document";
import { mergeTextFormatting } from "./textFormattingMerge";
import { mergeParagraphNumbering } from "@stll/docx-core/model";

const PARAGRAPH_REPLACE_KEYS = [
  "alignment",
  "bidi",
  "kinsoku",
  "overflowPunctuation",
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "snapToGrid",
  "beforeAutospacing",
  "afterAutospacing",
  "spacingExplicit",
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
  "shading",
  "keepNext",
  "keepLines",
  "widowControl",
  "pageBreakBefore",
  "contextualSpacing",
  "outlineLevel",
  "styleId",
  "suppressLineNumbers",
  "suppressAutoHyphens",
  "runInWithNext",
] as const satisfies readonly (keyof ParagraphFormatting)[];

type ParagraphReplaceKey = (typeof PARAGRAPH_REPLACE_KEYS)[number];

const copyDefinedParagraphProperty = <K extends ParagraphReplaceKey>(
  target: Pick<ParagraphFormatting, K>,
  source: Pick<ParagraphFormatting, K>,
  key: K,
): void => {
  const value = source[key];
  if (value !== undefined) {
    target[key] = value;
  }
};

/**
 * Merge custom tab stops across OOXML paragraph-property layers.
 *
 * Tabs cascade by position rather than replacing the inherited collection:
 * a higher-priority stop supersedes the stop at the same position, while a
 * `clear` stop removes that inherited position during layout. Keeping the
 * clear entry also lets the tab calculator suppress an automatic stop at the
 * same position.
 */
export function mergeParagraphTabStops(
  inherited: TabStop[] | undefined,
  direct: TabStop[],
): TabStop[];
export function mergeParagraphTabStops(
  inherited: TabStop[] | undefined,
  direct: TabStop[] | undefined,
): TabStop[] | undefined;
export function mergeParagraphTabStops(
  inherited: TabStop[] | undefined,
  direct: TabStop[] | undefined,
): TabStop[] | undefined {
  if (direct === undefined) {
    return inherited === undefined ? undefined : [...inherited];
  }

  const stopsByPosition = new Map<number, TabStop>();
  for (const stop of inherited ?? []) {
    stopsByPosition.set(stop.position, stop);
  }
  for (const stop of direct) {
    stopsByPosition.set(stop.position, stop);
  }

  return [...stopsByPosition.values()].toSorted((a, b) => a.position - b.position);
}

/**
 * A resolved cascade carries no captured bytes.
 *
 * `preserved` holds the markup of the one element it was read from, and the
 * two attribute remainders beside it hold the attributes of the `w:ind` and
 * `w:spacing` that element carried. A style's `w:pPr` is not a paragraph's
 * direct formatting, and a merge that let a spread carry either would write
 * the same bytes at two tiers and make an inherited value outrank the tier it
 * came from — the rule #873 states for every other inherited `w:pPr` value.
 * Every caller of the merge below builds a resolved cascade rather than a
 * saved element, so the drop is unconditional and belongs here, where a new
 * field cannot slip past it in a spread.
 *
 * The remainders inside `borders`, `shading` and `tabs` are not dropped, and
 * that is the same rule rather than an exception: those records replace
 * wholesale, so an inherited one is written back only where the whole record
 * is, and the attributes the author wrote on it belong with it.
 */
const withoutPreservedMarkup = (formatting: ParagraphFormatting): ParagraphFormatting => {
  const {
    preserved: _preserved,
    indentPreservedAttributes: _indentPreservedAttributes,
    spacingPreservedAttributes: _spacingPreservedAttributes,
    ...resolved
  } = formatting;
  return resolved;
};

/**
 * Merge paragraph properties for OOXML style cascade resolution.
 *
 * The source is the higher-priority layer. Most `w:pPr` properties replace an
 * inherited value when present; nested child containers merge by child field;
 * tabs merge by position; paragraph mark `w:rPr` uses the run-formatting
 * merge rules.
 */
export function mergeParagraphFormatting(
  target: ParagraphFormatting | undefined,
  source: ParagraphFormatting | undefined,
): ParagraphFormatting | undefined {
  if (!source) {
    return target;
  }
  if (!target) {
    const result = withoutPreservedMarkup(source);
    if (source.tabs !== undefined) {
      result.tabs = [...source.tabs];
    }
    return result;
  }

  const result: ParagraphFormatting = withoutPreservedMarkup(target);

  for (const key of PARAGRAPH_REPLACE_KEYS) {
    copyDefinedParagraphProperty(result, source, key);
  }

  if (source.indentFirstLine !== undefined) {
    result.hangingIndent = source.hangingIndent === true;
  }

  const mergedRunProperties = mergeTextFormatting(result.runProperties, source.runProperties);
  if (mergedRunProperties) {
    result.runProperties = mergedRunProperties;
  }

  if (source.borders !== undefined) {
    result.borders = { ...result.borders, ...source.borders };
  }
  // Not a spread: `w:numId` and `w:ilvl` inherit independently, and the fold
  // that says so owns the rule. A spread would leave the lower tier's id next
  // to the upper tier's `levelOnly` discriminator.
  const mergedNumbering = mergeParagraphNumbering(result.numPr, source.numPr);
  if (mergedNumbering !== undefined) {
    result.numPr = mergedNumbering;
  }
  if (source.frame !== undefined) {
    result.frame = { ...result.frame, ...source.frame };
  }
  if (source.tabs !== undefined) {
    result.tabs = mergeParagraphTabStops(result.tabs, source.tabs);
  }

  return result;
}
