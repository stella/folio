import type { ParagraphFormatting, TabStop } from "../types/document";
import { mergeTextFormatting } from "./textFormattingMerge";

const PARAGRAPH_REPLACE_KEYS = [
  "alignment",
  "bidi",
  "kinsoku",
  "overflowPunctuation",
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
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
    const result = { ...source };
    if (source.tabs !== undefined) {
      result.tabs = [...source.tabs];
    }
    return result;
  }

  const result: ParagraphFormatting = { ...target };

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
  if (source.numPr !== undefined) {
    result.numPr = { ...result.numPr, ...source.numPr };
  }
  if (source.frame !== undefined) {
    result.frame = { ...result.frame, ...source.frame };
  }
  if (source.tabs !== undefined) {
    result.tabs = mergeParagraphTabStops(result.tabs, source.tabs);
  }

  return result;
}
