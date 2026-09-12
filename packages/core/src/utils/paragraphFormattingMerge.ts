import type { ParagraphFormatting, TabStop } from "../types/document";
import { mergeTextFormatting } from "./textFormattingMerge";

type ParagraphFormattingMergeDisposition =
  | "replace"
  | "merge-borders"
  | "merge-frame"
  | "merge-numbering"
  | "merge-run-properties"
  | "merge-tabs";

type ParagraphFormattingMergeDescriptor<Field extends keyof ParagraphFormatting> = {
  readonly field: Field;
  readonly merge: ParagraphFormattingMergeDisposition;
};

/**
 * Total merge ownership for the complete modeled `w:pPr` surface.
 *
 * Adding a paragraph field cannot silently bypass the cascade: TypeScript
 * requires a merge decision here, and this table drives the implementation
 * below rather than serving as a detached audit list.
 */
export const PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS = Object.freeze({
  alignment: { field: "alignment", merge: "replace" },
  bidi: { field: "bidi", merge: "replace" },
  kinsoku: { field: "kinsoku", merge: "replace" },
  overflowPunctuation: { field: "overflowPunctuation", merge: "replace" },
  spaceBefore: { field: "spaceBefore", merge: "replace" },
  spaceAfter: { field: "spaceAfter", merge: "replace" },
  lineSpacing: { field: "lineSpacing", merge: "replace" },
  lineSpacingRule: { field: "lineSpacingRule", merge: "replace" },
  snapToGrid: { field: "snapToGrid", merge: "replace" },
  beforeAutospacing: { field: "beforeAutospacing", merge: "replace" },
  afterAutospacing: { field: "afterAutospacing", merge: "replace" },
  spacingExplicit: { field: "spacingExplicit", merge: "replace" },
  indentLeft: { field: "indentLeft", merge: "replace" },
  indentRight: { field: "indentRight", merge: "replace" },
  indentFirstLine: { field: "indentFirstLine", merge: "replace" },
  hangingIndent: { field: "hangingIndent", merge: "replace" },
  borders: { field: "borders", merge: "merge-borders" },
  shading: { field: "shading", merge: "replace" },
  tabs: { field: "tabs", merge: "merge-tabs" },
  keepNext: { field: "keepNext", merge: "replace" },
  keepLines: { field: "keepLines", merge: "replace" },
  widowControl: { field: "widowControl", merge: "replace" },
  pageBreakBefore: { field: "pageBreakBefore", merge: "replace" },
  contextualSpacing: { field: "contextualSpacing", merge: "replace" },
  numPr: { field: "numPr", merge: "merge-numbering" },
  numPrFromStyle: { field: "numPrFromStyle", merge: "replace" },
  outlineLevel: { field: "outlineLevel", merge: "replace" },
  styleId: { field: "styleId", merge: "replace" },
  frame: { field: "frame", merge: "merge-frame" },
  suppressLineNumbers: { field: "suppressLineNumbers", merge: "replace" },
  suppressAutoHyphens: { field: "suppressAutoHyphens", merge: "replace" },
  runProperties: { field: "runProperties", merge: "merge-run-properties" },
  runInWithNext: { field: "runInWithNext", merge: "replace" },
} as const satisfies {
  [Field in keyof ParagraphFormatting]-?: ParagraphFormattingMergeDescriptor<Field>;
});

const copyDefinedParagraphProperty = <K extends keyof ParagraphFormatting>(
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

  for (const descriptor of Object.values(PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS)) {
    switch (descriptor.merge) {
      case "replace":
        copyDefinedParagraphProperty(result, source, descriptor.field);
        break;
      case "merge-borders":
        if (source.borders !== undefined) {
          result.borders = { ...result.borders, ...source.borders };
        }
        break;
      case "merge-frame":
        if (source.frame !== undefined) {
          result.frame = { ...result.frame, ...source.frame };
        }
        break;
      case "merge-numbering":
        if (source.numPr !== undefined) {
          result.numPr = { ...result.numPr, ...source.numPr };
        }
        break;
      case "merge-run-properties": {
        const merged = mergeTextFormatting(result.runProperties, source.runProperties);
        if (merged !== undefined) {
          result.runProperties = merged;
        }
        break;
      }
      case "merge-tabs":
        if (source.tabs !== undefined) {
          result.tabs = mergeParagraphTabStops(result.tabs, source.tabs);
        }
        break;
      default: {
        const exhaustive: never = descriptor;
        return exhaustive;
      }
    }
  }

  if (source.indentFirstLine !== undefined) {
    result.hangingIndent = source.hangingIndent === true;
  }

  return result;
}
