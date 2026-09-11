import { panic } from "better-result";

import type { ParagraphFormatting, TabStop } from "../types/document";
import {
  PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR,
  PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST,
} from "../docx/paragraphPropertyDescriptor";
import { mergeTextFormatting } from "./textFormattingMerge";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

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

  const result: ParagraphFormatting = { ...target };
  for (const key of PARAGRAPH_FORMATTING_PROPERTY_KEY_LIST) {
    const sourceValue = source[key];
    if (sourceValue === undefined) {
      continue;
    }
    const { cascade } = PARAGRAPH_FORMATTING_PROPERTY_DESCRIPTOR[key];
    switch (cascade) {
      case "replace":
        Reflect.set(result, key, sourceValue);
        break;
      case "first-line":
        result.indentFirstLine = source.indentFirstLine;
        result.hangingIndent = source.hangingIndent === true;
        break;
      case "fieldwise":
        if (!isRecord(sourceValue)) {
          return panic(`Paragraph property ${key} requires a fieldwise record`);
        }
        Reflect.set(result, key, {
          ...(isRecord(result[key]) ? result[key] : {}),
          ...sourceValue,
        });
        break;
      case "run-properties": {
        const merged = mergeTextFormatting(result.runProperties, source.runProperties);
        if (merged) {
          result.runProperties = merged;
        }
        break;
      }
      case "tab-stops":
        result.tabs = mergeParagraphTabStops(result.tabs, source.tabs);
        break;
      case "preserve-derived":
        break;
      default: {
        const exhaustive: never = cascade;
        return exhaustive;
      }
    }
  }

  return result;
}
