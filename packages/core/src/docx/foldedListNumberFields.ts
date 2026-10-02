/**
 * `LISTNUM` fields a numbered paragraph draws as part of its list marker.
 *
 * The marker shows each field's cached result, so the field, and the tab that
 * followed it, must not also show on the line. They are not taken out of the
 * paragraph either: each stays in the content, at its source position, as the
 * markup it was read from, which shows nothing and is written back as it
 * stands. Being content is what keeps it in place: text typed around it moves
 * it as it moves any neighbour, a split leaves it in the half it stood in,
 * and a join carries it along.
 */

import type { ComplexField, ParagraphContent, PreservedInline, Run } from "../types/document";

/** Zero-width markup that can stand between a field and its tab. */
const RANGE_MARKER_TYPES: ReadonlySet<ParagraphContent["type"]> = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
]);

const isListNumberField = (content: ParagraphContent): content is ComplexField =>
  content.type === "complexField" &&
  (content.fieldType === "LISTNUM" ||
    content.instruction.trim().toUpperCase().startsWith("LISTNUM"));

/** A run that holds one tab and nothing else. */
export const isTabOnlyRun = (content: ParagraphContent): content is Run =>
  content.type === "run" && content.content.length === 1 && content.content[0]?.type === "tab";

const cachedText = (field: ComplexField): string => {
  let text = "";
  for (const run of field.fieldResult) {
    for (const piece of run.content) {
      if (piece.type === "text") {
        text += piece.text;
      }
    }
  }
  return text;
};

/** Whether `content` is a capture the fold left in place of a field or its tab. */
export const isFoldedListNumberCapture = (content: ParagraphContent): content is PreservedInline =>
  content.type === "preservedInline" && content.foldedListNumber !== undefined;

export type ListNumberFieldFold = {
  /** The content, each folded field and its tab replaced by its own markup. */
  content: ParagraphContent[];
  /** The cached display of each folded field that has one, in source order. */
  cached: string[];
  /** Every `LISTNUM` field the content holds, folded or not. */
  fieldCount: number;
};

/**
 * Replace every `LISTNUM` field, and the tab after it, by the markup it was
 * read from.
 *
 * `sourceMarkupOf` answers that markup. A field it has none for is left as
 * the field it is: nothing is written for it that was not read. Bookmark and
 * comment markers between a field and its tab stay where they are, and the
 * tab is still the field's when only they separate the two.
 */
export const foldListNumberFields = (
  content: readonly ParagraphContent[],
  sourceMarkupOf: (item: ComplexField | Run) => string | undefined,
): ListNumberFieldFold => {
  const folded: ParagraphContent[] = [];
  const cached: string[] = [];
  let fieldCount = 0;
  let awaitingTab = false;

  for (const item of content) {
    if (awaitingTab && RANGE_MARKER_TYPES.has(item.type)) {
      folded.push(item);
      continue;
    }
    if (awaitingTab) {
      awaitingTab = false;
      const tabMarkup = isTabOnlyRun(item) ? sourceMarkupOf(item) : undefined;
      if (tabMarkup !== undefined) {
        folded.push({ type: "preservedInline", xml: tabMarkup, text: "", foldedListNumber: "tab" });
        continue;
      }
    }
    if (isListNumberField(item)) {
      fieldCount += 1;
      const markup = sourceMarkupOf(item);
      if (markup !== undefined) {
        folded.push({ type: "preservedInline", xml: markup, text: "", foldedListNumber: "field" });
        const text = cachedText(item);
        if (text) {
          cached.push(text);
        }
        awaitingTab = true;
        continue;
      }
    }
    folded.push(item);
  }

  return { content: folded, cached, fieldCount };
};
