import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import { PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS } from "../../utils/paragraphFormattingMerge";
import type { ParagraphFormatting, TextFormatting } from "../../types/document";
import type {
  FolioContentPropertySet,
  FolioContentPropertyValue,
} from "../../compare/content-types";

const comparePropertyKeys = (left: { key: string }, right: { key: string }): number => {
  if (left.key === right.key) return 0;
  return left.key < right.key ? -1 : 1;
};

export const docxCanonicalPropertyValue = (value: unknown): FolioContentPropertyValue => {
  if (value === null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") return Object.is(value, -0) ? 0 : value;
  if (Array.isArray(value)) {
    return { type: "array", items: value.map(docxCanonicalPropertyValue) };
  }
  if (typeof value !== "object") {
    return panic("A modeled DOCX property is not representable as neutral data");
  }
  return {
    type: "object",
    entries: Object.entries(value)
      .filter(([, child]) => child !== undefined)
      .map(([key, child]) => ({ key, value: docxCanonicalPropertyValue(child) }))
      .toSorted(comparePropertyKeys),
  };
};

const modelPropertySet = <Value extends object, Field extends keyof Value>(
  value: Value | undefined,
  fields: readonly Field[],
): FolioContentPropertySet => {
  if (value === undefined) return [];
  const properties: { key: string; value: FolioContentPropertyValue }[] = [];
  for (const field of fields) {
    const property = value[field];
    if (property !== undefined) {
      properties.push({ key: String(field), value: docxCanonicalPropertyValue(property) });
    }
  }
  return properties.toSorted(comparePropertyKeys);
};

const PARAGRAPH_FORMATTING_FIELDS = Object.freeze(
  Object.values(PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS).map(({ field }) => field),
);

const TEXT_FORMATTING_FIELDS = Object.freeze(
  Object.values(TEXT_FORMATTING_PROPERTY_DESCRIPTORS).map(({ field }) => field),
);

/** @internal Canonical neutral projection of a modeled paragraph-property record. */
export const docxParagraphFormattingProperties = (
  formatting: ParagraphFormatting | undefined,
): FolioContentPropertySet => modelPropertySet(formatting, PARAGRAPH_FORMATTING_FIELDS);

/** @internal Canonical neutral projection of a modeled run-property record. */
export const docxTextFormattingProperties = (
  formatting: TextFormatting | undefined,
): FolioContentPropertySet => modelPropertySet(formatting, TEXT_FORMATTING_FIELDS);
