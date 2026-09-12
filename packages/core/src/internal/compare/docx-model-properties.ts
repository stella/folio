import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import { PARAGRAPH_FORMATTING_MERGE_DESCRIPTORS } from "../../utils/paragraphFormattingMerge";
import type {
  ParagraphFormatting,
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
  TextFormatting,
} from "../../types/document";
import {
  TABLE_CELL_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS,
  TABLE_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS,
  TABLE_ROW_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS,
  type CompareTableCellFormattingPropertyName,
  type CompareTableFormattingPropertyName,
  type CompareTableRowFormattingPropertyName,
} from "../../compare/table-format-properties";
import type {
  FolioContentProperty,
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
    return Object.freeze({
      type: "array",
      items: Object.freeze(value.map(docxCanonicalPropertyValue)),
    });
  }
  if (typeof value !== "object") {
    return panic("A modeled DOCX property is not representable as neutral data");
  }
  return Object.freeze({
    type: "object",
    entries: Object.freeze(
      Object.entries(value)
        .filter(([, child]) => child !== undefined)
        .map(([key, child]) => Object.freeze({ key, value: docxCanonicalPropertyValue(child) }))
        .toSorted(comparePropertyKeys),
    ),
  });
};

const modelPropertySet = <Value extends object, Field extends keyof Value & string>(
  value: Value | undefined,
  fields: readonly Field[],
): FolioContentPropertySet<Field> => {
  if (value === undefined) return Object.freeze([]);
  const properties: FolioContentProperty<Field>[] = [];
  for (const field of fields) {
    const property = value[field];
    if (property !== undefined) {
      properties.push(Object.freeze({ key: field, value: docxCanonicalPropertyValue(property) }));
    }
  }
  return Object.freeze(properties.toSorted(comparePropertyKeys));
};

type ComparisonPropertyDescriptor<Field extends string> = {
  readonly field: Field;
  readonly disposition: "property";
  readonly normalization: "exact" | "presence";
};

const OMITTED_COMPARISON_PROPERTY = Object.freeze({ type: "omitted" } as const);

const normalizedComparisonProperty = <Value extends object, Field extends keyof Value & string>(
  value: Value,
  descriptor: ComparisonPropertyDescriptor<Field>,
): Value[Field] | typeof OMITTED_COMPARISON_PROPERTY => {
  const property = value[descriptor.field];
  return property === undefined || (descriptor.normalization === "presence" && property === false)
    ? OMITTED_COMPARISON_PROPERTY
    : property;
};

const semanticFormattingValue = <Value extends object, Field extends keyof Value & string>(
  value: Value | undefined,
  descriptors: readonly ComparisonPropertyDescriptor<Field>[],
): Readonly<Partial<Pick<Value, Field>>> | null => {
  if (value === undefined) return null;
  const semantic: Partial<Pick<Value, Field>> = {};
  for (const descriptor of descriptors) {
    const property = normalizedComparisonProperty(value, descriptor);
    if (property === OMITTED_COMPARISON_PROPERTY) continue;
    Object.assign(semantic, { [descriptor.field]: property });
  }
  return Object.keys(semantic).length === 0 ? null : semantic;
};

const comparisonModelPropertySet = <Value extends object, Field extends keyof Value & string>(
  value: Value | undefined,
  descriptors: readonly ComparisonPropertyDescriptor<Field>[],
): FolioContentPropertySet<Field> => {
  if (value === undefined) return Object.freeze([]);
  const properties: FolioContentProperty<Field>[] = [];
  for (const descriptor of descriptors) {
    const property = normalizedComparisonProperty(value, descriptor);
    if (property === OMITTED_COMPARISON_PROPERTY) continue;
    properties.push(
      Object.freeze({
        key: descriptor.field,
        value: docxCanonicalPropertyValue(property),
      }),
    );
  }
  return Object.freeze(properties.toSorted(comparePropertyKeys));
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

/** @internal Directly authored table properties without transport source fields. */
export const docxTableFormattingSemanticValue = (
  formatting: TableFormatting | undefined,
): unknown => semanticFormattingValue(formatting, TABLE_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS);

/** @internal Directly authored row properties without transport source fields. */
export const docxTableRowFormattingSemanticValue = (
  formatting: TableRowFormatting | undefined,
): unknown =>
  semanticFormattingValue(formatting, TABLE_ROW_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS);

/** @internal Directly authored cell properties without structural or transport fields. */
export const docxTableCellFormattingSemanticValue = (
  formatting: TableCellFormatting | undefined,
): unknown =>
  semanticFormattingValue(formatting, TABLE_CELL_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS);

/** @internal Canonical neutral projection of directly authored table properties. */
export const docxTableFormattingProperties = (
  formatting: TableFormatting | undefined,
): FolioContentPropertySet<CompareTableFormattingPropertyName> =>
  comparisonModelPropertySet(formatting, TABLE_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS);

/** @internal Canonical neutral projection of directly authored row properties. */
export const docxTableRowFormattingProperties = (
  formatting: TableRowFormatting | undefined,
): FolioContentPropertySet<CompareTableRowFormattingPropertyName> =>
  comparisonModelPropertySet(formatting, TABLE_ROW_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS);

/** @internal Canonical neutral projection of directly authored cell properties. */
export const docxTableCellFormattingProperties = (
  formatting: TableCellFormatting | undefined,
): FolioContentPropertySet<CompareTableCellFormattingPropertyName> =>
  comparisonModelPropertySet(formatting, TABLE_CELL_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS);
