import type { TableCellFormatting, TableFormatting, TableRowFormatting } from "../types/document";
import type { FolioContentPropertyChange } from "./content-types";

type TableFormatPropertyNormalization = "exact" | "presence";

type TableFormatFieldDescriptor<Field extends string> =
  | {
      readonly field: Field;
      readonly disposition: "property";
      readonly normalization: TableFormatPropertyNormalization;
    }
  | {
      readonly field: Field;
      readonly disposition: "structure" | "transport";
    };

type TotalTableFormatFieldDescriptors<Value> = {
  readonly [Field in keyof Value]-?: TableFormatFieldDescriptor<Extract<Field, string>>;
};

const isTableFormatPropertyDescriptor = <Descriptor extends TableFormatFieldDescriptor<string>>(
  descriptor: Descriptor,
): descriptor is Extract<Descriptor, { readonly disposition: "property" }> =>
  descriptor.disposition === "property";

/**
 * Total ownership of the modeled table property surface.
 *
 * The DOCX adapter uses these records to build both its executable property
 * program and the public semantic delta. Adding a model field therefore
 * requires an explicit property, structure, or transport decision here.
 * @internal
 */
export const TABLE_FORMATTING_COMPARISON_FIELD_DESCRIPTORS = Object.freeze({
  width: Object.freeze({ field: "width", disposition: "property", normalization: "exact" }),
  justification: Object.freeze({
    field: "justification",
    disposition: "property",
    normalization: "exact",
  }),
  cellSpacing: Object.freeze({
    field: "cellSpacing",
    disposition: "property",
    normalization: "exact",
  }),
  indent: Object.freeze({ field: "indent", disposition: "property", normalization: "exact" }),
  borders: Object.freeze({ field: "borders", disposition: "property", normalization: "exact" }),
  cellMargins: Object.freeze({
    field: "cellMargins",
    disposition: "property",
    normalization: "exact",
  }),
  layout: Object.freeze({ field: "layout", disposition: "property", normalization: "exact" }),
  styleId: Object.freeze({ field: "styleId", disposition: "property", normalization: "exact" }),
  look: Object.freeze({ field: "look", disposition: "property", normalization: "exact" }),
  shading: Object.freeze({ field: "shading", disposition: "property", normalization: "exact" }),
  overlap: Object.freeze({ field: "overlap", disposition: "property", normalization: "exact" }),
  floating: Object.freeze({ field: "floating", disposition: "property", normalization: "exact" }),
  bidi: Object.freeze({ field: "bidi", disposition: "property", normalization: "presence" }),
  gridSourceXml: Object.freeze({ field: "gridSourceXml", disposition: "transport" }),
  sourceXml: Object.freeze({ field: "sourceXml", disposition: "transport" }),
} as const satisfies TotalTableFormatFieldDescriptors<TableFormatting>);

/** @internal */
export const TABLE_ROW_FORMATTING_COMPARISON_FIELD_DESCRIPTORS = Object.freeze({
  gridBefore: Object.freeze({
    field: "gridBefore",
    disposition: "property",
    normalization: "exact",
  }),
  widthBefore: Object.freeze({
    field: "widthBefore",
    disposition: "property",
    normalization: "exact",
  }),
  gridAfter: Object.freeze({
    field: "gridAfter",
    disposition: "property",
    normalization: "exact",
  }),
  widthAfter: Object.freeze({
    field: "widthAfter",
    disposition: "property",
    normalization: "exact",
  }),
  height: Object.freeze({ field: "height", disposition: "property", normalization: "exact" }),
  heightRule: Object.freeze({
    field: "heightRule",
    disposition: "property",
    normalization: "exact",
  }),
  header: Object.freeze({ field: "header", disposition: "property", normalization: "presence" }),
  cantSplit: Object.freeze({
    field: "cantSplit",
    disposition: "property",
    normalization: "presence",
  }),
  justification: Object.freeze({
    field: "justification",
    disposition: "property",
    normalization: "exact",
  }),
  hidden: Object.freeze({ field: "hidden", disposition: "property", normalization: "presence" }),
  conditionalFormat: Object.freeze({
    field: "conditionalFormat",
    disposition: "property",
    normalization: "exact",
  }),
  sourceXml: Object.freeze({ field: "sourceXml", disposition: "transport" }),
} as const satisfies TotalTableFormatFieldDescriptors<TableRowFormatting>);

/** @internal */
export const TABLE_CELL_FORMATTING_COMPARISON_FIELD_DESCRIPTORS = Object.freeze({
  width: Object.freeze({ field: "width", disposition: "property", normalization: "exact" }),
  borders: Object.freeze({ field: "borders", disposition: "property", normalization: "exact" }),
  margins: Object.freeze({ field: "margins", disposition: "property", normalization: "exact" }),
  shading: Object.freeze({ field: "shading", disposition: "property", normalization: "exact" }),
  verticalAlign: Object.freeze({
    field: "verticalAlign",
    disposition: "property",
    normalization: "exact",
  }),
  textDirection: Object.freeze({
    field: "textDirection",
    disposition: "property",
    normalization: "exact",
  }),
  gridSpan: Object.freeze({ field: "gridSpan", disposition: "structure" }),
  vMerge: Object.freeze({ field: "vMerge", disposition: "structure" }),
  fitText: Object.freeze({
    field: "fitText",
    disposition: "property",
    normalization: "presence",
  }),
  noWrap: Object.freeze({ field: "noWrap", disposition: "property", normalization: "presence" }),
  hideMark: Object.freeze({
    field: "hideMark",
    disposition: "property",
    normalization: "presence",
  }),
  conditionalFormat: Object.freeze({
    field: "conditionalFormat",
    disposition: "property",
    normalization: "exact",
  }),
  sourceXml: Object.freeze({ field: "sourceXml", disposition: "transport" }),
} as const satisfies TotalTableFormatFieldDescriptors<TableCellFormatting>);

/** @internal Canonical semantic table fields, derived from total ownership. */
export const TABLE_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS = Object.freeze(
  Object.values(TABLE_FORMATTING_COMPARISON_FIELD_DESCRIPTORS).filter(
    isTableFormatPropertyDescriptor,
  ),
);

/** @internal Canonical semantic row fields, derived from total ownership. */
export const TABLE_ROW_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS = Object.freeze(
  Object.values(TABLE_ROW_FORMATTING_COMPARISON_FIELD_DESCRIPTORS).filter(
    isTableFormatPropertyDescriptor,
  ),
);

/** @internal Canonical semantic cell fields, derived from total ownership. */
export const TABLE_CELL_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS = Object.freeze(
  Object.values(TABLE_CELL_FORMATTING_COMPARISON_FIELD_DESCRIPTORS).filter(
    isTableFormatPropertyDescriptor,
  ),
);

export type CompareTableFormattingPropertyName =
  (typeof TABLE_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS)[number]["field"];
export type CompareTableRowFormattingPropertyName =
  (typeof TABLE_ROW_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS)[number]["field"];
export type CompareTableCellFormattingPropertyName =
  (typeof TABLE_CELL_FORMATTING_COMPARISON_PROPERTY_DESCRIPTORS)[number]["field"];

export type CompareTableFormattingPropertyChange =
  FolioContentPropertyChange<CompareTableFormattingPropertyName>;
export type CompareTableRowFormattingPropertyChange =
  FolioContentPropertyChange<CompareTableRowFormattingPropertyName>;
export type CompareTableCellFormattingPropertyChange =
  FolioContentPropertyChange<CompareTableCellFormattingPropertyName>;

export type CompareTableCoordinate = {
  readonly tableIndex: number;
};

export type CompareTableRowCoordinate = {
  readonly tableIndex: number;
  readonly rowIndex: number;
};

export type CompareTableCellCoordinate = {
  readonly tableIndex: number;
  readonly rowIndex: number;
  readonly cellIndex: number;
};

/** Scope-specific semantic payload for one tracked table property change. */
export type CompareTableFormatDetails =
  | {
      readonly scope: "table";
      readonly base: Readonly<CompareTableCoordinate>;
      readonly target: Readonly<CompareTableCoordinate>;
      readonly properties: readonly CompareTableFormattingPropertyChange[];
    }
  | {
      readonly scope: "row";
      readonly base: Readonly<CompareTableRowCoordinate>;
      readonly target: Readonly<CompareTableRowCoordinate>;
      readonly properties: readonly CompareTableRowFormattingPropertyChange[];
    }
  | {
      readonly scope: "cell";
      readonly base: Readonly<CompareTableCellCoordinate>;
      readonly target: Readonly<CompareTableCellCoordinate>;
      readonly properties: readonly CompareTableCellFormattingPropertyChange[];
    };
