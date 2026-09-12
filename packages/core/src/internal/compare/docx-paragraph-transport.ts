import { panic } from "better-result";

import type {
  FolioAIBlockParagraphProperties,
  FolioAIBlockTableLocation,
  FolioAIParagraphSpacing,
} from "../../ai-edits/types";
import type {
  FolioContentBlock,
  FolioContentPropertyChange,
  FolioContentPropertySet,
  FolioContentPropertyValue,
  FolioContentTableLocation,
} from "../../compare/content-types";
import type { ParagraphAlignment, ParagraphFormatting } from "../../types/document";

type ParagraphTransportDisposition =
  | "alignment"
  | "list-level"
  | "spacing"
  | "style"
  | "unsupported";

/** One transport decision for every serializer-visible paragraph property. */
const DOCX_PARAGRAPH_TRANSPORT_DISPOSITIONS = Object.freeze({
  alignment: "alignment",
  bidi: "unsupported",
  kinsoku: "unsupported",
  overflowPunctuation: "unsupported",
  spaceBefore: "spacing",
  spaceAfter: "spacing",
  lineSpacing: "spacing",
  lineSpacingRule: "spacing",
  snapToGrid: "unsupported",
  beforeAutospacing: "spacing",
  afterAutospacing: "spacing",
  spacingExplicit: "unsupported",
  indentLeft: "unsupported",
  indentRight: "unsupported",
  indentFirstLine: "unsupported",
  hangingIndent: "unsupported",
  borders: "unsupported",
  shading: "unsupported",
  tabs: "unsupported",
  keepNext: "unsupported",
  keepLines: "unsupported",
  widowControl: "unsupported",
  pageBreakBefore: "unsupported",
  contextualSpacing: "unsupported",
  numPr: "list-level",
  numPrFromStyle: "unsupported",
  outlineLevel: "unsupported",
  styleId: "style",
  frame: "unsupported",
  suppressLineNumbers: "unsupported",
  suppressAutoHyphens: "unsupported",
  runProperties: "unsupported",
  runInWithNext: "unsupported",
} as const satisfies Record<keyof ParagraphFormatting, ParagraphTransportDisposition>);

const DOCX_PARAGRAPH_PROPERTY_FIELDS = Object.freeze({
  styleId: { field: "styleId" },
  listLevel: { field: "listLevel" },
  alignment: { field: "alignment" },
  spacing: { field: "spacing" },
} as const satisfies {
  [Field in keyof FolioAIBlockParagraphProperties]-?: { readonly field: Field };
});

const DOCX_PARAGRAPH_SPACING_FIELDS = Object.freeze({
  spaceBefore: { field: "spaceBefore" },
  spaceAfter: { field: "spaceAfter" },
  lineSpacing: { field: "lineSpacing" },
  lineSpacingRule: { field: "lineSpacingRule" },
  beforeAutospacing: { field: "beforeAutospacing" },
  afterAutospacing: { field: "afterAutospacing" },
} as const satisfies {
  [Field in keyof FolioAIParagraphSpacing]-?: { readonly field: Field };
});

/** Exact equality over the complete paragraph-operation property vocabulary. */
export const docxParagraphPropertiesEqual = (
  left: Readonly<FolioAIBlockParagraphProperties>,
  right: Readonly<FolioAIBlockParagraphProperties>,
): boolean =>
  Object.values(DOCX_PARAGRAPH_PROPERTY_FIELDS).every(({ field }) => {
    if (field !== "spacing") return left[field] === right[field];
    const leftSpacing = left.spacing;
    const rightSpacing = right.spacing;
    if (
      leftSpacing === undefined ||
      leftSpacing === null ||
      rightSpacing === undefined ||
      rightSpacing === null
    ) {
      return leftSpacing === rightSpacing;
    }
    return Object.values(DOCX_PARAGRAPH_SPACING_FIELDS).every(
      ({ field: spacingField }) =>
        leftSpacing[spacingField] === rightSpacing[spacingField],
    );
  });

const propertyValue = (
  properties: FolioContentPropertySet,
  key: keyof ParagraphFormatting,
): FolioContentPropertyValue | null =>
  properties.find((property) => property.key === key)?.value ?? null;

const stringValue = (value: FolioContentPropertyValue | null, field: string): string | null => {
  if (value === null || typeof value === "string") return value;
  return panic("A canonical DOCX paragraph property is not a string", { field });
};

const numberValue = (value: FolioContentPropertyValue | null, field: string): number | null => {
  if (value === null || typeof value === "number") return value;
  return panic("A canonical DOCX paragraph property is not a number", { field });
};

const booleanValue = (value: FolioContentPropertyValue | null, field: string): boolean | null => {
  if (value === null || typeof value === "boolean") return value;
  return panic("A canonical DOCX paragraph property is not a boolean", { field });
};

const alignmentValue = (
  value: FolioContentPropertyValue | null,
): ParagraphAlignment | null => {
  const alignment = stringValue(value, "alignment");
  switch (alignment) {
    case null:
    case "left":
    case "center":
    case "right":
    case "both":
    case "distribute":
    case "mediumKashida":
    case "highKashida":
    case "lowKashida":
    case "thaiDistribute":
      return alignment;
    default:
      return panic("A canonical DOCX paragraph alignment is invalid", { alignment });
  }
};

const spacingFromProperties = (
  properties: FolioContentPropertySet,
): FolioAIParagraphSpacing | null => {
  const spacing: FolioAIParagraphSpacing = {};
  const spaceBefore = numberValue(propertyValue(properties, "spaceBefore"), "spaceBefore");
  const spaceAfter = numberValue(propertyValue(properties, "spaceAfter"), "spaceAfter");
  const lineSpacing = numberValue(propertyValue(properties, "lineSpacing"), "lineSpacing");
  const lineSpacingRule = stringValue(
    propertyValue(properties, "lineSpacingRule"),
    "lineSpacingRule",
  );
  const beforeAutospacing = booleanValue(
    propertyValue(properties, "beforeAutospacing"),
    "beforeAutospacing",
  );
  const afterAutospacing = booleanValue(
    propertyValue(properties, "afterAutospacing"),
    "afterAutospacing",
  );
  if (spaceBefore !== null) spacing.spaceBefore = spaceBefore;
  if (spaceAfter !== null) spacing.spaceAfter = spaceAfter;
  if (lineSpacing !== null) spacing.lineSpacing = lineSpacing;
  switch (lineSpacingRule) {
    case null:
      break;
    case "auto":
    case "exact":
    case "atLeast":
      spacing.lineSpacingRule = lineSpacingRule;
      break;
    default:
      return panic("A canonical DOCX line-spacing rule is invalid", { lineSpacingRule });
  }
  if (beforeAutospacing !== null) spacing.beforeAutospacing = beforeAutospacing;
  if (afterAutospacing !== null) spacing.afterAutospacing = afterAutospacing;
  return Object.keys(spacing).length === 0 ? null : spacing;
};

type NumberingProjection = { readonly numId: number | null; readonly level: number | null };

const numberingValue = (value: FolioContentPropertyValue | null): NumberingProjection => {
  if (value === null) return { numId: null, level: null };
  if (typeof value !== "object" || value.type !== "object") {
    return panic("A canonical DOCX numbering property is not an object");
  }
  const numId = numberValue(
    value.entries.find(({ key }) => key === "numId")?.value ?? null,
    "numPr.numId",
  );
  const level = numberValue(
    value.entries.find(({ key }) => key === "ilvl")?.value ?? null,
    "numPr.ilvl",
  );
  return { numId, level };
};

/** Whether one authored delta has an exact operation-vocabulary lowering. */
export const docxParagraphPropertyChangeIsLowerable = (
  change: FolioContentPropertyChange,
): boolean => {
  if (!Object.hasOwn(DOCX_PARAGRAPH_TRANSPORT_DISPOSITIONS, change.key)) return false;
  const disposition = Reflect.get(DOCX_PARAGRAPH_TRANSPORT_DISPOSITIONS, change.key);
  if (disposition !== "list-level") return disposition !== "unsupported";
  const before = numberingValue(
    change.base.type === "present" ? change.base.value : null,
  );
  const revised = numberingValue(
    change.revised.type === "present" ? change.revised.value : null,
  );
  // The operation can move a paragraph within its current numbering
  // definition, or remove numbering. It cannot switch definitions.
  return revised.numId === null || before.numId === revised.numId;
};

const targetValue = (change: FolioContentPropertyChange): FolioContentPropertyValue | null =>
  change.revised.type === "present" ? change.revised.value : null;

/** Lower only changes proven representable by the paragraph operation vocabulary. */
export const docxParagraphPropertiesFromChanges = (
  changes: readonly FolioContentPropertyChange[],
): FolioAIBlockParagraphProperties => {
  const lowerable = changes.filter(docxParagraphPropertyChangeIsLowerable);
  const properties: FolioAIBlockParagraphProperties = {};
  const spacingChanges: FolioContentPropertySet[number][] = [];
  for (const change of lowerable) {
    const disposition = Reflect.get(DOCX_PARAGRAPH_TRANSPORT_DISPOSITIONS, change.key);
    switch (disposition) {
      case "style":
        properties.styleId = stringValue(targetValue(change), change.key);
        break;
      case "alignment":
        properties.alignment = alignmentValue(targetValue(change));
        break;
      case "spacing":
        if (change.revised.type === "present") {
          spacingChanges.push({ key: change.key, value: change.revised.value });
        }
        break;
      case "list-level":
        properties.listLevel = numberingValue(targetValue(change)).level;
        break;
      case "unsupported":
        return panic("An unsupported paragraph property passed its lowering filter", {
          property: change.key,
        });
      default:
        return panic("A canonical paragraph change has no transport disposition", {
          property: change.key,
        });
    }
  }
  if (
    lowerable.some(
      ({ key }) => Reflect.get(DOCX_PARAGRAPH_TRANSPORT_DISPOSITIONS, key) === "spacing",
    )
  ) {
    properties.spacing = spacingChanges.length === 0 ? null : spacingFromProperties(spacingChanges);
  }
  return properties;
};

/** Complete representable target paragraph state for one transport instruction. */
export const docxParagraphPropertiesFromBlock = (
  block: FolioContentBlock,
): FolioAIBlockParagraphProperties => {
  const authored = block.paragraphFormatting.authored;
  const styleId = stringValue(propertyValue(authored, "styleId"), "styleId");
  const alignment = alignmentValue(propertyValue(authored, "alignment"));
  const spacing = spacingFromProperties(authored);
  const numbering = numberingValue(propertyValue(authored, "numPr"));
  return {
    styleId,
    alignment,
    spacing,
    listLevel: numbering.level,
  };
};

/** Strip neutral structural identities from a consumer-facing cell location. */
export const docxTableLocationFromContent = ({
  outerTableIndex,
  tableIndex,
  rowIndex,
  cellIndex,
  gridColumnIndex,
  columnSpan,
  rowSpan,
  paragraphIndex,
}: FolioContentTableLocation): FolioAIBlockTableLocation => ({
  outerTableIndex,
  tableIndex,
  rowIndex,
  cellIndex,
  gridColumnIndex,
  columnSpan,
  rowSpan,
  paragraphIndex,
});
