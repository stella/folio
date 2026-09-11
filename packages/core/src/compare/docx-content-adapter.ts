import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import { folioAIBlockIdStability } from "../ai-edits/block-identity";
import type {
  FolioAIBlock,
  FolioAIBlockParagraphProperties,
  FolioAIBlockPreviewRun,
  FolioAIBlockTableLocation,
  FolioAIParagraphSpacing,
} from "../ai-edits/types";
import type { ParagraphAlignment, TextFormatting } from "../types/document";
import type {
  FolioContentBlock,
  FolioContentInputBlock,
  FolioContentInputRun,
  FolioContentPropertyInput,
  FolioContentPropertyInputValue,
  FolioContentPropertyChange,
  FolioContentPropertySet,
  FolioContentPropertyValue,
  FolioContentTableLocation,
} from "./content-types";
import type { DocxAuthoredRun } from "./docx-operation-plan";

type DocxBlockFieldDisposition =
  | "identity"
  | "kind"
  | "text"
  | "block-property"
  | "paragraph-format"
  | "runs"
  | "boundaries"
  | "table"
  | "container";

const DOCX_BLOCK_FIELD_DISPOSITIONS = Object.freeze({
  id: "identity",
  idStability: "identity",
  kind: "kind",
  text: "text",
  headingLevel: "block-property",
  displayLabel: "block-property",
  styleId: "paragraph-format",
  directAlignment: "paragraph-format",
  directSpacing: "paragraph-format",
  listLevel: "paragraph-format",
  previewRuns: "runs",
  structuralBoundaries: "boundaries",
  table: "table",
  containerPath: "container",
} as const satisfies Record<keyof FolioAIBlock, DocxBlockFieldDisposition>);

type FieldsWithDisposition<Disposition extends DocxBlockFieldDisposition> = {
  [Field in keyof typeof DOCX_BLOCK_FIELD_DISPOSITIONS]: (typeof DOCX_BLOCK_FIELD_DISPOSITIONS)[Field] extends Disposition
    ? Field
    : never;
}[keyof typeof DOCX_BLOCK_FIELD_DISPOSITIONS];

type DocxParagraphField = FieldsWithDisposition<"paragraph-format">;

const canonicalPropertyInput = (
  entries: readonly { readonly key: string; readonly value: FolioContentPropertyInputValue }[],
): FolioContentPropertyInput =>
  entries.toSorted((left, right) => (left.key < right.key ? -1 : left.key > right.key ? 1 : 0));

const modelValueToProperty = (value: unknown): FolioContentPropertyInputValue => {
  if (value === null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") return Object.is(value, -0) ? 0 : value;
  if (Array.isArray(value)) {
    return { type: "array", items: value.map(modelValueToProperty) };
  }
  if (typeof value !== "object") {
    return panic("A modeled DOCX formatting value is not representable as neutral data");
  }
  const entries: { key: string; value: FolioContentPropertyInputValue }[] = [];
  for (const [key, child] of Object.entries(value)) {
    if (child !== undefined) entries.push({ key, value: modelValueToProperty(child) });
  }
  return { type: "object", entries: canonicalPropertyInput(entries) };
};

const stringProperty = (value: FolioContentPropertyValue | null, field: string): string | null => {
  if (value === null || typeof value === "string") return value;
  return panic("A captured DOCX paragraph property has the wrong value type", { field });
};

const numberProperty = (value: FolioContentPropertyValue | null, field: string): number | null => {
  if (value === null || typeof value === "number") return value;
  return panic("A captured DOCX paragraph property has the wrong value type", { field });
};

const booleanProperty = (value: FolioContentPropertyValue, field: string): boolean => {
  if (typeof value === "boolean") return value;
  return panic("A captured DOCX paragraph property has the wrong value type", { field });
};

type SpacingCodec<Field extends keyof FolioAIParagraphSpacing> = {
  readonly field: Field;
  readonly apply: (spacing: FolioAIParagraphSpacing, value: FolioContentPropertyValue) => void;
};

type SpacingCodecs = {
  readonly [Field in keyof FolioAIParagraphSpacing]: SpacingCodec<Field>;
};

const DOCX_SPACING_CODECS = Object.freeze({
  spaceBefore: Object.freeze({
    field: "spaceBefore",
    apply: (spacing, value) => {
      const decoded = numberProperty(value, "directSpacing.spaceBefore");
      if (decoded === null) return panic("Paragraph spacing cannot contain null");
      spacing.spaceBefore = decoded;
    },
  }),
  spaceAfter: Object.freeze({
    field: "spaceAfter",
    apply: (spacing, value) => {
      const decoded = numberProperty(value, "directSpacing.spaceAfter");
      if (decoded === null) return panic("Paragraph spacing cannot contain null");
      spacing.spaceAfter = decoded;
    },
  }),
  lineSpacing: Object.freeze({
    field: "lineSpacing",
    apply: (spacing, value) => {
      const decoded = numberProperty(value, "directSpacing.lineSpacing");
      if (decoded === null) return panic("Paragraph spacing cannot contain null");
      spacing.lineSpacing = decoded;
    },
  }),
  lineSpacingRule: Object.freeze({
    field: "lineSpacingRule",
    apply: (spacing, value) => {
      const decoded = stringProperty(value, "directSpacing.lineSpacingRule");
      switch (decoded) {
        case "auto":
        case "exact":
        case "atLeast":
          spacing.lineSpacingRule = decoded;
          return;
        default:
          return panic("Captured paragraph spacing has an invalid line-spacing rule", {
            rule: decoded,
          });
      }
    },
  }),
  beforeAutospacing: Object.freeze({
    field: "beforeAutospacing",
    apply: (spacing, value) => {
      spacing.beforeAutospacing = booleanProperty(value, "directSpacing.beforeAutospacing");
    },
  }),
  afterAutospacing: Object.freeze({
    field: "afterAutospacing",
    apply: (spacing, value) => {
      spacing.afterAutospacing = booleanProperty(value, "directSpacing.afterAutospacing");
    },
  }),
} as const satisfies SpacingCodecs);

const spacingCodecByField = new Map<string, SpacingCodec<keyof FolioAIParagraphSpacing>>(
  Object.values(DOCX_SPACING_CODECS).map((codec) => [codec.field, codec]),
);

const spacingProperty = (
  value: FolioContentPropertyValue | null,
): FolioAIBlockParagraphProperties["spacing"] => {
  if (value === null) return null;
  if (typeof value !== "object" || value.type !== "object") {
    return panic("Captured paragraph spacing is not an object");
  }
  const spacing: FolioAIParagraphSpacing = {};
  for (const entry of value.entries) {
    const codec = spacingCodecByField.get(entry.key);
    if (!codec) {
      return panic("Captured paragraph spacing contains an unsupported field", {
        field: entry.key,
      });
    }
    codec.apply(spacing, entry.value);
  }
  return spacing;
};

const paragraphAlignmentProperty = (
  value: FolioContentPropertyValue | null,
): ParagraphAlignment | null => {
  const decoded = stringProperty(value, "directAlignment");
  switch (decoded) {
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
      return decoded;
    default:
      return panic("Captured paragraph alignment has an invalid value", { alignment: decoded });
  }
};

type ParagraphCodec<Field extends DocxParagraphField> = {
  readonly field: Field;
  readonly read: (block: FolioAIBlock) => unknown;
  readonly apply: (
    properties: FolioAIBlockParagraphProperties,
    value: FolioContentPropertyValue | null,
  ) => void;
};

type ParagraphCodecs = {
  readonly [Field in DocxParagraphField]: ParagraphCodec<Field>;
};

const DOCX_PARAGRAPH_CODECS = Object.freeze({
  styleId: Object.freeze({
    field: "styleId",
    read: (block) => block.styleId,
    apply: (properties, value) => {
      properties.styleId = stringProperty(value, "styleId");
    },
  }),
  directAlignment: Object.freeze({
    field: "directAlignment",
    read: (block) => block.directAlignment,
    apply: (properties, value) => {
      properties.alignment = paragraphAlignmentProperty(value);
    },
  }),
  directSpacing: Object.freeze({
    field: "directSpacing",
    read: (block) => block.directSpacing,
    apply: (properties, value) => {
      properties.spacing = spacingProperty(value);
    },
  }),
  listLevel: Object.freeze({
    field: "listLevel",
    read: (block) => block.listLevel,
    apply: (properties, value) => {
      properties.listLevel = numberProperty(value, "listLevel");
    },
  }),
} as const satisfies ParagraphCodecs);

const paragraphCodecByField = new Map<string, ParagraphCodec<DocxParagraphField>>(
  Object.values(DOCX_PARAGRAPH_CODECS).map((codec) => [codec.field, codec]),
);

const textFormattingProperties = (
  formatting: TextFormatting | undefined,
): FolioContentPropertyInput => {
  if (formatting === undefined) return [];
  const properties: { key: string; value: FolioContentPropertyInputValue }[] = [];
  for (const { field } of Object.values(TEXT_FORMATTING_PROPERTY_DESCRIPTORS)) {
    const value = formatting[field];
    if (value !== undefined) properties.push({ key: field, value: modelValueToProperty(value) });
  }
  return canonicalPropertyInput(properties);
};

const previewRunToContentRun = (run: FolioAIBlockPreviewRun): FolioContentInputRun => ({
  text: run.text,
  effectiveFormatting: textFormattingProperties(run.effectiveFormatting),
  authoredFormatting: textFormattingProperties(run.authoredFormatting),
});

type DocxBlockPropertyField = FieldsWithDisposition<"block-property">;

type BlockPropertyCodec<Field extends DocxBlockPropertyField> = {
  readonly field: Field;
  readonly read: (block: FolioAIBlock) => unknown;
};

type BlockPropertyCodecs = {
  readonly [Field in DocxBlockPropertyField]: BlockPropertyCodec<Field>;
};

const DOCX_BLOCK_PROPERTY_CODECS = Object.freeze({
  headingLevel: Object.freeze({
    field: "headingLevel",
    read: (block) => block.headingLevel,
  }),
  displayLabel: Object.freeze({
    field: "displayLabel",
    read: (block) => block.displayLabel,
  }),
} as const satisfies BlockPropertyCodecs);

const blockProperties = (block: FolioAIBlock): FolioContentPropertyInput => {
  const properties: { key: string; value: FolioContentPropertyInputValue }[] = [];
  for (const codec of Object.values(DOCX_BLOCK_PROPERTY_CODECS)) {
    const value = codec.read(block);
    if (value !== undefined) {
      properties.push({ key: codec.field, value: modelValueToProperty(value) });
    }
  }
  return canonicalPropertyInput(properties);
};

const paragraphProperties = (block: FolioAIBlock): FolioContentPropertyInput => {
  const properties: { key: string; value: FolioContentPropertyInputValue }[] = [];
  for (const codec of Object.values(DOCX_PARAGRAPH_CODECS)) {
    const value = codec.read(block);
    if (value !== undefined) {
      properties.push({ key: codec.field, value: modelValueToProperty(value) });
    }
  }
  return canonicalPropertyInput(properties);
};

/** Lower canonical paragraph deltas through the same codecs that captured them. @internal */
export const docxParagraphPropertiesFromChanges = (
  changes: readonly FolioContentPropertyChange[],
): FolioAIBlockParagraphProperties => {
  const properties: FolioAIBlockParagraphProperties = {};
  for (const change of changes) {
    const codec = paragraphCodecByField.get(change.key);
    if (!codec) {
      return panic("The DOCX adapter received an unsupported paragraph property", {
        property: change.key,
      });
    }
    codec.apply(
      properties,
      change.revised.type === "present" ? change.revised.value : null,
    );
  }
  return properties;
};

/** Lower one canonical target block's complete authored paragraph state. @internal */
export const docxParagraphPropertiesFromBlock = (
  block: FolioContentBlock,
): FolioAIBlockParagraphProperties => {
  const properties: FolioAIBlockParagraphProperties = {};
  for (const codec of Object.values(DOCX_PARAGRAPH_CODECS)) {
    const property = block.paragraphFormatting.authored.find(({ key }) => key === codec.field);
    codec.apply(properties, property?.value ?? null);
  }
  return properties;
};

/** Strip neutral structural identities from a consumer-facing cell location. @internal */
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

const ownedTextFormatting = (formatting: TextFormatting | undefined): Readonly<TextFormatting> =>
  Object.freeze(structuredClone(formatting ?? {}));

/** Map canonical target offsets to their authored DOCX run properties without re-diffing. @internal */
export const docxAuthoredRunsForBlock = (
  contentBlock: FolioContentBlock,
  docxBlock: FolioAIBlock,
): readonly DocxAuthoredRun[] => {
  if (
    contentBlock.identity.id !== docxBlock.id ||
    contentBlock.text !== docxBlock.text
  ) {
    return panic("A canonical block no longer names its DOCX source projection", {
      contentBlockId: contentBlock.identity.id,
      docxBlockId: docxBlock.id,
    });
  }
  const runs = docxBlock.previewRuns;
  if (!runs || runs.length === 0) {
    return Object.freeze([
      Object.freeze({
        startOffset: 0,
        endOffset: contentBlock.text.length,
        formatting: ownedTextFormatting(undefined),
      }),
    ]);
  }
  const captured: DocxAuthoredRun[] = [];
  let offset = 0;
  for (const run of runs) {
    const endOffset = offset + run.text.length;
    if (!contentBlock.text.startsWith(run.text, offset)) {
      return panic("DOCX run boundaries do not reconstruct their canonical block", {
        blockId: docxBlock.id,
        offset,
      });
    }
    captured.push(
      Object.freeze({
        startOffset: offset,
        endOffset,
        formatting: ownedTextFormatting(run.authoredFormatting),
      }),
    );
    offset = endOffset;
  }
  if (offset !== contentBlock.text.length) {
    return panic("DOCX run boundaries do not cover their canonical block", {
      blockId: docxBlock.id,
      expected: contentBlock.text.length,
      actual: offset,
    });
  }
  return Object.freeze(captured);
};

/** Own the authored DOCX run projection for one canonical UTF-16 range. @internal */
export const docxAuthoredRunsForRange = (
  contentBlock: FolioContentBlock,
  docxBlock: FolioAIBlock,
  startOffset: number,
  endOffset: number,
): readonly DocxAuthoredRun[] => {
  if (
    !Number.isSafeInteger(startOffset) ||
    !Number.isSafeInteger(endOffset) ||
    startOffset < 0 ||
    endOffset < startOffset ||
    endOffset > contentBlock.text.length
  ) {
    return panic("A DOCX authored-run projection received an invalid canonical range", {
      blockId: contentBlock.identity.id,
      startOffset,
      endOffset,
    });
  }
  if (startOffset === endOffset) {
    return Object.freeze([
      Object.freeze({
        startOffset: 0,
        endOffset: 0,
        formatting: docxAuthoredFormattingAt(contentBlock, docxBlock, startOffset),
      }),
    ]);
  }
  const projected: DocxAuthoredRun[] = [];
  for (const run of docxAuthoredRunsForBlock(contentBlock, docxBlock)) {
    const overlapStart = Math.max(startOffset, run.startOffset);
    const overlapEnd = Math.min(endOffset, run.endOffset);
    if (overlapStart >= overlapEnd) continue;
    projected.push(
      Object.freeze({
        startOffset: overlapStart - startOffset,
        endOffset: overlapEnd - startOffset,
        formatting: run.formatting,
      }),
    );
  }
  if (projected.at(-1)?.endOffset !== endOffset - startOffset) {
    return panic("A DOCX authored-run range does not reconstruct its canonical text", {
      blockId: contentBlock.identity.id,
      startOffset,
      endOffset,
    });
  }
  return Object.freeze(projected);
};

/** Authored DOCX formatting at one canonical target offset. @internal */
export const docxAuthoredFormattingAt = (
  contentBlock: FolioContentBlock,
  docxBlock: FolioAIBlock,
  offset: number,
): Readonly<TextFormatting> => {
  const runs = docxAuthoredRunsForBlock(contentBlock, docxBlock);
  const run = runs.find(
    ({ startOffset, endOffset }) =>
      (offset >= startOffset && offset < endOffset) ||
      (offset === contentBlock.text.length && endOffset === offset),
  );
  return run?.formatting ?? ownedTextFormatting(undefined);
};

/** Convert the DOCX reviewed-view projection into the neutral comparison contract. @internal */
export const docxBlockToContentInput = (block: FolioAIBlock): FolioContentInputBlock => ({
  identity: {
    type: folioAIBlockIdStability(block) === "stable" ? "persistent-hint" : "positional",
    id: block.id,
  },
  kind: block.kind,
  text: block.text,
  blockProperties: blockProperties(block),
  paragraphFormatting: {
    authored: paragraphProperties(block),
    // The current DOCX snapshot exposes authored paragraph properties. It does
    // not claim style-resolved values as authored presentation.
    effective: [],
  },
  runs: block.previewRuns?.map(previewRunToContentRun) ?? [],
  structuralBoundaries: block.structuralBoundaries ?? [],
  ...(block.table !== undefined && {
    table: {
      outerTableIdentity: {
        type: "positional",
        id: `outer-table-${String(block.table.outerTableIndex)}`,
      },
      tableIdentity: { type: "positional", id: `table-${String(block.table.tableIndex)}` },
      rowIdentity: {
        type: "positional",
        id: `table-${String(block.table.tableIndex)}-row-${String(block.table.rowIndex)}`,
      },
      cellIdentity: {
        type: "positional",
        id: `table-${String(block.table.tableIndex)}-row-${String(block.table.rowIndex)}-cell-${String(block.table.cellIndex)}`,
      },
      ...block.table,
    },
  }),
  containerPath:
    block.containerPath?.map(({ kind, id }) => ({
      kind,
      identity: { type: "persistent-hint", id },
    })) ?? [],
});
