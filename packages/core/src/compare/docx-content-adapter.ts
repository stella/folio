import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import { folioAIBlockIdStability } from "../ai-edits/block-identity";
import type { FolioAIBlock, FolioAIBlockPreviewRun } from "../ai-edits/types";
import type { TextFormatting } from "../types/document";
import type {
  FolioContentInputBlock,
  FolioContentInputRun,
  FolioContentPropertyInput,
  FolioContentPropertyInputValue,
} from "./content-types";

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

const selectedProperties = (
  block: FolioAIBlock,
  disposition: "block-property" | "paragraph-format",
): FolioContentPropertyInput => {
  const properties: { key: string; value: FolioContentPropertyInputValue }[] = [];
  for (const [field, fieldDisposition] of Object.entries(DOCX_BLOCK_FIELD_DISPOSITIONS)) {
    if (fieldDisposition !== disposition) continue;
    const value = Reflect.get(block, field);
    if (value !== undefined) properties.push({ key: field, value: modelValueToProperty(value) });
  }
  return canonicalPropertyInput(properties);
};

/** Convert the DOCX reviewed-view projection into the neutral comparison contract. @internal */
export const docxBlockToContentInput = (block: FolioAIBlock): FolioContentInputBlock => ({
  identity: {
    type: folioAIBlockIdStability(block) === "stable" ? "persistent-hint" : "positional",
    id: block.id,
  },
  kind: block.kind,
  text: block.text,
  blockProperties: selectedProperties(block, "block-property"),
  paragraphFormatting: selectedProperties(block, "paragraph-format"),
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
