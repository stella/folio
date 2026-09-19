// Shapes folio-reserved-values/no-bare-reserved-compare must not flag.
// `scripts/reserved-values-lint.test.ts` lints this file and asserts zero reports.

import type { ParagraphFormatting, TableCellFormatting } from "@stll/docx-core/model";

declare const isNumberingReference: (numId: number | undefined) => boolean;
declare const NO_OUTLINE_LEVEL: number;

// Read through the owning reader.
export const hasNumbering = (formatting: ParagraphFormatting): boolean =>
  isNumberingReference(formatting.numPr?.numId);

// Compared against a named constant rather than the literal.
export const isBodyText = (formatting: ParagraphFormatting): boolean =>
  formatting.outlineLevel === NO_OUTLINE_LEVEL;

// A value the field accepts that means itself.
export const isRestart = (cell: TableCellFormatting): boolean => cell.vMerge === "restart";

// A field with no reserved value.
export const isCentred = (formatting: ParagraphFormatting): boolean =>
  formatting.alignment === "center";

// Same token, different vocabulary: neither name is a model field.
declare const overflowY: string;
declare const operation: { action: string };
export const scrolls = (): boolean => overflowY === "auto";
export const clears = (): boolean => operation.action === "clear";
