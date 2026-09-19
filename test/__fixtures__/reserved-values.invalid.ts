// Deliberate violations of folio-reserved-values/no-bare-reserved-compare.
// `scripts/reserved-values-lint.test.ts` lints this file and asserts the count.

import type { ParagraphFormatting, TableCellFormatting } from "@stll/docx-core/model";

// A numeric sentinel compared directly.
export const isBodyText = (formatting: ParagraphFormatting): boolean =>
  formatting.outlineLevel === 9;

// The same decision spelled as a bound.
export const isHeading = (formatting: ParagraphFormatting): boolean =>
  formatting.outlineLevel !== undefined && formatting.outlineLevel <= 8;

// `w:numId` 0 names no numbering definition.
export const hasNumbering = (formatting: ParagraphFormatting): boolean =>
  formatting.numPr?.numId !== 0;

// A string sentinel on a cell property.
export const isMergeContinuation = (cell: TableCellFormatting): boolean =>
  cell.vMerge === "continue";

// A reserved border token, read through a computed member access.
export const hasVisibleTopBorder = (cell: TableCellFormatting): boolean =>
  cell.borders?.top?.["style"] !== "nil";
