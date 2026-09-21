// Deliberate violations of folio-reserved-values/no-bare-reserved-compare.
// `scripts/reserved-values-lint.test.ts` lints this file and asserts the count.

import type { ParagraphNumberingSlots, TableCellFormatting } from "@stll/docx-core/model";

// `w:numId` 0 names no numbering definition.
export const hasNumbering = (numbering: ParagraphNumberingSlots): boolean => numbering.numId !== 0;

// A string sentinel on a cell property.
export const isMergeContinuation = (cell: TableCellFormatting): boolean =>
  cell.vMerge === "continue";

// A reserved border token, read through a computed member access.
export const hasVisibleTopBorder = (cell: TableCellFormatting): boolean =>
  cell.borders?.top?.["style"] !== "nil";
