import { validateDocumentModel } from "@stll/docx-core";
import type { ValidateDocumentModelIssue, ValidateDocumentModelResult } from "@stll/docx-core";

import type { Document } from "../types/document";
import type { BlockContent } from "../types/document";

export class DocxModelValidationError extends Error {
  readonly issues: ValidateDocumentModelIssue[];

  constructor(context: string, issues: ValidateDocumentModelIssue[]) {
    super(`${context}:\n${formatDocumentModelIssues(issues).join("\n")}`);
    this.name = "DocxModelValidationError";
    this.issues = issues;
  }
}

export const validateFolioDocumentModel = (document: Document): ValidateDocumentModelResult => {
  const result = validateDocumentModel(document);
  const issues = [...result.issues];
  if (document.package.document.comments?.length === 0) {
    issues.push({
      path: "package.document.comments",
      message: "Empty comment lists must be absent.",
      severity: "error",
    });
  }
  const validateBlocks = (blocks: readonly BlockContent[], path: string): void => {
    for (const [index, block] of blocks.entries()) {
      const blockPath = `${path}[${String(index)}]`;
      if (block.type === "paragraph") {
        if (block.formatting !== undefined && Object.keys(block.formatting).length === 0) {
          issues.push({
            path: `${blockPath}.formatting`,
            message: "Empty paragraph formatting must be absent.",
            severity: "error",
          });
        }
        continue;
      }
      if (block.type === "table") {
        for (const [rowIndex, row] of block.rows.entries())
          for (const [cellIndex, cell] of row.cells.entries())
            validateBlocks(
              cell.content,
              `${blockPath}.rows[${String(rowIndex)}].cells[${String(cellIndex)}].content`,
            );
      } else if (block.type === "blockSdt") {
        validateBlocks(block.content, `${blockPath}.content`);
      }
    }
  };
  validateBlocks(document.package.document.content, "package.document.content");
  return { valid: issues.every(({ severity }) => severity !== "error"), issues };
};

export const assertValidFolioDocumentModel = (document: Document, context: string): void => {
  const validation = validateFolioDocumentModel(document);
  if (validation.valid) {
    return;
  }

  throw new DocxModelValidationError(context, validation.issues);
};

export const formatDocumentModelIssues = (issues: ValidateDocumentModelIssue[]): string[] =>
  issues.map((issue) => {
    const prefix = issue.severity === "warning" ? "DOCX model warning" : "DOCX model error";
    return `${prefix} at ${issue.path}: ${issue.message}`;
  });
