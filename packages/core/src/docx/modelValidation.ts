import { validateDocumentModel } from "@stll/docx-core";
import type {
  ValidateDocumentModelIssue,
  ValidateDocumentModelResult,
} from "@stll/docx-core";
import { isCanonicallyEmptyParagraphFormatting } from "@stll/docx-core/model";

import type { Document } from "../types/document";

export class DocxModelValidationError extends Error {
  readonly issues: ValidateDocumentModelIssue[];

  constructor(context: string, issues: ValidateDocumentModelIssue[]) {
    super(`${context}:\n${formatDocumentModelIssues(issues).join("\n")}`);
    this.name = "DocxModelValidationError";
    this.issues = issues;
  }
}

export const validateFolioDocumentModel = (
  document: Document,
): ValidateDocumentModelResult => {
  const issues: ValidateDocumentModelIssue[] = [];
  try {
    issues.push(...validateDocumentModel(document).issues);
  } catch {
    issues.push({
      path: "package.document",
      message: "Document structure is malformed.",
      severity: "error",
    });
  }

  const isRecord = (value: unknown): value is Record<string, unknown> =>
    typeof value === "object" && value !== null && !Array.isArray(value);
  const addMalformedIssue = (path: string, expected: string): void => {
    issues.push({ path, message: `Expected ${expected}.`, severity: "error" });
  };

  const input: unknown = document;
  if (!isRecord(input)) {
    addMalformedIssue("document", "an object");
    return { valid: false, issues };
  }
  const documentPackage = input["package"];
  if (!isRecord(documentPackage)) {
    addMalformedIssue("package", "an object");
    return { valid: false, issues };
  }
  const documentBody = documentPackage["document"];
  if (!isRecord(documentBody)) {
    addMalformedIssue("package.document", "an object");
    return { valid: false, issues };
  }

  if (
    Array.isArray(documentBody["comments"]) &&
    documentBody["comments"].length === 0
  ) {
    issues.push({
      path: "package.document.comments",
      message: "Empty comment lists must be absent.",
      severity: "error",
    });
  }
  const validateBlocks = (blocks: unknown, path: string): void => {
    if (!Array.isArray(blocks)) {
      addMalformedIssue(path, "an array");
      return;
    }
    for (const [index, block] of blocks.entries()) {
      const blockPath = `${path}[${String(index)}]`;
      if (!isRecord(block)) {
        addMalformedIssue(blockPath, "an object");
        continue;
      }
      if (block["type"] === "paragraph") {
        if (
          block["formatting"] !== undefined &&
          !isRecord(block["formatting"])
        ) {
          addMalformedIssue(`${blockPath}.formatting`, "an object");
        } else if (
          block["formatting"] !== undefined &&
          isRecord(block["formatting"]) &&
          isCanonicallyEmptyParagraphFormatting(block["formatting"])
        ) {
          issues.push({
            path: `${blockPath}.formatting`,
            message: "Empty paragraph formatting must be absent.",
            severity: "error",
          });
        }
        continue;
      }
      if (block["type"] === "table") {
        if (!Array.isArray(block["rows"])) {
          addMalformedIssue(`${blockPath}.rows`, "an array");
          continue;
        }
        for (const [rowIndex, row] of block["rows"].entries()) {
          const rowPath = `${blockPath}.rows[${String(rowIndex)}]`;
          if (!isRecord(row)) {
            addMalformedIssue(rowPath, "an object");
            continue;
          }
          if (!Array.isArray(row["cells"])) {
            addMalformedIssue(`${rowPath}.cells`, "an array");
            continue;
          }
          for (const [cellIndex, cell] of row["cells"].entries()) {
            const cellPath = `${rowPath}.cells[${String(cellIndex)}]`;
            if (!isRecord(cell)) {
              addMalformedIssue(cellPath, "an object");
              continue;
            }
            validateBlocks(cell["content"], `${cellPath}.content`);
          }
        }
      } else if (
        block["type"] === "blockSdt" ||
        block["type"] === "blockCustomXml"
      ) {
        validateBlocks(block["content"], `${blockPath}.content`);
      }
    }
  };
  validateBlocks(documentBody["content"], "package.document.content");
  const validateStoryMap = (stories: unknown, path: string): void => {
    if (stories === undefined) return;
    if (!(stories instanceof Map)) {
      addMalformedIssue(path, "a map");
      return;
    }
    for (const [id, story] of stories) {
      const storyPath = `${path}[${JSON.stringify(id)}]`;
      if (!isRecord(story)) {
        addMalformedIssue(storyPath, "an object");
        continue;
      }
      validateBlocks(story["content"], `${storyPath}.content`);
    }
  };
  const validateNoteStories = (notes: unknown, path: string): void => {
    if (notes === undefined) return;
    if (!Array.isArray(notes)) {
      addMalformedIssue(path, "an array");
      return;
    }
    for (const [index, note] of notes.entries()) {
      const notePath = `${path}[${String(index)}]`;
      if (!isRecord(note)) {
        addMalformedIssue(notePath, "an object");
        continue;
      }
      validateBlocks(note["content"], `${notePath}.content`);
    }
  };
  validateStoryMap(documentPackage["headers"], "package.headers");
  validateStoryMap(documentPackage["footers"], "package.footers");
  validateNoteStories(documentPackage["footnotes"], "package.footnotes");
  validateNoteStories(documentPackage["endnotes"], "package.endnotes");
  if (Array.isArray(documentBody["comments"])) {
    for (const [index, comment] of documentBody["comments"].entries()) {
      const commentPath = `package.document.comments[${String(index)}]`;
      if (!isRecord(comment)) {
        addMalformedIssue(commentPath, "an object");
        continue;
      }
      validateBlocks(comment["content"], `${commentPath}.content`);
    }
  }
  return {
    valid: issues.every(({ severity }) => severity !== "error"),
    issues,
  };
};

export const assertValidFolioDocumentModel = (
  document: Document,
  context: string,
): void => {
  const validation = validateFolioDocumentModel(document);
  if (validation.valid) {
    return;
  }

  throw new DocxModelValidationError(context, validation.issues);
};

export const formatDocumentModelIssues = (
  issues: ValidateDocumentModelIssue[],
): string[] =>
  issues.map((issue) => {
    const prefix =
      issue.severity === "warning" ? "DOCX model warning" : "DOCX model error";
    return `${prefix} at ${issue.path}: ${issue.message}`;
  });
