import type { Mark, Node as PMNode } from "prosemirror-model";

import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import { readBookmarkBoundaryAttrs } from "./bookmarkBoundaryAttrs";
import {
  readCharacterSpacingMarkAttrs,
  readCharacterStyleMarkAttrs,
  readCommentMarkAttrs,
  readEmphasisMarkAttrs,
  readFieldAttrs,
  readFontFamilyMarkAttrs,
  readFontSizeMarkAttrs,
  readFootnoteRefMarkAttrs,
  readHardBreakAttrs,
  readHighlightMarkAttrs,
  readHyperlinkMarkAttrs,
  readImageAttrs,
  readLanguageMarkAttrs,
  readMathAttrs,
  readBlockSdtAttrs,
  readParagraphAttrs,
  readRunFormattingOverrideMarkAttrs,
  readRunPropertyChangeMarkAttrs,
  readRunShadingMarkAttrs,
  readSdtAttrs,
  readShapeAttrs,
  readStrikeMarkAttrs,
  readSymbolAttrs,
  readTabAttrs,
  readTableAttrs,
  readTableCellAttrs,
  readTableRowAttrs,
  readTextBoxAttrs,
  readTextColorMarkAttrs,
  readTextEffectMarkAttrs,
  readTrackedChangeMarkAttrs,
  readUnderlineMarkAttrs,
} from "./attrs";
import { readTextBoxAnchorAttrs } from "./textBoxAnchorAttrs";

export type ProseMirrorDocumentValidationIssue = {
  path: string;
  message: string;
};

export type ValidateProseMirrorDocumentResult = {
  valid: boolean;
  issues: ProseMirrorDocumentValidationIssue[];
};

export class ProseMirrorDocumentValidationError extends Error {
  readonly issues: ProseMirrorDocumentValidationIssue[];

  constructor(context: string, issues: ProseMirrorDocumentValidationIssue[]) {
    super(`${context}:\n${formatProseMirrorDocumentIssues(issues).join("\n")}`);
    this.name = "ProseMirrorDocumentValidationError";
    this.issues = issues;
  }
}

const validDocumentCache = new WeakSet<PMNode>();
const validNodeCache = new WeakSet<PMNode>();

export const validateProseMirrorDocument = (doc: PMNode): ValidateProseMirrorDocumentResult => {
  const issues: ProseMirrorDocumentValidationIssue[] = [];

  if (doc.type.name !== "doc") {
    issues.push({
      path: "doc.type.name",
      message: `Expected doc, got ${doc.type.name}.`,
    });
  }

  validateNode(doc, "doc", issues);
  validateBookmarkBoundaryStructure(doc, issues);

  return {
    valid: issues.length === 0,
    issues,
  };
};

type OpenBookmarkBoundary = {
  id: number;
  path: string;
};

const validateBookmarkBoundaryStructure = (
  doc: PMNode,
  issues: ProseMirrorDocumentValidationIssue[],
): void => {
  // Bookmark boundaries pair by id, not by stack order. Word permits ranges
  // to overlap (start A, start B, end A, end B), so crossing pairs are valid.
  const open = new Map<number, OpenBookmarkBoundary>();
  const startedIds = new Set<number>();

  const registerStart = (id: number, path: string): void => {
    if (startedIds.has(id)) {
      issues.push({ path, message: `Bookmark id ${id} has more than one start boundary.` });
      return;
    }
    startedIds.add(id);
    open.set(id, { id, path });
  };

  const registerEnd = (id: number, path: string): void => {
    const start = open.get(id);
    if (!start) {
      issues.push({ path, message: `Bookmark id ${id} has no open start boundary.` });
      return;
    }
    open.delete(id);
  };

  const visit = (node: PMNode, path: string): void => {
    const paragraphAttrs = node.type.name === "paragraph" ? readParagraphAttrs(node) : null;
    if (paragraphAttrs?.ok) {
      for (const [index, bookmark] of (paragraphAttrs.value.bookmarks ?? []).entries()) {
        registerStart(bookmark.id, `${path}.paragraph.attrs.bookmarks[${index}]`);
      }
    }

    if (node.type.name === "bookmarkBoundary") {
      const result = readBookmarkBoundaryAttrs(node);
      if (result.ok) {
        const attrs = result.value;
        const hasHyperlink = node.marks.some((mark) => mark.type.name === "hyperlink");
        const trackedChanges = node.marks.filter(
          (mark) => mark.type.name === "insertion" || mark.type.name === "deletion",
        );
        if (trackedChanges.length > 1) {
          issues.push({
            path,
            message: "Bookmark boundaries cannot carry multiple tracked-change parents.",
          });
        } else if (trackedChanges.length === 1 && !hasHyperlink) {
          issues.push({
            path,
            message:
              "Bookmark boundaries inside tracked changes require a hyperlink serialization parent.",
          });
        }
        if (attrs.type === "start") {
          registerStart(attrs.id, path);
        } else {
          registerEnd(attrs.id, path);
        }
      }
    }

    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    node.forEach((child, _offset, index) => {
      visit(child, `${path}.content[${index}]`);
    });
    if (paragraphAttrs?.ok) {
      for (const [index, bookmark] of (paragraphAttrs.value.bookmarks ?? []).entries()) {
        registerEnd(bookmark.id, `${path}.paragraph.attrs.bookmarks[${index}]`);
      }
    }
  };

  visit(doc, "doc");
  for (const boundary of open.values()) {
    issues.push({
      path: boundary.path,
      message: `Bookmark id ${boundary.id} has no matching end boundary.`,
    });
  }
};

export const assertValidProseMirrorDocument = (doc: PMNode, context: string): void => {
  if (validDocumentCache.has(doc)) {
    return;
  }

  const validation = validateProseMirrorDocument(doc);
  if (validation.valid) {
    validDocumentCache.add(doc);
    return;
  }

  throw new ProseMirrorDocumentValidationError(context, validation.issues);
};

export const formatProseMirrorDocumentIssues = (
  issues: ProseMirrorDocumentValidationIssue[],
): string[] =>
  issues.map((issue) => `ProseMirror document error at ${issue.path}: ${issue.message}`);

const validateNode = (
  node: PMNode,
  path: string,
  issues: ProseMirrorDocumentValidationIssue[],
): void => {
  if (validNodeCache.has(node)) {
    return;
  }

  const issueCountBeforeNode = issues.length;

  validateNodeAttrs(node, path, issues);
  validateMarks(node.marks, path, issues);

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, _offset, index) => {
    validateNode(child, `${path}.content[${index}]`, issues);
  });

  if (issues.length === issueCountBeforeNode) {
    validNodeCache.add(node);
  }
};

const validateNodeAttrs = (
  node: PMNode,
  path: string,
  issues: ProseMirrorDocumentValidationIssue[],
): void => {
  switch (node.type.name) {
    case "doc":
    case "text":
    case "horizontalRule":
    case "pageBreak":
    case "renderedPageBreak":
      return;

    case "bookmarkBoundary":
      appendAttrIssues(path, readBookmarkBoundaryAttrs(node), issues);
      return;

    case "tab":
      appendAttrIssues(path, readTabAttrs(node), issues);
      return;

    case "symbol":
      appendAttrIssues(path, readSymbolAttrs(node), issues);
      return;

    case "hardBreak":
      appendAttrIssues(path, readHardBreakAttrs(node), issues);
      return;

    case "paragraph":
      appendAttrIssues(path, readParagraphAttrs(node), issues);
      return;

    case "table":
      appendAttrIssues(path, readTableAttrs(node), issues);
      return;

    case "tableRow":
      appendAttrIssues(path, readTableRowAttrs(node), issues);
      return;

    case "tableCell":
    case "tableHeader":
      appendAttrIssues(path, readTableCellAttrs(node), issues);
      return;

    case "image":
      appendAttrIssues(path, readImageAttrs(node), issues);
      return;

    case "field":
      appendAttrIssues(path, readFieldAttrs(node), issues);
      if (node.childCount > 0) {
        issues.push({
          path: `${path}.content`,
          message: "Ordinary fields cannot contain structured result children.",
        });
      }
      return;

    case "structuredField":
      {
        const fieldAttrs = readFieldAttrs(node);
        appendAttrIssues(path, fieldAttrs, issues);
        if (fieldAttrs.ok) {
          const hasStructuredHyperlink = node.content.content.some((child) =>
            child.marks.some((mark) => mark.type.name === "hyperlink"),
          );
          if (fieldAttrs.value.fieldKind === "complex") {
            issues.push({
              path: `${path}.content`,
              message: "Complex fields cannot contain structured result children.",
            });
          } else if (!hasStructuredHyperlink) {
            issues.push({
              path: `${path}.content`,
              message: "Structured simple fields require hyperlink content.",
            });
          }
          // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
          node.forEach((child, _offset, index) => {
            const childPath = `${path}.content[${index}]`;
            const hasHyperlink = child.marks.some((mark) => mark.type.name === "hyperlink");
            if (child.type.name === "bookmarkBoundary" && !hasHyperlink) {
              issues.push({
                path: childPath,
                message: "Bookmark boundaries inside fields require a hyperlink parent.",
              });
            }
          });
        }
      }
      return;

    case "math":
      appendAttrIssues(path, readMathAttrs(node), issues);
      return;

    case "sdt":
      appendAttrIssues(path, readSdtAttrs(node), issues);
      return;

    case "blockSdt":
      // Validate attrs through the typed reader so malformed projections
      // (wrong sdtType, non-string rawPropertiesXml, etc.) surface here
      // instead of leaking into the serializer or downstream consumers.
      appendAttrIssues(path, readBlockSdtAttrs(node), issues);
      for (let i = 0; i < node.childCount; i += 1) {
        validateNode(node.child(i), `${path}.content[${i}]`, issues);
      }
      return;

    case "shape":
      appendAttrIssues(path, readShapeAttrs(node), issues);
      return;

    case "textBox":
      appendAttrIssues(path, readTextBoxAttrs(node), issues);
      return;

    case "textBoxAnchor":
      appendAttrIssues(path, readTextBoxAnchorAttrs(node), issues);
      return;

    default:
      issues.push({
        path: `${path}.type.name`,
        message: `Unsupported ProseMirror node type ${node.type.name}.`,
      });
  }
};

const validateMarks = (
  marks: readonly Mark[],
  path: string,
  issues: ProseMirrorDocumentValidationIssue[],
): void => {
  for (const [index, mark] of marks.entries()) {
    const markPath = `${path}.marks[${index}]`;
    switch (mark.type.name) {
      case "bold":
      case "italic":
      case "subscript":
      case "superscript":
      case "allCaps":
      case "smallCaps":
      case "emboss":
      case "imprint":
      case "hidden":
      case "textShadow":
      case "textOutline":
      case "rtl":
        continue;

      case "textEffect":
        appendAttrIssues(markPath, readTextEffectMarkAttrs(mark), issues);
        continue;

      case "underline":
        appendAttrIssues(markPath, readUnderlineMarkAttrs(mark), issues);
        continue;

      case "strike":
        appendAttrIssues(markPath, readStrikeMarkAttrs(mark), issues);
        continue;

      case "textColor":
        appendAttrIssues(markPath, readTextColorMarkAttrs(mark), issues);
        continue;

      case "highlight":
        appendAttrIssues(markPath, readHighlightMarkAttrs(mark), issues);
        continue;

      case "runShading":
        appendAttrIssues(markPath, readRunShadingMarkAttrs(mark), issues);
        continue;

      case "fontSize":
        appendAttrIssues(markPath, readFontSizeMarkAttrs(mark), issues);
        continue;

      case "fontFamily":
        appendAttrIssues(markPath, readFontFamilyMarkAttrs(mark), issues);
        continue;

      case "language":
        appendAttrIssues(markPath, readLanguageMarkAttrs(mark), issues);
        continue;

      case "characterSpacing":
        appendAttrIssues(markPath, readCharacterSpacingMarkAttrs(mark), issues);
        continue;

      case "characterStyle":
        appendAttrIssues(markPath, readCharacterStyleMarkAttrs(mark), issues);
        continue;

      case "emphasisMark":
        appendAttrIssues(markPath, readEmphasisMarkAttrs(mark), issues);
        continue;

      case "footnoteRef":
        appendAttrIssues(markPath, readFootnoteRefMarkAttrs(mark), issues);
        continue;

      case "comment":
        appendAttrIssues(markPath, readCommentMarkAttrs(mark), issues);
        continue;

      case "insertion":
      case "deletion":
        appendAttrIssues(markPath, readTrackedChangeMarkAttrs(mark), issues);
        continue;

      case "runPropertyChange":
        appendAttrIssues(markPath, readRunPropertyChangeMarkAttrs(mark), issues);
        continue;

      case "runFormattingOverride":
        appendAttrIssues(markPath, readRunFormattingOverrideMarkAttrs(mark), issues);
        continue;

      case "hyperlink":
        appendAttrIssues(markPath, readHyperlinkMarkAttrs(mark), issues);
        continue;

      default:
        issues.push({
          path: `${markPath}.type.name`,
          message: `Unsupported ProseMirror mark type ${mark.type.name}.`,
        });
    }
  }
};

const appendAttrIssues = <T>(
  path: string,
  result: ReadProseMirrorAttrsResult<T>,
  issues: ProseMirrorDocumentValidationIssue[],
): void => {
  if (result.ok) {
    return;
  }

  for (const issue of result.issues) {
    issues.push(withPathPrefix(path, issue));
  }
};

const withPathPrefix = (
  prefix: string,
  issue: ProseMirrorAttrIssue,
): ProseMirrorDocumentValidationIssue => ({
  path: `${prefix}.${issue.path}`,
  message: issue.message,
});
