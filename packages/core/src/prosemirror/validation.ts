import type { Mark, Node as PMNode } from "prosemirror-model";

import { readNoteMarkerAttrs } from "../internal/noteMarkerAttrs";
import { subtreeHolds } from "../internal/subtreeHolds";
import type { PositionedBookmarkMarker } from "../types/document";
import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import {
  readBlockCustomXmlAttrs,
  readBlockSdtAttrs,
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
  readInlineWrapperMarkAttrs,
  readLanguageMarkAttrs,
  readMathAttrs,
  readPageBreakRunAttrs,
  readParagraphAttrs,
  readPreservedBlockAttrs,
  readPreservedXmlAttrs,
  readRunFormattingOverrideMarkAttrs,
  readRunIdentityMarkAttrs,
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
import { readBookmarkBoundaryAttrs } from "./bookmarkBoundaryAttrs";
import { readCommentRangeBoundaryAttrs } from "./commentRangeBoundaryAttrs";
import { readCommentReferenceAttrs } from "./commentReferenceAttrs";
import { HYPHEN_ELEMENT_MARK_NAME } from "./conversion/hyphenTextCarriers";
import { fieldRequiresStructuredContent } from "./fieldRepresentation";
import { readHyphenElementMarkAttrs } from "./hyphenElementAttrs";
import { readMoveRangeBoundaryAttrs } from "./moveRangeBoundaryAttrs";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";
import { isRunContentInlineNodeName } from "./runContentProjection";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { RUN_IDENTITY_MARK_NAME } from "./runIdentity";
import { type ParagraphAttrs, PRESERVED_XML_LEVELS } from "./schema/nodes";
import { readTextBoxAnchorAttrs } from "./textBoxAnchorAttrs";

const COMPLEX_FIELD_RESULT_MARK_NAMES = new Set([
  ...RUN_FORMATTING_MARK_NAMES,
  RUN_IDENTITY_MARK_NAME,
  HYPHEN_ELEMENT_MARK_NAME,
  "runPropertyChange",
  "footnoteRef",
]);

const hasUnsupportedComplexFieldResultContent = (node: PMNode): boolean =>
  node.content.content.some((child) => {
    if (child.type.name === "bookmarkBoundary") {
      return false;
    }
    const preservedXmlAttrs =
      child.type.name === "preservedXml"
        ? readPreservedXmlAttrs(child)
        : undefined;
    const inlinePreservedXml =
      preservedXmlAttrs?.ok &&
      preservedXmlAttrs.value.level === PRESERVED_XML_LEVELS.inline;
    return (
      inlinePreservedXml ||
      !isRunContentInlineNodeName(child.type.name) ||
      child.marks.some(
        (mark) =>
          !COMPLEX_FIELD_RESULT_MARK_NAMES.has(mark.type.name) &&
          // What the whole field carries, its result carries along.
          !mark.isInSet(node.marks),
      )
    );
  });

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

export const validateProseMirrorDocument = (
  doc: PMNode,
): ValidateProseMirrorDocumentResult => {
  const issues: ProseMirrorDocumentValidationIssue[] = [];

  if (doc.type.name !== "doc") {
    issues.push({
      path: "doc.type.name",
      message: `Expected doc, got ${doc.type.name}.`,
    });
  }

  validateNode(doc, "doc", issues);
  validateBookmarkBoundaryStructure(doc, issues, holdsNoBookmarkBoundary);

  return {
    valid: issues.length === 0,
    issues,
  };
};

/**
 * Where a bookmark boundary came from. A boundary is either a node or a
 * positioned marker on a table or row. Positioned starts participate in
 * pairing because their other boundary may be an inline node; unmatched
 * positioned markers are tolerated because imported documents can omit their
 * other carrier. A
 * `paragraph.attrs.bookmarks` entry is input data the conversion could not pair
 * (see `collectPairedBookmarkIds`) and is written back out as a start and an end
 * around that one paragraph. The distinction decides how strictly a duplicate
 * id is treated.
 */
type BookmarkBoundaryOrigin = "boundary" | "paragraph-attr" | "positioned";

type OpenBookmarkBoundary = {
  id: number;
  name: string | null;
  path: string;
  paragraph: string;
  origin: BookmarkBoundaryOrigin;
};

/**
 * Locate a bookmark boundary by the paragraph holding it, so a failure names a
 * place a reader can find rather than only a node path.
 *
 * The paragraph's text is deliberately left out. Validation messages travel
 * into logs and error reports, and document text is not log data.
 */
const describeParagraph = (
  attrs: ReadProseMirrorAttrsResult<ParagraphAttrs> | null,
  path: string,
): string => {
  const paraId = attrs?.ok === true ? attrs.value.paraId : undefined;
  return paraId !== undefined && paraId.length > 0
    ? `paragraph ${paraId}`
    : `paragraph at ${path}`;
};

const describeBookmark = (id: number, name: string | null): string =>
  name === null || name.length === 0
    ? `Bookmark id ${id}`
    : `Bookmark "${name}" (id ${id})`;

// A node that carries a bookmark boundary itself: a boundary node, a
// paragraph's unpaired bookmarks, or a table's or row's positioned markers.
// Attrs that do not read count too, so a subtree the pass skips never holds
// a boundary.
const carriesBookmarkBoundary = (node: PMNode): boolean => {
  switch (node.type.name) {
    case "bookmarkBoundary":
    case "blockBookmarkBoundary":
      return true;
    case "paragraph": {
      const attrs = readParagraphAttrs(node);
      return !attrs.ok || (attrs.value.bookmarks?.length ?? 0) > 0;
    }
    case "table": {
      const attrs = readTableAttrs(node);
      return !attrs.ok || (attrs.value._bookmarks?.length ?? 0) > 0;
    }
    case "tableRow": {
      const attrs = readTableRowAttrs(node);
      return !attrs.ok || (attrs.value._bookmarks?.length ?? 0) > 0;
    }
    default:
      return false;
  }
};

// Nodes are immutable and an edit shares every subtree it does not touch, so
// whether a subtree holds a boundary is decided once per node. The pairing
// pass then walks only the subtrees that hold one, and an edit costs the path
// it changed rather than every node of the document.
const holdsBookmarkBoundary = subtreeHolds(carriesBookmarkBoundary);
const holdsNoBookmarkBoundary = (node: PMNode): boolean =>
  !holdsBookmarkBoundary(node);

/** The pairing pass over every node, skipping none: the reference for the skip. */
export const bookmarkIssuesWalkingEveryNode = (
  doc: PMNode,
): ProseMirrorDocumentValidationIssue[] => {
  const issues: ProseMirrorDocumentValidationIssue[] = [];
  validateBookmarkBoundaryStructure(doc, issues, () => false);
  return issues;
};

/** The pairing pass as validation runs it. */
export const bookmarkIssues = (
  doc: PMNode,
): ProseMirrorDocumentValidationIssue[] => {
  const issues: ProseMirrorDocumentValidationIssue[] = [];
  validateBookmarkBoundaryStructure(doc, issues, holdsNoBookmarkBoundary);
  return issues;
};

const validateBookmarkBoundaryStructure = (
  doc: PMNode,
  issues: ProseMirrorDocumentValidationIssue[],
  skip: (node: PMNode) => boolean,
  /** The indices of the document's children `skip` keeps, when known. */
  holders?: readonly number[],
): void => {
  // Bookmark boundaries pair by id, not by stack order. Word permits ranges
  // to overlap (start A, start B, end A, end B), so crossing pairs are valid.
  //
  // A duplicate id is an error EXCEPT when both occurrences are paragraph-attr
  // bookmarks. Real documents repeat an id that way — a stray `_GoBack`, a
  // template assembled from several sources, a file round-tripped through
  // another editor — and Word and LibreOffice keep the first start and open the
  // file rather than refusing it. Rejecting the document instead made every such
  // file unopenable here, and the attr carries no range to corrupt: it
  // serializes as a start and an end around its own paragraph, so the tolerated
  // occurrence writes back exactly what was read.
  //
  // A duplicate involving a `bookmarkBoundary` node stays an error. Those this
  // codebase emits itself, in pairs, so a collision there means an editing
  // operation (a paste carrying an id the document already uses) produced an
  // ambiguous anchor rather than an input document being untidy.
  const open = new Map<number, OpenBookmarkBoundary>();
  const startedBy = new Map<number, BookmarkBoundaryOrigin>();

  /** Whether the boundary opened, and so needs a matching end registered later. */
  const registerStart = (boundary: OpenBookmarkBoundary): boolean => {
    const { id, name, path, paragraph, origin } = boundary;
    const previousOrigin = startedBy.get(id);
    if (previousOrigin !== undefined) {
      if (origin === "paragraph-attr" && previousOrigin === "paragraph-attr") {
        return false;
      }
      issues.push({
        path,
        message: `${describeBookmark(id, name)} has more than one start boundary (${paragraph}).`,
      });
      return false;
    }
    startedBy.set(id, origin);
    open.set(id, boundary);
    return true;
  };

  const registerEnd = (id: number, path: string, paragraph: string): void => {
    if (!open.has(id)) {
      issues.push({
        path,
        message: `${describeBookmark(id, null)} has no open start boundary (${paragraph}).`,
      });
      return;
    }
    open.delete(id);
  };

  const visit = (node: PMNode, path: string, paragraph: string): void => {
    const paragraphAttrs =
      node.type.name === "paragraph" ? readParagraphAttrs(node) : null;
    const enclosingParagraph =
      node.type.name === "paragraph"
        ? describeParagraph(paragraphAttrs, path)
        : paragraph;

    // Only the attr bookmarks that actually opened are closed after the
    // children. A tolerated duplicate never opened, so registering its end
    // would report an unmatched end that is not there.
    const openedAttrBookmarks: { id: number; path: string }[] = [];
    if (paragraphAttrs?.ok) {
      for (const [index, bookmark] of (
        paragraphAttrs.value.bookmarks ?? []
      ).entries()) {
        const bookmarkPath = `${path}.paragraph.attrs.bookmarks[${index}]`;
        const opened = registerStart({
          id: bookmark.id,
          name: bookmark.name,
          path: bookmarkPath,
          paragraph: enclosingParagraph,
          origin: "paragraph-attr",
        });
        if (opened) {
          openedAttrBookmarks.push({ id: bookmark.id, path: bookmarkPath });
        }
      }
    }

    if (
      node.type.name === "bookmarkBoundary" ||
      node.type.name === "blockBookmarkBoundary"
    ) {
      const result = readBookmarkBoundaryAttrs(node);
      if (result.ok) {
        const attrs = result.value;
        const trackedChanges = node.marks.filter(
          (mark) =>
            mark.type.name === "insertion" || mark.type.name === "deletion",
        );
        if (trackedChanges.length > 1) {
          issues.push({
            path,
            message: `Bookmark boundaries cannot carry multiple tracked-change parents (${enclosingParagraph}).`,
          });
        }
        if (attrs.type === "start") {
          registerStart({
            id: attrs.id,
            name: attrs.name ?? null,
            path,
            paragraph: enclosingParagraph,
            origin: "boundary",
          });
        } else {
          registerEnd(attrs.id, path, enclosingParagraph);
        }
      }
    }

    let positionedBookmarks: readonly PositionedBookmarkMarker[] = [];
    if (node.type.name === "table") {
      const result = readTableAttrs(node);
      if (result.ok) {
        positionedBookmarks = result.value._bookmarks ?? [];
      }
    } else if (node.type.name === "tableRow") {
      const result = readTableRowAttrs(node);
      if (result.ok) {
        positionedBookmarks = result.value._bookmarks ?? [];
      }
    }
    const positionedBookmarksByChildIndex = new Map<
      number,
      { marker: PositionedBookmarkMarker["marker"]; path: string }[]
    >();
    for (const [index, positioned] of positionedBookmarks.entries()) {
      const markerPath = `${path}.${node.type.name}.attrs._bookmarks[${index}]`;
      if (positioned.index > node.childCount) {
        issues.push({
          path: `${markerPath}.index`,
          message: `Expected a child position between 0 and ${node.childCount}.`,
        });
        continue;
      }
      const markers = positionedBookmarksByChildIndex.get(positioned.index);
      const atPosition = {
        marker: positioned.marker,
        path: markerPath,
      };
      if (markers === undefined) {
        positionedBookmarksByChildIndex.set(positioned.index, [atPosition]);
      } else {
        markers.push(atPosition);
      }
    }
    const visitPositionedBookmarks = (childIndex: number): void => {
      for (const {
        marker,
        path: markerPath,
      } of positionedBookmarksByChildIndex.get(childIndex) ?? []) {
        if (marker.type === "bookmarkStart") {
          registerStart({
            id: marker.id,
            name: marker.name,
            path: markerPath,
            paragraph: enclosingParagraph,
            origin: "positioned",
          });
        } else if (open.has(marker.id)) {
          registerEnd(marker.id, markerPath, enclosingParagraph);
        }
      }
    };

    // A table or row marker stands before the child whose index it records.
    // Bucketing once preserves authored order within that position while the
    // child walk stays linear in document-controlled children and markers.
    if (node === doc && holders !== undefined) {
      // The document carries no positioned markers of its own.
      for (const index of holders)
        visit(
          node.child(index),
          `${path}.content[${index}]`,
          enclosingParagraph,
        );
    } else {
      // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
      node.forEach((child, _offset, index) => {
        visitPositionedBookmarks(index);
        if (!skip(child))
          visit(child, `${path}.content[${index}]`, enclosingParagraph);
      });
    }
    visitPositionedBookmarks(node.childCount);

    for (const { id, path: bookmarkPath } of openedAttrBookmarks) {
      registerEnd(id, bookmarkPath, enclosingParagraph);
    }
  };

  visit(doc, "doc", "the document root");
  for (const boundary of open.values()) {
    if (boundary.origin === "positioned") {
      continue;
    }
    issues.push({
      path: boundary.path,
      message: `${describeBookmark(boundary.id, boundary.name)} has no matching end boundary (${boundary.paragraph}).`,
    });
  }
};

export const assertValidProseMirrorDocument = (
  doc: PMNode,
  context: string,
): void => {
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

/** A story document that passed validation, as the next one is checked against it. */
export type ValidatedStory = {
  doc: PMNode;
  /** The indices of its children holding a bookmark boundary. */
  holders: readonly number[];
};

/**
 * {@link assertValidProseMirrorDocument} for a story document, checking only
 * the top-level children that `previous`, the last document that passed,
 * did not hold: the run of children from the start and the run to the end
 * that are the very nodes `previous` held were valid there, and the bookmark
 * pass visits the children holding a boundary, kept ones by their shifted
 * index. Gives what the next document is checked against.
 */
export const assertValidProseMirrorDocumentKeeping = (
  doc: PMNode,
  context: string,
  previous: ValidatedStory | undefined,
): ValidatedStory => {
  const before = previous?.doc;
  const count = doc.childCount;
  const old = before?.childCount ?? 0;
  let start = 0;
  while (
    start < count &&
    start < old &&
    doc.child(start) === before?.child(start)
  )
    start += 1;
  let after = 0;
  while (
    after < count - start &&
    after < old - start &&
    doc.child(count - 1 - after) === before?.child(old - 1 - after)
  )
    after += 1;
  const issues: ProseMirrorDocumentValidationIssue[] = [];
  if (doc.type.name !== "doc") {
    issues.push({
      path: "doc.type.name",
      message: `Expected doc, got ${doc.type.name}.`,
    });
  }
  validateNodeAttrs(doc, "doc", issues);
  validateMarks(doc.marks, "doc", issues);
  for (let index = start; index < count - after; index += 1)
    validateNode(doc.child(index), `doc.content[${index}]`, issues);
  const shift = count - old;
  const holders: number[] = [];
  for (const index of previous?.holders ?? [])
    if (index < start) holders.push(index);
  for (let index = start; index < count - after; index += 1)
    if (holdsBookmarkBoundary(doc.child(index))) holders.push(index);
  for (const index of previous?.holders ?? [])
    if (index >= old - after) holders.push(index + shift);
  validateBookmarkBoundaryStructure(
    doc,
    issues,
    holdsNoBookmarkBoundary,
    holders,
  );
  if (issues.length > 0)
    throw new ProseMirrorDocumentValidationError(context, issues);
  validNodeCache.add(doc);
  validDocumentCache.add(doc);
  return { doc, holders };
};

export const formatProseMirrorDocumentIssues = (
  issues: ProseMirrorDocumentValidationIssue[],
): string[] =>
  issues.map(
    (issue) => `ProseMirror document error at ${issue.path}: ${issue.message}`,
  );

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

    // Either level: the same marker, so the same attribute reader.
    case "bookmarkBoundary":
    case "blockBookmarkBoundary":
      appendAttrIssues(path, readBookmarkBoundaryAttrs(node), issues);
      return;

    case "commentReference":
      appendAttrIssues(path, readCommentReferenceAttrs(node), issues);
      return;

    case "commentRangeBoundary":
    case "blockCommentRangeBoundary":
      appendAttrIssues(path, readCommentRangeBoundaryAttrs(node), issues);
      return;

    case "rangeAnchor":
      appendAttrIssues(path, readRangeAnchorAttrs(node), issues);
      return;

    case "moveRangeBoundary":
    case "blockMoveRangeBoundary":
      appendAttrIssues(path, readMoveRangeBoundaryAttrs(node), issues);
      return;

    case "tab":
      appendAttrIssues(path, readTabAttrs(node), issues);
      return;

    case "symbol":
      appendAttrIssues(path, readSymbolAttrs(node), issues);
      return;

    case "noteMarker":
      appendAttrIssues(path, readNoteMarkerAttrs(node), issues);
      return;

    case "preservedXml":
      appendAttrIssues(path, readPreservedXmlAttrs(node), issues);
      return;

    case "preservedBlock":
      appendAttrIssues(path, readPreservedBlockAttrs(node), issues);
      return;

    case "hardBreak":
      appendAttrIssues(path, readHardBreakAttrs(node), issues);
      return;

    case "pageBreakRun":
      appendAttrIssues(path, readPageBreakRunAttrs(node), issues);
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
          if (!fieldRequiresStructuredContent(node.content.content)) {
            issues.push({
              path: `${path}.content`,
              message:
                "Structured fields require run metadata, distinct run marks or inline structure.",
            });
          } else if (
            fieldAttrs.value.fieldKind === "complex" &&
            hasStructuredHyperlink
          ) {
            issues.push({
              path: `${path}.content`,
              message:
                "Complex field results cannot contain hyperlink content.",
            });
          } else if (
            fieldAttrs.value.fieldKind === "complex" &&
            hasUnsupportedComplexFieldResultContent(node)
          ) {
            issues.push({
              path: `${path}.content`,
              message:
                "Complex field results can contain only run-representable content.",
            });
          }
          // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
          node.forEach((child, _offset, index) => {
            const childPath = `${path}.content[${index}]`;
            const hasHyperlink = child.marks.some(
              (mark) => mark.type.name === "hyperlink",
            );
            if (
              fieldAttrs.value.fieldKind === "complex" &&
              child.type.name === "bookmarkBoundary"
            ) {
              issues.push({
                path: childPath,
                message:
                  "Complex field results cannot contain bookmark boundaries.",
              });
            } else if (
              child.type.name === "bookmarkBoundary" &&
              !hasHyperlink
            ) {
              issues.push({
                path: childPath,
                message:
                  "Bookmark boundaries inside fields require a hyperlink parent.",
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
      // (a wrong sdtType, a malformed preserved record, etc.) surface here
      // instead of leaking into the serializer or downstream consumers.
      appendAttrIssues(path, readBlockSdtAttrs(node), issues);
      for (let i = 0; i < node.childCount; i += 1) {
        validateNode(node.child(i), `${path}.content[${i}]`, issues);
      }
      return;

    case "blockCustomXml":
      appendAttrIssues(path, readBlockCustomXmlAttrs(node), issues);
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

/**
 * Marks already validated with no issues. A mark's attrs never change, and a
 * conversion shares one instance per distinct mark across thousands of runs,
 * so each distinct mark is validated once.
 */
const validMarks = new WeakSet<Mark>();

const validateMarks = (
  marks: readonly Mark[],
  path: string,
  issues: ProseMirrorDocumentValidationIssue[],
): void => {
  for (const [index, mark] of marks.entries()) {
    if (validMarks.has(mark)) {
      continue;
    }
    const issueCount = issues.length;
    validateMark(mark, `${path}.marks[${index}]`, issues);
    if (issues.length === issueCount) {
      validMarks.add(mark);
    }
  }
};

const validateMark = (
  mark: Mark,
  markPath: string,
  issues: ProseMirrorDocumentValidationIssue[],
): void => {
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
      return;

    case "hyphenElement":
      appendAttrIssues(markPath, readHyphenElementMarkAttrs(mark), issues);
      return;

    case "textEffect":
      appendAttrIssues(markPath, readTextEffectMarkAttrs(mark), issues);
      return;

    case "underline":
      appendAttrIssues(markPath, readUnderlineMarkAttrs(mark), issues);
      return;

    case "strike":
      appendAttrIssues(markPath, readStrikeMarkAttrs(mark), issues);
      return;

    case "textColor":
      appendAttrIssues(markPath, readTextColorMarkAttrs(mark), issues);
      return;

    case "highlight":
      appendAttrIssues(markPath, readHighlightMarkAttrs(mark), issues);
      return;

    case "runShading":
      appendAttrIssues(markPath, readRunShadingMarkAttrs(mark), issues);
      return;

    case "fontSize":
      appendAttrIssues(markPath, readFontSizeMarkAttrs(mark), issues);
      return;

    case "fontFamily":
      appendAttrIssues(markPath, readFontFamilyMarkAttrs(mark), issues);
      return;

    case "language":
      appendAttrIssues(markPath, readLanguageMarkAttrs(mark), issues);
      return;

    case "characterSpacing":
      appendAttrIssues(markPath, readCharacterSpacingMarkAttrs(mark), issues);
      return;

    case "characterStyle":
      appendAttrIssues(markPath, readCharacterStyleMarkAttrs(mark), issues);
      return;

    case "emphasisMark":
      appendAttrIssues(markPath, readEmphasisMarkAttrs(mark), issues);
      return;

    case "footnoteRef":
      appendAttrIssues(markPath, readFootnoteRefMarkAttrs(mark), issues);
      return;

    case "comment":
      appendAttrIssues(markPath, readCommentMarkAttrs(mark), issues);
      return;

    case "insertion":
    case "deletion":
      appendAttrIssues(markPath, readTrackedChangeMarkAttrs(mark), issues);
      return;

    case "runPropertyChange":
      appendAttrIssues(markPath, readRunPropertyChangeMarkAttrs(mark), issues);
      return;

    case "runIdentity":
      appendAttrIssues(markPath, readRunIdentityMarkAttrs(mark), issues);
      return;

    case "runFormattingOverride":
      appendAttrIssues(
        markPath,
        readRunFormattingOverrideMarkAttrs(mark),
        issues,
      );
      return;

    case "hyperlink":
      appendAttrIssues(markPath, readHyperlinkMarkAttrs(mark), issues);
      return;

    case "inlineWrapper":
      appendAttrIssues(markPath, readInlineWrapperMarkAttrs(mark), issues);
      return;

    default:
      issues.push({
        path: `${markPath}.type.name`,
        message: `Unsupported ProseMirror mark type ${mark.type.name}.`,
      });
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
