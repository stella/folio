/**
 * The check an operation batch's result passes before it is committed.
 *
 * An operation that reaches the document is a promise that the document can
 * still be saved: the batch reports it as applied, issues a receipt and an
 * undo handle, and the next thing the caller does may be `save()`. The model
 * validator the save runs (`validateFolioDocumentModel`, the one
 * `repackDocx` asserts) is therefore the judge here too, run on what the
 * batch produced, while the batch can still be refused.
 *
 * Two constraints shape it:
 *
 * - Cost. Converting a whole story back to the model and validating it costs
 *   about as much as the apply itself on a long document. A batch only
 *   changes the top-level blocks it touched, and ProseMirror keeps every
 *   untouched child node by identity, so only the window between the first
 *   and the last changed child is converted and validated.
 * - Blame. The apply layer holds the story, not the package: comments, notes
 *   and header/footer parts are the host's. A window can therefore report an
 *   error the batch did not cause (a comment the package defines, a range the
 *   window cuts). Only an error the same window did not already have before
 *   the batch counts; an error the document already carried is the
 *   document's, and refusing every later edit near it would help no one.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { ValidateDocumentModelIssue } from "@stll/docx-core";

import { validateFolioDocumentModel } from "../docx/modelValidation";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import type { Comment, Document, HeaderFooter, NumberingDefinitions } from "../types/document";
import { sectionPropertiesOf } from "../prosemirror/sectionCarrier";

/** What the story's package knows that the story itself does not. */
export type FolioOperationResultValidationContext = {
  /**
   * The document's numbering definitions; `null` when the story states none,
   * `undefined` when it cannot say (no numbering plugin), which leaves
   * numbering references unchecked.
   */
  numbering: NumberingDefinitions | null | undefined;
  /** Comment ids the batch allocated; the host holds the comments themselves. */
  createdCommentIds: readonly number[];
  /**
   * The author the host stamps on those comments, when the batch names one;
   * the host writes the batch's author, so a blank one fails there too.
   */
  commentAuthor?: string;
};

type ChangedWindow = { from: number; beforeTo: number; afterTo: number };

/**
 * The top-level children that differ between two versions of one story,
 * found by node identity from both ends. `null` when every child is shared.
 */
const changedTopLevelWindow = (before: PMNode, after: PMNode): ChangedWindow | null => {
  if (before === after) {
    return null;
  }
  const beforeCount = before.childCount;
  const afterCount = after.childCount;
  let from = 0;
  while (from < beforeCount && from < afterCount && before.child(from) === after.child(from)) {
    from += 1;
  }
  let beforeTo = beforeCount;
  let afterTo = afterCount;
  while (
    beforeTo > from &&
    afterTo > from &&
    before.child(beforeTo - 1) === after.child(afterTo - 1)
  ) {
    beforeTo -= 1;
    afterTo -= 1;
  }
  if (from === beforeTo && from === afterTo) {
    return null;
  }
  return { from, beforeTo, afterTo };
};

const TOP_LEVEL_PATH = /^package\.document\.content\[(\d+)\]/u;

/** Rebase a window-relative block path onto the story's own block index. */
const rebasePath = (path: string, offset: number): string =>
  path.replace(
    TOP_LEVEL_PATH,
    (_match, index: string) => `package.document.content[${String(Number(index) + offset)}]`,
  );

const syntheticComment = (id: number, author: string): Comment => ({ id, author, content: [] });

type KnownSectionParts = {
  headers: Map<string, HeaderFooter>;
  footers: Map<string, HeaderFooter>;
};

/** A story can identify its existing references even though the host owns the parts. */
const knownSectionParts = (doc: PMNode): KnownSectionParts => {
  const headers = new Map<string, HeaderFooter>();
  const footers = new Map<string, HeaderFooter>();
  doc.descendants((node) => {
    if (node.type.name !== "paragraph") return true;
    const section = sectionPropertiesOf(node);
    for (const { rId, type } of section?.headerReferences ?? []) {
      headers.set(rId, { type: "header", hdrFtrType: type, content: [] });
    }
    for (const { rId, type } of section?.footerReferences ?? []) {
      footers.set(rId, { type: "footer", hdrFtrType: type, content: [] });
    }
    return false;
  });
  return { headers, footers };
};

const windowErrors = (
  doc: PMNode,
  from: number,
  to: number,
  context: FolioOperationResultValidationContext,
  knownParts?: KnownSectionParts,
): ValidateDocumentModelIssue[] => {
  if (to <= from) {
    return [];
  }
  const children: PMNode[] = [];
  for (let index = from; index < to; index += 1) {
    children.push(doc.child(index));
  }
  let content;
  try {
    // Nothing is saved from this conversion: the validator reads references
    // and structure, not which tier a paragraph property came from, so there
    // is deliberately no property-source base to restore against.
    content = proseDocToBlocks(doc.copy(Fragment.fromArray(children)), []);
  } catch (error) {
    return [
      {
        path: `package.document.content[${String(from)}]`,
        message: `The content cannot be converted for saving: ${
          error instanceof Error ? error.message : String(error)
        }`,
        severity: "error",
      },
    ];
  }
  const document: Document = {
    package: {
      document: {
        content,
        comments: context.createdCommentIds.map((id) =>
          syntheticComment(id, context.commentAuthor ?? ""),
        ),
      },
      ...(context.numbering !== null &&
        context.numbering !== undefined && { numbering: context.numbering }),
      ...(knownParts && { headers: knownParts.headers, footers: knownParts.footers }),
    },
  };
  return validateFolioDocumentModel(document)
    .issues.filter(
      (issue) =>
        issue.severity === "error" &&
        // Without the numbering plugin the story cannot say which instances
        // exist, so a reference is not something it can hold against a batch.
        !(context.numbering === undefined && issue.path.endsWith(".numPr.numId")),
    )
    .map(({ path, message, severity }) => ({ path: rebasePath(path, from), message, severity }));
};

/**
 * The save-validator errors the batch introduced: errors in the changed window
 * of `after` that the same window of `before` did not already report. Empty
 * when the result is as saveable as the story it started from.
 */
export const findIntroducedModelErrors = (
  before: PMNode,
  after: PMNode,
  context: FolioOperationResultValidationContext,
): ValidateDocumentModelIssue[] => {
  const window = changedTopLevelWindow(before, after);
  if (window === null) {
    return [];
  }
  let afterErrors = windowErrors(after, window.from, window.afterTo, context);
  // A newly changed window can include a section record that was already in
  // the story. Its header/footer parts live in the host, outside this window.
  // Supply only references known before the batch, so a new dangling rId
  // still fails validation.
  let parts: KnownSectionParts | undefined;
  if (afterErrors.some(({ message }) => message.startsWith("Section references missing "))) {
    parts = knownSectionParts(before);
    afterErrors = windowErrors(after, window.from, window.afterTo, context, parts);
  }
  if (afterErrors.length === 0) {
    return afterErrors;
  }
  // The story before the batch held none of the batch's comments.
  const known = new Map<string, number>();
  for (const { message } of windowErrors(
    before,
    window.from,
    window.beforeTo,
    {
      ...context,
      createdCommentIds: [],
    },
    parts,
  )) {
    known.set(message, (known.get(message) ?? 0) + 1);
  }
  return afterErrors.filter(({ message }) => {
    const remaining = known.get(message) ?? 0;
    if (remaining === 0) {
      return true;
    }
    known.set(message, remaining - 1);
    return false;
  });
};

export const describeModelError = ({ path, message }: ValidateDocumentModelIssue): string =>
  `The result would not save: ${path}: ${message}`;
