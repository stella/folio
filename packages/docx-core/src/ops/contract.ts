import { noteContentWithAutomaticMark, noteUsesCustomMark } from "./noteMarks";
import { documentStories, replaceStoryBody, storyBody } from "./stories";
/**
 * The seed contract: what a document must be for operations to apply to it.
 *
 * A host establishes it once, when a document enters the journal: run
 * `ensureParaIds` over the package bytes (it stamps every story, headers,
 * footers and notes included, and makes ids unique), parse, then
 * {@link normalizeForOps}. {@link validateOpsDocument} states the result:
 *
 * - every paragraph of each editable story has a `w14:paraId`;
 * - no paragraph id repeats anywhere in the package, compared as hex, so an
 *   inverse can always recreate the ids it removed;
 * - no revision id (tracked changes, property changes) or content-control id
 *   repeats, so the one record carrying an id is the one it names;
 * - the body's section view says what its blocks say;
 * - each editable story holds no empty run, text node, or revision wrapper.
 *
 * Headers, footers and notes are independently addressable stories. Text-box
 * and comment paragraph ids also count toward package-wide uniqueness.
 *
 * Every operation checks the contract before applying and leaves it holding,
 * so a document an operation produced is not checked again: documents are
 * values, never modified in place.
 */

import { Result, TaggedError } from "better-result";

import type { BlockContent, Document, TableCell, TableRow } from "../model/document";
import { sectionsInStep, storyParagraphs, withBodyContent } from "./blocks";
import { countIds, countKeys, packageIdentityKeys, packageParagraphIds } from "./ids";
import {
  leafSpans,
  asParagraphContent,
  childNodes,
  type InlineNode,
  isEmptyRecord,
  rebuildNode,
} from "./leaves";
import { DOCUMENT_OP_REFUSAL_REASONS, type DocumentOpRefusalReason } from "./refusal";

/** A document that does not meet the seed contract. */
export class DocumentOpsContractError extends TaggedError("DocumentOpsContractError")<{
  message: string;
  reason: DocumentOpRefusalReason;
}> {}

const holdsEmptyRecord = (nodes: readonly InlineNode[]): boolean =>
  nodes.some((node) => isEmptyRecord(node) || holdsEmptyRecord(childNodes(node) ?? []));

const violation = (document: Document): DocumentOpsContractError | undefined => {
  const body = document.package.document;
  const paragraphs = documentStories(document).flatMap((story) =>
    storyParagraphs(storyBody(document, story)),
  );
  if (paragraphs.some(({ paragraph }) => paragraph.paraId === undefined)) {
    return new DocumentOpsContractError({
      reason: DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID,
      message: "A story paragraph has no paraId; ensureParaIds establishes one.",
    });
  }
  for (const notes of [document.package.footnotes, document.package.endnotes]) {
    const ids = notes?.map(({ id }) => id) ?? [];
    if (new Set(ids).size !== ids.length)
      return new DocumentOpsContractError({
        reason: DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID,
        message: "A note collection repeats a stable note id.",
      });
  }
  for (const story of documentStories(document)) {
    const misplaced = storyParagraphs(storyBody(document, story)).some(({ paragraph }) =>
      leafSpans(paragraph.content).some(
        ({ node }) =>
          node.type === "noteMarker" &&
          (story === "main" ||
            (story.kind !== "footnote" && story.kind !== "endnote") ||
            node.kind !== story.kind),
      ),
    );
    if (misplaced)
      return new DocumentOpsContractError({
        reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        message: "An automatic note mark must belong to its matching note story.",
      });
  }
  const repeated = [...countIds(packageParagraphIds(document.package))].find(
    ([, count]) => count > 1,
  );
  if (repeated !== undefined) {
    return new DocumentOpsContractError({
      reason: DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_BLOCK_ID,
      message: `${repeated[1]} paragraphs in the package are ${repeated[0]}.`,
    });
  }
  const repeatedRecord = [...countKeys(packageIdentityKeys(document.package))].find(
    ([, count]) => count > 1,
  );
  if (repeatedRecord !== undefined) {
    return new DocumentOpsContractError({
      reason: DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID,
      message: `${repeatedRecord[1]} records in the package carry ${repeatedRecord[0]}.`,
    });
  }
  if (!sectionsInStep(body)) {
    return new DocumentOpsContractError({
      reason: DOCUMENT_OP_REFUSAL_REASONS.SECTIONS_OUT_OF_STEP,
      message: "The body's section view does not match its blocks.",
    });
  }
  if (paragraphs.some(({ paragraph }) => holdsEmptyRecord(paragraph.content))) {
    return new DocumentOpsContractError({
      reason: DOCUMENT_OP_REFUSAL_REASONS.EMPTY_RECORD,
      message:
        "A document story holds an empty run, text node, or revision wrapper; normalizeForOps removes them.",
    });
  }
  return undefined;
};

/** Documents known to meet the contract: checked ones and ones an operation produced. */
const MEETS_CONTRACT = new WeakSet<Document>();

/** Check a document against the seed contract. */
export const validateOpsDocument = (
  document: Document,
): Result<Document, DocumentOpsContractError> => {
  if (MEETS_CONTRACT.has(document)) {
    return Result.ok(document);
  }
  const found = violation(document);
  if (found !== undefined) {
    return Result.err(found);
  }
  MEETS_CONTRACT.add(document);
  return Result.ok(document);
};

/** Record that an operation produced `document` from one that met the contract. */
export const meetsContract = (document: Document): void => {
  MEETS_CONTRACT.add(document);
};

/** The contract checked afresh, for tests that hold operations to it. */
export const contractViolation = (document: Document): DocumentOpsContractError | undefined =>
  violation(document);

const withoutEmpty = (nodes: readonly InlineNode[]): readonly InlineNode[] => {
  let changed = false;
  const out: InlineNode[] = [];
  for (const node of nodes) {
    if (isEmptyRecord(node)) {
      changed = true;
      continue;
    }
    const children = childNodes(node);
    const kept = children === undefined ? children : withoutEmpty(children);
    if (kept !== children && kept !== undefined) {
      changed = true;
      const rebuilt = rebuildNode(node, kept);
      // Runs and revision wrappers left with nothing go; other containers stay as markup.
      if (!isEmptyRecord(rebuilt)) out.push(rebuilt);
      continue;
    }
    out.push(node);
  }
  return changed ? out : nodes;
};

const normalizeBlocks = (blocks: readonly BlockContent[]): BlockContent[] => {
  const out: BlockContent[] = [];
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph": {
        const content = withoutEmpty(block.content);
        out.push(
          content === block.content ? block : { ...block, content: asParagraphContent(content) },
        );
        break;
      }
      case "table": {
        const rows: TableRow[] = [];
        for (const row of block.rows) {
          const cells: TableCell[] = [];
          for (const cell of row.cells)
            cells.push({ ...cell, content: normalizeBlocks(cell.content) });
          rows.push({ ...row, cells });
        }
        out.push({ ...block, rows });
        break;
      }
      case "blockSdt":
      case "blockCustomXml":
        out.push({ ...block, content: normalizeBlocks(block.content) });
        break;
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        out.push(block);
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
  return out;
};

/**
 * The document with every empty run and empty text node of each editable story
 * removed, and its section view derived again. Nothing else changes.
 */
export const normalizeForOps = (document: Document): Document => {
  let current = document;
  for (const story of documentStories(document)) {
    if (story !== "main" && (story.kind === "footnote" || story.kind === "endnote")) {
      const note =
        story.kind === "footnote"
          ? current.package.footnotes?.find(({ id }) => id === story.id)
          : current.package.endnotes?.find(({ id }) => id === story.id);
      if (note)
        current = replaceStoryBody({
          document: current,
          story,
          body: {
            content: noteContentWithAutomaticMark({
              note,
              customMark: noteUsesCustomMark(current, note),
            }),
          },
        });
    }
    const body = storyBody(current, story);
    current = replaceStoryBody({
      document: current,
      story,
      body: withBodyContent(body, normalizeBlocks(body.content)),
    });
  }
  return current;
};
