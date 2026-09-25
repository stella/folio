/**
 * The seed contract: what a document must be for operations to apply to it.
 *
 * A host establishes it once, when a document enters the journal: run
 * `ensureParaIds` over the package bytes (it stamps every story, headers,
 * footers and notes included, and makes ids unique), parse, then
 * {@link normalizeForOps}. {@link validateOpsDocument} states the result:
 *
 * - every paragraph of the main story has a `w14:paraId`;
 * - no paragraph id repeats anywhere in the package, compared as hex, so an
 *   inverse can always recreate the ids it removed;
 * - no revision id (tracked changes, property changes) or content-control id
 *   repeats, so the one record carrying an id is the one it names;
 * - the body's section view says what its blocks say;
 * - the main story holds no empty run or empty text node.
 *
 * Paragraphs in text boxes, headers, footers, notes and comments are other
 * stories: they are not addressable in schema version 1, and their ids count
 * toward uniqueness like any other.
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
  const paragraphs = storyParagraphs(body);
  if (paragraphs.some(({ paragraph }) => paragraph.paraId === undefined)) {
    return new DocumentOpsContractError({
      reason: DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID,
      message: "A main-story paragraph has no paraId; ensureParaIds establishes one.",
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
      message: "The main story holds an empty run or text node; normalizeForOps removes them.",
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
      // A run left with nothing is empty too; a container stays, as markup.
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
 * The document with every empty run and empty text node of the main story
 * removed, and its section view derived again. Nothing else changes.
 */
export const normalizeForOps = (document: Document): Document => {
  const body = document.package.document;
  return {
    ...document,
    package: {
      ...document.package,
      document: withBodyContent(body, normalizeBlocks(body.content)),
    },
  };
};
