import { TaggedError } from "better-result";

import type { DocumentOpType } from "./types";

/**
 * Why an operation was not applied. A refused operation changes nothing: the
 * document it was applied to is returned as it was.
 */
export const DOCUMENT_OP_REFUSAL_REASONS = Object.freeze({
  /** No paragraph in the story carries the id. */
  BLOCK_NOT_FOUND: "blockNotFound",
  /** An offset is not an integer from `0` to the paragraph's length, or a range runs backwards. */
  INVALID_OFFSET: "invalidOffset",
  /** A range starts and ends in different paragraphs or stories. */
  CROSS_BLOCK_RANGE: "crossBlockRange",
  /** A position falls between the two halves of a surrogate pair. */
  SPLITS_SURROGATE_PAIR: "splitsSurrogatePair",
  /** Inserted text is empty, holds a tab or line-break character, or cannot be written to XML. */
  INVALID_TEXT: "invalidText",
  /**
   * A paragraph id the operation (or its inverse) creates is not eight hex
   * digits below `0x80000000`, or is the reserved `00000000`.
   */
  INVALID_BLOCK_ID: "invalidBlockId",
  /**
   * An id the operation brings in (a paragraph id, or a revision or
   * content-control id a replacement carries) is already used in the package.
   */
  ID_COLLISION: "idCollision",
  /** The paragraphs to join or replace are not adjacent siblings of one container. */
  NOT_ADJACENT: "notAdjacent",
  /**
   * The operation would move or remove a section break (joining a paragraph
   * that ends a section, replacing paragraphs across or at one); that is a
   * section operation.
   */
  SECTION_BOUNDARY: "sectionBoundary",
  /** Direct text would land inside tracked-deleted or moved-away content. */
  INSIDE_TRACKED_DELETION: "insideTrackedDeletion",
  /**
   * The operation cuts a record carrying a revision or content-control id in
   * two and names too few `newIds` for the halves after the first.
   */
  NEEDS_NEW_IDS: "needsNewIds",
  /** A new revision or content-control id is not an integer from 0 to 2^31 - 1. */
  INVALID_NEW_ID: "invalidNewId",
  /**
   * The operation would merge two records carrying the same revision or
   * content-control id: nothing could tell them apart again to undo it.
   */
  SHARED_ID: "sharedId",
  /**
   * What the operation expects (paragraphs to replace, a slice to delete,
   * the values a patch replaced) is not what is there.
   */
  STALE: "stale",
  /**
   * The records at a position are not the shape the operation needs: an open
   * end meets a record it cannot merge with, records to join differ, a slice
   * would merge records that were separate, or there are fewer records to cut
   * than the operation names.
   */
  STRUCTURE_MISMATCH: "structureMismatch",
  /** Inserted content is empty, or holds an empty run or text node. */
  EMPTY_CONTENT: "emptyContent",
  /** A replacement names no paragraph to replace, or none to put in its place. */
  EMPTY_BLOCK_LIST: "emptyBlockList",
  /** Seed contract: a main-story paragraph has no `paraId`. */
  MISSING_BLOCK_ID: "missingBlockId",
  /** Seed contract: two paragraphs in the package carry one id. */
  DUPLICATE_BLOCK_ID: "duplicateBlockId",
  /** Seed contract: two records in the package carry one revision or content-control id. */
  DUPLICATE_RECORD_ID: "duplicateRecordId",
  /** Seed contract: the body's section view does not match its blocks. */
  SECTIONS_OUT_OF_STEP: "sectionsOutOfStep",
  /** Seed contract: the main story holds an empty run or empty text node. */
  EMPTY_RECORD: "emptyRecord",
} as const);

/** One of {@link DOCUMENT_OP_REFUSAL_REASONS}. */
export type DocumentOpRefusalReason =
  (typeof DOCUMENT_OP_REFUSAL_REASONS)[keyof typeof DOCUMENT_OP_REFUSAL_REASONS];

/** An operation the document cannot take as stated. */
export class DocumentOpRefusal extends TaggedError("DocumentOpRefusal")<{
  message: string;
  reason: DocumentOpRefusalReason;
  opType: DocumentOpType;
}> {}
