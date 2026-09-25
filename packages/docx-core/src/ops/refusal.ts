import { TaggedError } from "better-result";

import type { DocumentOpType } from "./types";

/**
 * Why an operation was not applied. A refused operation changes nothing: the
 * document it was applied to is returned as it was.
 */
export const DOCUMENT_OP_REFUSAL_REASONS = Object.freeze({
  /** No paragraph in the story carries the id. */
  BLOCK_NOT_FOUND: "blockNotFound",
  /** More than one paragraph in the story carries the id. */
  AMBIGUOUS_BLOCK_ID: "ambiguousBlockId",
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
  /** A paragraph id the operation brings in is already used somewhere in the package. */
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
   * The operation would cut a tracked change or content control in two, so
   * both halves carried its one id, or merge two of them into one record
   * carrying only one.
   */
  SPLITS_IDENTIFIED_CONTAINER: "splitsIdentifiedContainer",
  /** The content an operation expects (paragraphs to replace, a slice to delete) is not what is there. */
  STALE: "stale",
  /**
   * The records at a position are not the shape the operation needs: an open
   * end meets a record it cannot merge with, records to join differ, or there
   * are fewer records to cut than the operation names.
   */
  STRUCTURE_MISMATCH: "structureMismatch",
  /** Inserted content is empty. */
  EMPTY_CONTENT: "emptyContent",
  /** A replacement names no paragraph to replace, or none to put in its place. */
  EMPTY_BLOCK_LIST: "emptyBlockList",
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
