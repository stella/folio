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
  /** A new paragraph id is not an 8-digit `ST_LongHexNumber` below `0x80000000`, or is zero. */
  INVALID_BLOCK_ID: "invalidBlockId",
  /** A new paragraph id is already used in the story. */
  ID_COLLISION: "idCollision",
  /** The paragraphs to join or replace are not adjacent siblings of one container. */
  NOT_ADJACENT: "notAdjacent",
  /** The paragraph to join ends a section; removing its mark is a section operation. */
  SECTION_BOUNDARY: "sectionBoundary",
  /** Direct text would land inside tracked-deleted or moved-away content. */
  INSIDE_TRACKED_DELETION: "insideTrackedDeletion",
  /**
   * A split would cut a tracked change or content control in two, and both
   * halves would carry the one revision or control id.
   */
  SPLITS_IDENTIFIED_CONTAINER: "splitsIdentifiedContainer",
  /** The paragraphs a replacement expects are not the ones in the document. */
  STALE: "stale",
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
