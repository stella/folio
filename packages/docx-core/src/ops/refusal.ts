import { TaggedError, type TaggedErrorClass } from "better-result";

import type { ParagraphMarkChange } from "../model/content";
import type { DocumentOpType } from "./types";

const DocumentOpRefusalBase: TaggedErrorClass<"DocumentOpRefusal"> =
  TaggedError("DocumentOpRefusal");

/**
 * Why an operation was not applied. A refused operation changes nothing: the
 * document it was applied to is returned as it was.
 */
export const DOCUMENT_OP_REFUSAL_REASONS: Readonly<{
  /** The journal envelope uses a schema this reader does not implement. */
  readonly UNSUPPORTED_SCHEMA: "unsupportedSchema";
  /** No paragraph in the story carries the id. */
  readonly BLOCK_NOT_FOUND: "blockNotFound";
  /** A raw operation is missing required fields or supplies their wrong shape. */
  readonly INVALID_OPERATION: "invalidOperation";
  /** The requested topology edit changes nothing. */
  readonly NO_CHANGE: "noChange";
  /** Removed content carries anchors or review records that require coordinated resolution. */
  readonly DEPENDENT_RECORDS: "dependentRecords";
  /** An offset is not an integer from `0` to the paragraph's length, or a range runs backwards. */
  readonly INVALID_OFFSET: "invalidOffset";
  /** A range starts and ends in different paragraphs or stories. */
  readonly CROSS_BLOCK_RANGE: "crossBlockRange";
  /** A position falls between the two halves of a surrogate pair. */
  readonly SPLITS_SURROGATE_PAIR: "splitsSurrogatePair";
  /** Inserted text is empty, holds a tab or line-break character, or cannot be written to XML. */
  readonly INVALID_TEXT: "invalidText";
  /**
   * A paragraph id the operation (or its inverse) creates is not eight hex
   * digits below `0x80000000`, or is the reserved `00000000`.
   */
  readonly INVALID_BLOCK_ID: "invalidBlockId";
  /**
   * An id the operation brings in (a paragraph id, or a revision or
   * content-control id a replacement carries) is already used in the package.
   */
  readonly ID_COLLISION: "idCollision";
  /** The paragraphs to join or replace are not adjacent siblings of one container. */
  readonly NOT_ADJACENT: "notAdjacent";
  /**
   * The operation would move or remove a section break (joining a paragraph
   * that ends a section, replacing paragraphs across or at one); that is a
   * section operation. Or it would put one inside a table cell, where no
   * section ends.
   */
  readonly SECTION_BOUNDARY: "sectionBoundary";
  /** Direct text would land inside tracked-deleted or moved-away content. */
  readonly INSIDE_TRACKED_DELETION: "insideTrackedDeletion";
  /**
   * The operation cuts a record carrying a revision or content-control id in
   * two and names too few `newIds` for the halves after the first.
   */
  readonly NEEDS_NEW_IDS: "needsNewIds";
  /** A new revision or content-control id (a stamp's too) is not an integer from 0 to 2^31 - 1. */
  readonly INVALID_NEW_ID: "invalidNewId";
  /**
   * The operation would merge two records carrying the same revision or
   * content-control id: nothing could tell them apart again to undo it.
   */
  readonly SHARED_ID: "sharedId";
  /**
   * What the operation expects (paragraphs to replace, a slice to delete,
   * the values a patch replaced) is not what is there.
   */
  readonly STALE: "stale";
  /**
   * The records at a position are not the shape the operation needs: an open
   * end meets a record it cannot merge with, records to join differ, a slice
   * would merge records that were separate, or there are fewer records to cut
   * than the operation names.
   */
  readonly STRUCTURE_MISMATCH: "structureMismatch";
  /** Inserted content is empty, or holds an empty run or text node. */
  readonly EMPTY_CONTENT: "emptyContent";
  /** A replacement names no paragraph to replace, or none to put in its place. */
  readonly EMPTY_BLOCK_LIST: "emptyBlockList";
  /**
   * The operation would give a tracked split a mark or property change of
   * its own, or would put a pasted paragraph's review over the review the
   * destination already carries.
   */
  readonly REVISION_CONFLICT: "revisionConflict";
  /**
   * The operation cannot preserve an existing table, row, cell or grid
   * review, or would overwrite one of the same kind. Resolve it before
   * changing the table.
   */
  readonly PENDING_TABLE_CHANGE: "pendingTableChange";
  /**
   * A suggestion would fold into a run's pending property change that is not
   * its own. A run holds one property change and a suggestion is resolved by
   * its id: rejecting it would discard the other change with its own.
   */
  readonly PENDING_RUN_PROPERTY_CHANGE: "pendingRunPropertyChange";
  /**
   * A suggestion would fold into a paragraph's pending property change that
   * is not its own, for the same reason as
   * {@link DOCUMENT_OP_REFUSAL_REASONS.PENDING_RUN_PROPERTY_CHANGE}; or a
   * tracked insertion after a container's final paragraph would take over
   * the mark of one whose properties carry a pending change.
   */
  readonly PENDING_PARAGRAPH_PROPERTY_CHANGE: "pendingParagraphPropertyChange";
  /**
   * A tracked join would run a paragraph, through one whose own mark is
   * already going, into the table after it while that table's first row or
   * cell has a pending insertion and cannot receive what runs in.
   */
  readonly JOIN_INTO_PENDING_TABLE: "joinIntoPendingTable";
  /** The operation would put a mark on a paragraph that ends its story body or table cell. */
  readonly CONTAINER_FINAL_MARK: "containerFinalMark";
  /**
   * A paragraph break inside an inline content control: the control lives in
   * one paragraph, so the break has no form.
   */
  readonly INSIDE_INLINE_CONTROL: "insideInlineControl";
  /**
   * No tracked change can record the edit, or no resolution op can carry it
   * out yet: the paragraph mark's run properties, a deleted paragraph whose
   * hyperlink or content control holds a comment anchor, a section-ending
   * mark resolved across a standing range boundary, a change inside a
   * simple field holding content the field cannot hold.
   */
  readonly UNTRACKABLE: "untrackable";
  /**
   * A tracked paragraph would take over the mark of a paragraph whose break
   * is pending deletion or moving away. No placement of that record keeps
   * both laws: accepting equals the direct edit, and rejecting the insertion
   * alone gives the original back.
   */
  readonly PENDING_BREAK_TAKEOVER: "pendingBreakTakeover";
  /**
   * A tracked operation would record a change of a paragraph mark whose
   * break is pending deletion (deleted, moving away, or added and then
   * deleted), which holds one such record: joining or deleting through it,
   * putting a table or row against it, or, after a container's final
   * paragraph, taking it over, since the container's new final paragraph
   * keeps no record a pending change moves to it.
   */
  readonly PENDING_PARAGRAPH_MARK_DELETION: "pendingParagraphMarkDeletion";
  /**
   * As {@link DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION},
   * for a paragraph mark whose break is pending insertion (inserted or moved
   * here) and not deleted since.
   */
  readonly PENDING_PARAGRAPH_MARK_INSERTION: "pendingParagraphMarkInsertion";
  /**
   * A paragraph the operation would edit separately joins the next one
   * through a mark pending deletion (deleted or moving away): accepting that
   * deletion gives the joined paragraph the next paragraph's properties, so
   * the edit would land on another paragraph or merge two named ones.
   */
  readonly PENDING_PARAGRAPH_JOIN: "pendingParagraphJoin";
  /**
   * A tracked column insertion in a table whose columns already carry a
   * pending change. A table records one previous grid (`w:tblGridChange`), so
   * a second insertion could not be rejected alone: the columns of one change
   * go in one operation (`InsertColumnOp.count`).
   */
  readonly PENDING_GRID_CHANGE: "pendingGridChange";
  /**
   * A suggested link change would leave a piece of a link holding only
   * bookmark or comment markers, or would copy content that names a record of
   * its own (preserved run markup, a marker inside a copied text box, or a
   * note holding a tracked change, comment or bookmark).
   */
  readonly UNRECORDABLE_LINK_SUGGESTION: "unrecordableLinkSuggestion";
  /** Seed contract: a main-story paragraph has no `paraId`. */
  readonly MISSING_BLOCK_ID: "missingBlockId";
  /** Seed contract: two paragraphs in the package carry one id. */
  readonly DUPLICATE_BLOCK_ID: "duplicateBlockId";
  /** Seed contract: two records in the package carry one revision or content-control id. */
  readonly DUPLICATE_RECORD_ID: "duplicateRecordId";
  /** Seed contract: the body's section view does not match its blocks. */
  readonly SECTIONS_OUT_OF_STEP: "sectionsOutOfStep";
  /** Seed contract: the main story holds an empty run or empty text node. */
  readonly EMPTY_RECORD: "emptyRecord";
}> = Object.freeze({
  /** The journal envelope uses a schema this reader does not implement. */
  UNSUPPORTED_SCHEMA: "unsupportedSchema",
  /** No paragraph in the story carries the id. */
  BLOCK_NOT_FOUND: "blockNotFound",
  /** A raw operation is missing required fields or supplies their wrong shape. */
  INVALID_OPERATION: "invalidOperation",
  /** The requested topology edit changes nothing. */
  NO_CHANGE: "noChange",
  /** Removed content carries anchors or review records that require coordinated resolution. */
  DEPENDENT_RECORDS: "dependentRecords",
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
   * section operation. Or it would put one inside a table cell, where no
   * section ends.
   */
  SECTION_BOUNDARY: "sectionBoundary",
  /** Direct text would land inside tracked-deleted or moved-away content. */
  INSIDE_TRACKED_DELETION: "insideTrackedDeletion",
  /**
   * The operation cuts a record carrying a revision or content-control id in
   * two and names too few `newIds` for the halves after the first.
   */
  NEEDS_NEW_IDS: "needsNewIds",
  /** A new revision or content-control id (a stamp's too) is not an integer from 0 to 2^31 - 1. */
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
  /**
   * The operation would give a tracked split a mark or property change of
   * its own, or would put a pasted paragraph's review over the review the
   * destination already carries.
   */
  REVISION_CONFLICT: "revisionConflict",
  /**
   * The operation would change a table, row, cell or grid that already
   * carries a pending change of the kind it would record: a row pending
   * insertion or deletion, a table, row or cell property change, a grid
   * change, or cells already suggested for deletion; or would change the
   * columns or cells of a table under review. Each element holds one record
   * of a kind: resolve the pending one first.
   */
  PENDING_TABLE_CHANGE: "pendingTableChange",
  /**
   * A suggestion would fold into a run's pending property change that is not
   * its own. A run holds one property change and a suggestion is resolved by
   * its id: rejecting it would discard the other change with its own.
   */
  PENDING_RUN_PROPERTY_CHANGE: "pendingRunPropertyChange",
  /**
   * A suggestion would fold into a paragraph's pending property change that
   * is not its own, for the same reason as
   * {@link DOCUMENT_OP_REFUSAL_REASONS.PENDING_RUN_PROPERTY_CHANGE}; or a
   * tracked insertion after a container's final paragraph would take over
   * the mark of one whose properties carry a pending change.
   */
  PENDING_PARAGRAPH_PROPERTY_CHANGE: "pendingParagraphPropertyChange",
  /**
   * A tracked join would run a paragraph, through one whose own mark is
   * already going, into the table after it while that table's first row or
   * cell is under review: the table cannot receive what runs in.
   */
  JOIN_INTO_PENDING_TABLE: "joinIntoPendingTable",
  /** The operation would put a mark on a paragraph that ends its story body or table cell. */
  CONTAINER_FINAL_MARK: "containerFinalMark",
  /**
   * A paragraph break inside an inline content control: the control lives in
   * one paragraph, so the break has no form.
   */
  INSIDE_INLINE_CONTROL: "insideInlineControl",
  /**
   * No tracked change can record the edit, or no resolution op can carry it
   * out yet: the paragraph mark's run properties, a deleted paragraph whose
   * hyperlink or content control holds a comment anchor, a section-ending
   * mark resolved across a standing range boundary, a change inside a
   * simple field holding content the field cannot hold.
   */
  UNTRACKABLE: "untrackable",
  /** A second tracked column insertion beside a pending grid change. */
  PENDING_GRID_CHANGE: "pendingGridChange",
  /**
   * A tracked paragraph would take over the mark of a paragraph whose break
   * is pending deletion or moving away. No placement of that record keeps
   * both laws: accepting equals the direct edit, and rejecting the insertion
   * alone gives the original back.
   */
  PENDING_BREAK_TAKEOVER: "pendingBreakTakeover",
  /**
   * A tracked operation would record a change of a paragraph mark whose
   * break is pending deletion (deleted, moving away, or added and then
   * deleted), which holds one such record: joining or deleting through it,
   * putting a table or row against it, or, after a container's final
   * paragraph, taking it over, since the container's new final paragraph
   * keeps no record a pending change moves to it.
   */
  PENDING_PARAGRAPH_MARK_DELETION: "pendingParagraphMarkDeletion",
  /**
   * As {@link DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION},
   * for a paragraph mark whose break is pending insertion (inserted or moved
   * here) and not deleted since.
   */
  PENDING_PARAGRAPH_MARK_INSERTION: "pendingParagraphMarkInsertion",
  /**
   * A paragraph the operation would edit separately joins the next one
   * through a mark pending deletion (deleted or moving away): accepting that
   * deletion gives the joined paragraph the next paragraph's properties, so
   * the edit would land on another paragraph or merge two named ones.
   */
  PENDING_PARAGRAPH_JOIN: "pendingParagraphJoin",
  /**
   * A suggested link change would leave a piece of a link holding only
   * bookmark or comment markers, or would copy content that names a record of
   * its own (preserved run markup, a marker inside a copied text box, or a
   * note holding a tracked change, comment or bookmark).
   */
  UNRECORDABLE_LINK_SUGGESTION: "unrecordableLinkSuggestion",
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

/** Which operations a refusal reason can refuse. */
export const REFUSAL_SCOPES: Readonly<{
  /** Any operation, direct or tracked. */
  readonly ANY: "any";
  /**
   * Only a tracked one: a pending review record, or the tracked change
   * itself, cannot take the edit, which applied directly may stand.
   */
  readonly TRACKED: "tracked";
}> = Object.freeze({ ANY: "any", TRACKED: "tracked" } as const);

/** One of {@link REFUSAL_SCOPES}. */
export type RefusalScope = (typeof REFUSAL_SCOPES)[keyof typeof REFUSAL_SCOPES];

const { ANY, TRACKED } = REFUSAL_SCOPES;

/** The scope of every refusal reason. */
export const DOCUMENT_OP_REFUSAL_SCOPES: Readonly<
  Record<DocumentOpRefusalReason, RefusalScope>
> = Object.freeze({
  unsupportedSchema: ANY,
  blockNotFound: ANY,
  invalidOperation: ANY,
  noChange: ANY,
  dependentRecords: ANY,
  invalidOffset: ANY,
  crossBlockRange: ANY,
  splitsSurrogatePair: ANY,
  invalidText: ANY,
  invalidBlockId: ANY,
  idCollision: ANY,
  notAdjacent: ANY,
  sectionBoundary: ANY,
  insideTrackedDeletion: ANY,
  needsNewIds: ANY,
  invalidNewId: ANY,
  sharedId: ANY,
  stale: ANY,
  structureMismatch: ANY,
  emptyContent: ANY,
  emptyBlockList: ANY,
  revisionConflict: TRACKED,
  pendingTableChange: ANY,
  pendingRunPropertyChange: TRACKED,
  pendingParagraphPropertyChange: TRACKED,
  joinIntoPendingTable: TRACKED,
  containerFinalMark: ANY,
  insideInlineControl: ANY,
  untrackable: TRACKED,
  pendingBreakTakeover: TRACKED,
  pendingParagraphMarkDeletion: TRACKED,
  pendingParagraphMarkInsertion: TRACKED,
  pendingParagraphJoin: ANY,
  pendingGridChange: TRACKED,
  unrecordableLinkSuggestion: TRACKED,
  missingBlockId: ANY,
  duplicateBlockId: ANY,
  duplicateRecordId: ANY,
  sectionsOutOfStep: ANY,
  emptyRecord: ANY,
} as const satisfies Record<DocumentOpRefusalReason, RefusalScope>);

/** The reasons only a tracked operation is refused for. */
export const TRACKED_REFUSAL_REASONS: readonly DocumentOpRefusalReason[] =
  Object.freeze(
    Object.values(DOCUMENT_OP_REFUSAL_REASONS).filter(
      (reason) => DOCUMENT_OP_REFUSAL_SCOPES[reason] === TRACKED,
    ),
  );

/** A refusal reason naming what a paragraph mark under review has pending. */
type PendingParagraphMarkReason =
  | typeof DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION
  | typeof DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_INSERTION;

/**
 * Why a tracked edit cannot record over `mark`, by what it has pending: a
 * break that accepting every change takes (deleted, moving away, or added
 * and then deleted) is a pending deletion; one only added is a pending
 * insertion.
 */
export const pendingParagraphMarkReason = ({
  kind,
  deletion,
}: ParagraphMarkChange): PendingParagraphMarkReason => {
  switch (kind) {
    case "del":
    case "moveFrom":
      return DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION;
    case "ins":
    case "moveTo":
      return deletion === undefined
        ? DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_INSERTION
        : DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_DELETION;
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};

/** An operation the document cannot take as stated. */
export class DocumentOpRefusal extends DocumentOpRefusalBase<{
  message: string;
  reason: DocumentOpRefusalReason;
  opType: DocumentOpType | undefined;
  /**
   * Paragraph-mark revision whose resolution the document blocks: it would
   * move text across a standing range boundary, or leave its container
   * without a final paragraph.
   */
  blockedParagraphMarkRevisionId?: number;
}> {}
