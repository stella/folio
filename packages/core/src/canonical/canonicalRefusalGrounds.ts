import type { DocumentOpRefusalReason } from "@stll/docx-core/ops";

import {
  CANONICAL_CONDITION,
  CANONICAL_GAP,
  CANONICAL_RULE,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";

const STALE = { condition: CANONICAL_CONDITION.stale } as const;
const INVALID_REQUEST = {
  condition: CANONICAL_CONDITION.invalidRequest,
} as const;

/**
 * The ground of every document operation refusal: a moved-on document is
 * stale, an exhausted id space names it, a link suggestion review records
 * cannot hold and content OOXML has no revision record for are the
 * document's rules, content the tracked-change records cannot express yet is
 * a gap, and every other reason describes a request the document does not
 * allow.
 */
const DOCUMENT_OP_REFUSAL_GROUNDS = {
  unsupportedSchema: INVALID_REQUEST,
  blockNotFound: INVALID_REQUEST,
  invalidOperation: INVALID_REQUEST,
  noChange: INVALID_REQUEST,
  dependentRecords: INVALID_REQUEST,
  invalidOffset: INVALID_REQUEST,
  crossBlockRange: INVALID_REQUEST,
  splitsSurrogatePair: INVALID_REQUEST,
  invalidText: INVALID_REQUEST,
  invalidBlockId: INVALID_REQUEST,
  idCollision: INVALID_REQUEST,
  notAdjacent: INVALID_REQUEST,
  sectionBoundary: INVALID_REQUEST,
  insideTrackedDeletion: INVALID_REQUEST,
  needsNewIds: { condition: CANONICAL_CONDITION.idSpace },
  invalidNewId: INVALID_REQUEST,
  sharedId: INVALID_REQUEST,
  stale: STALE,
  structureMismatch: INVALID_REQUEST,
  emptyContent: INVALID_REQUEST,
  emptyBlockList: INVALID_REQUEST,
  revisionConflict: INVALID_REQUEST,
  pendingTableChange: INVALID_REQUEST,
  pendingRunPropertyChange: INVALID_REQUEST,
  pendingParagraphPropertyChange: INVALID_REQUEST,
  joinIntoPendingTable: INVALID_REQUEST,
  containerFinalMark: INVALID_REQUEST,
  insideInlineControl: INVALID_REQUEST,
  untrackable: { gap: CANONICAL_GAP.trackedChangeRecords },
  tableSnapshotPresence: { rule: CANONICAL_RULE.tableSnapshotPresence },
  pendingBreakTakeover: INVALID_REQUEST,
  pendingParagraphMarkDeletion: INVALID_REQUEST,
  pendingParagraphMarkInsertion: INVALID_REQUEST,
  pendingParagraphJoin: INVALID_REQUEST,
  pendingGridChange: INVALID_REQUEST,
  unrecordableLinkSuggestion: { rule: CANONICAL_RULE.linkSuggestionShape },
  noRevisionRecord: { rule: CANONICAL_RULE.revisionContentShape },
  missingBlockId: INVALID_REQUEST,
  duplicateBlockId: INVALID_REQUEST,
  duplicateRecordId: INVALID_REQUEST,
  sectionsOutOfStep: STALE,
  emptyRecord: INVALID_REQUEST,
} as const satisfies Record<DocumentOpRefusalReason, CanonicalRefusalGround>;

/** The refusal ground a document operation refused for `reason` reports. */
export const groundForDocumentOpRefusal = (
  reason: DocumentOpRefusalReason,
): CanonicalRefusalGround => DOCUMENT_OP_REFUSAL_GROUNDS[reason];
