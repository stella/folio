import {
  DOCUMENT_OP_REFUSAL_REASONS,
  type DocumentOpRefusal,
} from "@stll/docx-core/ops";

import type { FolioAIEditSkipReason } from "../ai-edits/types";
import {
  CANONICAL_GAP,
  type CanonicalGap,
} from "../types/canonicalCapabilities";

type CompilerRefusalSkip = {
  reason: FolioAIEditSkipReason;
  message: string;
  refusal?: { gap: CanonicalGap };
};

/**
 * The skip a refused compilation reports. A deliberate refusal names its own
 * reason, as the legacy engine reports it; any other is a capability the
 * compiler lacks.
 */
export const compilerRefusalSkip = ({
  reason,
  message,
}: DocumentOpRefusal): CompilerRefusalSkip => {
  switch (reason) {
    case DOCUMENT_OP_REFUSAL_REASONS.PENDING_RUN_PROPERTY_CHANGE:
    case DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE:
      return { reason, message };
    // The paragraph whose mark is already going stands before a table under
    // review: resolving that removed mark first lets the join through.
    case DOCUMENT_OP_REFUSAL_REASONS.JOIN_INTO_PENDING_TABLE:
      return { reason: "pendingParagraphMarkDeletion", message };
    case DOCUMENT_OP_REFUSAL_REASONS.UNSUPPORTED_SCHEMA:
    case DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND:
    case DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION:
    case DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE:
    case DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS:
    case DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET:
    case DOCUMENT_OP_REFUSAL_REASONS.CROSS_BLOCK_RANGE:
    case DOCUMENT_OP_REFUSAL_REASONS.SPLITS_SURROGATE_PAIR:
    case DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT:
    case DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID:
    case DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION:
    case DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT:
    case DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY:
    case DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION:
    case DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS:
    case DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID:
    case DOCUMENT_OP_REFUSAL_REASONS.SHARED_ID:
    case DOCUMENT_OP_REFUSAL_REASONS.STALE:
    case DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH:
    case DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT:
    case DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST:
    case DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT:
    case DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK:
    case DOCUMENT_OP_REFUSAL_REASONS.INSIDE_INLINE_CONTROL:
    case DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE:
    case DOCUMENT_OP_REFUSAL_REASONS.PENDING_BREAK_TAKEOVER:
    case DOCUMENT_OP_REFUSAL_REASONS.UNRECORDABLE_LINK_SUGGESTION:
    case DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID:
    case DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_BLOCK_ID:
    case DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID:
    case DOCUMENT_OP_REFUSAL_REASONS.SECTIONS_OUT_OF_STEP:
    case DOCUMENT_OP_REFUSAL_REASONS.EMPTY_RECORD:
      return {
        reason: "unsupportedBlock",
        message,
        refusal: { gap: CANONICAL_GAP.publicUnsupportedInline },
      };
    default: {
      const unreachable: never = reason;
      return unreachable;
    }
  }
};
