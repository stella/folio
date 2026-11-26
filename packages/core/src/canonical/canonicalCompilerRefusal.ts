import {
  DOCUMENT_OP_REFUSAL_REASONS,
  type DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "@stll/docx-core/ops";

import type { FolioAIEditSkipReason } from "../ai-edits/types";
import type { CanonicalRefusalGround } from "../types/canonicalCapabilities";
import { groundForDocumentOpRefusal } from "./canonicalRefusalGrounds";

type CompilerRefusalSkip = {
  reason: FolioAIEditSkipReason;
  message: string;
  refusal?: CanonicalRefusalGround;
};

/**
 * How the canonical session reports a compiler refusal: `typed`, a
 * deliberate refusal the document's review records call for, under its own
 * public reason and no ground, as the legacy engine reports one; or
 * `grounded`, under the public reason that fits it and the ground its
 * document operation reason has (`groundForDocumentOpRefusal`).
 */
type CompilerRefusalDisposition =
  | { type: "typed"; reason: FolioAIEditSkipReason }
  | { type: "grounded"; reason: FolioAIEditSkipReason };

const typed = (reason: FolioAIEditSkipReason) =>
  ({ type: "typed", reason }) as const;
const grounded = (reason: FolioAIEditSkipReason) =>
  ({ type: "grounded", reason }) as const;
const UNSUPPORTED = grounded("unsupportedBlock");

const R = DOCUMENT_OP_REFUSAL_REASONS;

/**
 * Every compiler refusal reason's public disposition. A deliberate refusal
 * names its own reason, as the legacy engine reports one, and so does a
 * change that would change nothing. A document that moved on is a stale
 * range, and an offset inside a character splits it. Every other reason
 * reports its ground besides: `untrackable` the shapes the tracked-change
 * records do not express yet, `unrecordableLinkSuggestion` the rule the
 * review records state, and the rest a request the document does not allow.
 */
const COMPILER_REFUSAL_DISPOSITIONS = {
  [R.PENDING_RUN_PROPERTY_CHANGE]: typed("pendingRunPropertyChange"),
  [R.PENDING_PARAGRAPH_PROPERTY_CHANGE]: typed(
    "pendingParagraphPropertyChange",
  ),
  [R.JOIN_INTO_PENDING_TABLE]: typed("pendingTableInsertion"),
  [R.PENDING_TABLE_CHANGE]: typed("pendingTableChange"),
  [R.PENDING_GRID_CHANGE]: typed("pendingTableChange"),
  [R.PENDING_BREAK_TAKEOVER]: typed("pendingParagraphMarkDeletion"),
  [R.PENDING_PARAGRAPH_MARK_DELETION]: typed("pendingParagraphMarkDeletion"),
  [R.PENDING_PARAGRAPH_MARK_INSERTION]: typed("pendingParagraphMarkInsertion"),
  // Both engines refuse a public property edit of such a paragraph with this
  // reason (PENDING_PARAGRAPH_MARK_PROPERTY_REFUSAL).
  [R.PENDING_PARAGRAPH_JOIN]: typed("pendingParagraphMarkDeletion"),
  [R.UNSUPPORTED_SCHEMA]: UNSUPPORTED,
  [R.BLOCK_NOT_FOUND]: grounded("missingBlock"),
  [R.INVALID_OPERATION]: UNSUPPORTED,
  [R.NO_CHANGE]: typed("noopOperation"),
  [R.DEPENDENT_RECORDS]: UNSUPPORTED,
  [R.INVALID_OFFSET]: UNSUPPORTED,
  [R.CROSS_BLOCK_RANGE]: UNSUPPORTED,
  [R.SPLITS_SURROGATE_PAIR]: grounded("splitsCharacter"),
  [R.INVALID_TEXT]: UNSUPPORTED,
  [R.INVALID_BLOCK_ID]: UNSUPPORTED,
  [R.ID_COLLISION]: UNSUPPORTED,
  [R.NOT_ADJACENT]: UNSUPPORTED,
  [R.SECTION_BOUNDARY]: UNSUPPORTED,
  [R.INSIDE_TRACKED_DELETION]: UNSUPPORTED,
  [R.NEEDS_NEW_IDS]: UNSUPPORTED,
  [R.INVALID_NEW_ID]: UNSUPPORTED,
  [R.SHARED_ID]: UNSUPPORTED,
  [R.STALE]: grounded("staleRange"),
  [R.STRUCTURE_MISMATCH]: UNSUPPORTED,
  [R.EMPTY_CONTENT]: UNSUPPORTED,
  [R.EMPTY_BLOCK_LIST]: UNSUPPORTED,
  [R.REVISION_CONFLICT]: UNSUPPORTED,
  [R.CONTAINER_FINAL_MARK]: UNSUPPORTED,
  [R.INSIDE_INLINE_CONTROL]: UNSUPPORTED,
  [R.UNTRACKABLE]: UNSUPPORTED,
  [R.UNRECORDABLE_LINK_SUGGESTION]: UNSUPPORTED,
  [R.MISSING_BLOCK_ID]: UNSUPPORTED,
  [R.DUPLICATE_BLOCK_ID]: UNSUPPORTED,
  [R.DUPLICATE_RECORD_ID]: UNSUPPORTED,
  [R.SECTIONS_OUT_OF_STEP]: grounded("staleRange"),
  [R.EMPTY_RECORD]: UNSUPPORTED,
} as const satisfies Record<
  DocumentOpRefusalReason,
  CompilerRefusalDisposition
>;

/** The skip a refused compilation reports. */
export const compilerRefusalSkip = ({
  reason,
  message,
}: DocumentOpRefusal): CompilerRefusalSkip => {
  const disposition = COMPILER_REFUSAL_DISPOSITIONS[reason];
  switch (disposition.type) {
    case "typed":
      return { reason: disposition.reason, message };
    case "grounded":
      return {
        reason: disposition.reason,
        message,
        refusal: groundForDocumentOpRefusal(reason),
      };
    default: {
      const unreachable: never = disposition;
      return unreachable;
    }
  }
};
