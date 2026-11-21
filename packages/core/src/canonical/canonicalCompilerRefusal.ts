import {
  DOCUMENT_OP_REFUSAL_REASONS,
  type DocumentOpRefusal,
  type DocumentOpRefusalReason,
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
 * How the canonical session reports a compiler refusal: `typed`, a
 * deliberate refusal the document's review records call for, under its own
 * public reason; or `unsupported`, a shape the engine cannot represent yet
 * (or a compiler output its model refuses), a capability gap.
 */
type CompilerRefusalDisposition =
  { type: "typed"; reason: FolioAIEditSkipReason } | { type: "unsupported" };

const typed = (reason: FolioAIEditSkipReason) =>
  ({ type: "typed", reason }) as const;
const UNSUPPORTED = { type: "unsupported" } as const;

const R = DOCUMENT_OP_REFUSAL_REASONS;

/**
 * Every compiler refusal reason's public disposition. A deliberate refusal
 * names its own reason, as the legacy engine reports one; `untrackable`
 * gathers shapes the engine does not record yet, and `revisionConflict` a
 * tracked split the compiler stated with records of its own, so both report
 * a gap.
 */
const COMPILER_REFUSAL_DISPOSITIONS = {
  [R.PENDING_RUN_PROPERTY_CHANGE]: typed("pendingRunPropertyChange"),
  [R.PENDING_PARAGRAPH_PROPERTY_CHANGE]: typed(
    "pendingParagraphPropertyChange",
  ),
  [R.JOIN_INTO_PENDING_TABLE]: typed("pendingTableInsertion"),
  [R.PENDING_TABLE_CHANGE]: typed("pendingTableChange"),
  [R.PENDING_BREAK_TAKEOVER]: typed("pendingParagraphMarkDeletion"),
  [R.PENDING_PARAGRAPH_MARK]: typed("pendingParagraphMarkDeletion"),
  [R.UNSUPPORTED_SCHEMA]: UNSUPPORTED,
  [R.BLOCK_NOT_FOUND]: UNSUPPORTED,
  [R.INVALID_OPERATION]: UNSUPPORTED,
  [R.NO_CHANGE]: UNSUPPORTED,
  [R.DEPENDENT_RECORDS]: UNSUPPORTED,
  [R.INVALID_OFFSET]: UNSUPPORTED,
  [R.CROSS_BLOCK_RANGE]: UNSUPPORTED,
  [R.SPLITS_SURROGATE_PAIR]: UNSUPPORTED,
  [R.INVALID_TEXT]: UNSUPPORTED,
  [R.INVALID_BLOCK_ID]: UNSUPPORTED,
  [R.ID_COLLISION]: UNSUPPORTED,
  [R.NOT_ADJACENT]: UNSUPPORTED,
  [R.SECTION_BOUNDARY]: UNSUPPORTED,
  [R.INSIDE_TRACKED_DELETION]: UNSUPPORTED,
  [R.NEEDS_NEW_IDS]: UNSUPPORTED,
  [R.INVALID_NEW_ID]: UNSUPPORTED,
  [R.SHARED_ID]: UNSUPPORTED,
  [R.STALE]: UNSUPPORTED,
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
  [R.SECTIONS_OUT_OF_STEP]: UNSUPPORTED,
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
    case "unsupported":
      return {
        reason: "unsupportedBlock",
        message,
        refusal: { gap: CANONICAL_GAP.publicUnsupportedInline },
      };
    default: {
      const unreachable: never = disposition;
      return unreachable;
    }
  }
};
