import type { FolioAIEditOperation } from "./types";

/** Shared cutover ledger identifiers; noncanonical v1 consumers keep their existing executor. */
export const CANONICAL_PUBLIC_OPERATION_CAPABILITIES = {
  COMMENTS: "publicOps.comments",
  SUGGESTED_MODE: "publicOps.suggestedMode",
  TABLE_PROJECTION: "publicOps.tableProjection",
  UNSUPPORTED_INLINE: "publicOps.unsupportedInline",
  HEADLESS_SESSION: "publicOps.headlessSession",
} as const;

export type CanonicalPublicOperationCapabilityId =
  (typeof CANONICAL_PUBLIC_OPERATION_CAPABILITIES)[keyof typeof CANONICAL_PUBLIC_OPERATION_CAPABILITIES];

export type CanonicalPublicOperationRefusal = {
  gap: CanonicalPublicOperationCapabilityId;
};

/** A new public kind cannot bypass a canonical compiler/refusal decision. */
export const CANONICAL_PUBLIC_OPERATION_DISPOSITIONS = {
  replaceInBlock: "compile",
  replaceRange: "compile",
  replaceBlock: "compile",
  splitBlock: "compile",
  formatRange: "compile",
  mergeBlockWithNext: "compile",
  setBlockParagraphProperties: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.UNSUPPORTED_INLINE,
  insertAfterBlock: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.UNSUPPORTED_INLINE,
  insertBeforeBlock: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.UNSUPPORTED_INLINE,
  deleteBlock: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.UNSUPPORTED_INLINE,
  commentOnBlock: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.COMMENTS,
  commentOnRange: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.COMMENTS,
  insertTable: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  insertSignatureTable: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  deleteTable: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  insertTableRow: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  deleteTableRow: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  insertTableColumn: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  deleteTableColumn: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  mergeTableCells: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
  splitTableCell: CANONICAL_PUBLIC_OPERATION_CAPABILITIES.TABLE_PROJECTION,
} as const satisfies Record<
  FolioAIEditOperation["type"],
  "compile" | CanonicalPublicOperationCapabilityId
>;
