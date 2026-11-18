import { expect } from "bun:test";

import type { Document } from "../../model/document";
import { type AppliedDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpType,
} from "../types";
import { GENERATED_OP_KINDS, type OpSeed } from "./documentArbitraries";

export const NUM_RUNS = 10_000;

export type Tally = Map<DocumentOpType | "refused", number>;

export const count = (tally: Tally, key: DocumentOpType | "refused"): void => {
  tally.set(key, (tally.get(key) ?? 0) + 1);
};

/**
 * A property that passes because nothing applied proves nothing. Each such
 * property also runs on a fixed sample (test/property-coverage.ts) in which
 * every generated kind takes its turn, and every kind must apply there: what
 * the property's own seed drew, or how far its run got, decides nothing.
 */
export const COVERAGE_RUNS = GENERATED_OP_KINDS.length * 20;
/**
 * Sequences of two to eight operations. Each run's first operation takes its
 * turn too, so every kind also meets an unedited document as often as above.
 */
export const SEQUENCE_COVERAGE_RUNS = COVERAGE_RUNS;

/** The `index`th seed of a coverage sample, of the kind whose turn it is. */
export const inTurn = (seed: OpSeed, index: number): OpSeed => ({
  ...seed,
  kind: index % GENERATED_OP_KINDS.length,
});

export const expectEveryKindApplied = (tally: Tally): void => {
  for (const kind of GENERATED_OP_KINDS) {
    expect({ kind, applied: (tally.get(kind) ?? 0) > 0 }).toEqual({
      kind,
      applied: true,
    });
  }
};

/**
 * What each operation's inverse is made of: the table in `apply.ts`. Every
 * primary inverse is one operation, except a run patch's, which restores one stretch
 * of prior values per operation, and a split's or join's, which may give the
 * paragraph keeping its id its own review fields back and restore authored
 * content that the plain-run seam merge would otherwise coalesce. Note
 * lifecycle composition also restores owned bodies and exhausted threads.
 */
export const NOTE_LIFECYCLE_INVERSE_KINDS = {
  [DOCUMENT_OP_TYPES.REMOVE_NOTE]: [
    DOCUMENT_OP_TYPES.INSERT_CONTENT,
    DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
  ],
  [DOCUMENT_OP_TYPES.DELETE_BLOCKS]: [
    DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
    DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
  ],
  [DOCUMENT_OP_TYPES.INSERT_BLOCKS]: [
    DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
    DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
  ],
  [DOCUMENT_OP_TYPES.INSERT_CONTENT]: [DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE],
  [DOCUMENT_OP_TYPES.DELETE_RANGE]: [DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE],
  [DOCUMENT_OP_TYPES.REPLACE_INLINE]: [DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE],
  [DOCUMENT_OP_TYPES.RESOLVE_REVISION]: [
    DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
    DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
  ],
} as const satisfies Record<string, readonly DocumentOpType[]>;

export const INVERSE_KINDS = {
  setPackageResources: ["setPackageResources"],
  createComment: ["restoreCommentState"],
  updateCommentContent: ["restoreCommentState"],
  setCommentResolution: ["restoreCommentState"],
  deleteComment: ["restoreCommentState"],
  restoreCommentState: ["restoreCommentState"],
  createHeaderFooter: ["restoreStoryParts"],
  removeHeaderFooter: ["restoreStoryParts"],
  addNote: ["restoreStoryParts"],
  removeNote: ["restoreStoryParts", ...NOTE_LIFECYCLE_INVERSE_KINDS.removeNote],
  setSectionProps: ["restoreStoryParts"],
  restoreStoryParts: ["restoreStoryParts"],
  deleteBlocks: [
    "insertBlocks",
    "replaceBlocks",
    "replaceInline",
    "setParagraphReview",
    ...NOTE_LIFECYCLE_INVERSE_KINDS.deleteBlocks,
  ],
  insertBlocks: ["replaceBlocks", ...NOTE_LIFECYCLE_INVERSE_KINDS.insertBlocks],
  insertTable: ["setContainerBlocks"],
  deleteTable: ["setContainerBlocks", "replaceInline", "setParagraphReview"],
  setContainerBlocks: ["setContainerBlocks"],
  insertRow: ["setTableRows"],
  deleteRow: ["setTableRows", "replaceInline", "setParagraphReview"],
  setTableRows: ["setTableRows"],
  insertText: ["deleteRange", "replaceInline", "restoreStoryParts"],
  insertContent: [
    "deleteRange",
    "replaceInline",
    "restoreStoryParts",
    ...NOTE_LIFECYCLE_INVERSE_KINDS.insertContent,
  ],
  deleteRange: [
    "insertContent",
    "replaceInline",
    "restoreStoryParts",
    ...NOTE_LIFECYCLE_INVERSE_KINDS.deleteRange,
  ],
  splitInline: ["joinInline", "restoreStoryParts"],
  joinInline: ["splitInline", "restoreStoryParts"],
  setRunProps: ["setRunProps"],
  setParagraphProps: ["setParagraphProps", "setParagraphReview"],
  splitBlock: [
    "joinBlocks",
    "replaceInline",
    "setParagraphReview",
    "restoreStoryParts",
  ],
  joinBlocks: [
    "splitBlock",
    "replaceInline",
    "setParagraphReview",
    "restoreStoryParts",
  ],
  replaceBlocks: ["replaceBlocks", "restoreStoryParts"],
  setParagraphReview: ["setParagraphReview"],
  replaceInline: [
    "replaceInline",
    "restoreStoryParts",
    ...NOTE_LIFECYCLE_INVERSE_KINDS.replaceInline,
  ],
  resolveRevision: [
    "replaceInline",
    "setParagraphReview",
    "joinBlocks",
    "replaceBlocks",
    ...NOTE_LIFECYCLE_INVERSE_KINDS.resolveRevision,
  ],
  createNumberingInstance: ["deleteNumberingInstance"],
  deleteNumberingInstance: ["createNumberingInstance"],
  setSectionEndpoint: ["setSectionEndpoint"],
} as const satisfies Record<DocumentOpType, readonly DocumentOpType[]>;

export const applyAll = (
  document: Document,
  ops: readonly DocumentOp[],
): Document => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) {
    throw applied.error;
  }
  return applied.value.document;
};

export const expectRestores = (
  applied: AppliedDocumentOp,
  original: Document,
): void => {
  // An operation leaves the seed contract holding: checked afresh, not from memory.
  expect(contractViolation(applied.document)).toBeUndefined();
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) {
    throw restored.error;
  }
  expect(restored.value.document).toStrictEqual(original);
  expect(contractViolation(restored.value.document)).toBeUndefined();
  const replayed = JSON.parse(JSON.stringify(applied.inverse)) as DocumentOp[];
  expect(applyAll(applied.document, replayed)).toStrictEqual(original);
  expect(
    applyAll(restored.value.document, restored.value.inverse),
  ).toStrictEqual(applied.document);
};
