/**
 * Public vocabulary of {@link ./compare.compareDocx}: the change list, the
 * options that pin its output, and the failures it reports.
 */

import { TaggedError } from "better-result";

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import type {
  FolioAIBlockTableLocation,
  FolioAIEditSkippedOperation,
  FolioAIInlineFormatting,
} from "../ai-edits/types";

/** Everything {@link compareDocx} needs; nothing it reads from the ambient clock. */
export type CompareDocxOptions = {
  /** Author recorded on every generated tracked change. */
  author: string;
  /**
   * ISO-8601 date stamped on every generated tracked change. Required rather
   * than defaulted so a caller cannot get a nondeterministic package by
   * omission; pass the base document's own timestamp, a release date, or a
   * fixed epoch.
   */
  timestamp: string;
};

/** Where one change sits in the base or target document. */
export type CompareChangeLocation = {
  story: FolioDocumentStoryHandle;
  /** Innermost table cell, when the change is inside a table. */
  cell?: FolioAIBlockTableLocation;
};

/** One run of characters whose inline formatting differs, in base-block offsets. */
export type CompareFormatRange = {
  startOffset: number;
  endOffset: number;
  /** Only the properties that differ, set to the target document's value. */
  formatting: FolioAIInlineFormatting;
};

/**
 * One difference between the two documents, in target-document order with
 * base-only entries slotted where they sat. Every variant is plain JSON, so a
 * caller can hand the list to an agent instead of the package.
 *
 * A `move` is reported, not represented: the tracked-change grammar has no
 * "this paragraph came from there" mark that Word round-trips, so the buffer
 * carries a deletion at the source and an insertion at the destination while
 * the change list keeps the relocation visible.
 */
export type CompareChange =
  | {
      kind: "insert";
      location: CompareChangeLocation;
      /** Id of the inserted block in the target document. */
      targetBlockId: string;
      after: string;
    }
  | {
      kind: "delete";
      location: CompareChangeLocation;
      baseBlockId: string;
      before: string;
    }
  | {
      kind: "replace";
      location: CompareChangeLocation;
      baseBlockId: string;
      targetBlockId: string;
      before: string;
      after: string;
    }
  | {
      kind: "move";
      location: CompareChangeLocation;
      /** The block's id where it sat in the base document. */
      baseBlockId: string;
      /** The same content's id where it sits in the target document. */
      targetBlockId: string;
      text: string;
    }
  | {
      kind: "format";
      location: CompareChangeLocation;
      baseBlockId: string;
      targetBlockId: string;
      text: string;
      ranges: readonly CompareFormatRange[];
    }
  | {
      kind: "table-row-insert";
      location: CompareChangeLocation;
      tableIndex: number;
      /** Row index in the target document's table. */
      rowIndex: number;
      /** The new row's cell texts, in physical cell order. */
      cells: readonly string[];
      targetBlockIds: readonly string[];
    }
  | {
      kind: "table-row-delete";
      location: CompareChangeLocation;
      tableIndex: number;
      /** Row index in the base document's table. */
      rowIndex: number;
      cells: readonly string[];
      baseBlockIds: readonly string[];
    };

/** Why a part of the package is absent from `changes`. */
export const COMPARE_UNSUPPORTED_REASONS = Object.freeze([
  /** Header, footer, footnote, and endnote stories are out of scope this iteration. */
  "secondary-story",
  /** The story exists only in the target package; creating a part is not a text edit. */
  "story-missing-in-base",
  /** The story exists only in the base package. */
  "story-missing-in-target",
] as const);

export type CompareUnsupportedReason = (typeof COMPARE_UNSUPPORTED_REASONS)[number];

/**
 * A package part the comparison did not cover. Reported rather than dropped so
 * a caller can tell "no differences" from "not looked at".
 */
export type CompareUnsupportedPart = {
  reason: CompareUnsupportedReason;
  baseStory: FolioDocumentStoryHandle | null;
  targetStory: FolioDocumentStoryHandle | null;
};

export type CompareResult = {
  /** The base package carrying the generated tracked changes. */
  buffer: ArrayBuffer;
  changes: readonly CompareChange[];
  unsupported: readonly CompareUnsupportedPart[];
};

export class InvalidCompareDocxOptionsError extends TaggedError("InvalidCompareDocxOptionsError")<{
  message: string;
  option: "timestamp";
  receivedValue: unknown;
}> {}

export class CompareDocxParseError extends TaggedError("CompareDocxParseError")<{
  message: string;
  side: "base" | "target";
  cause: unknown;
}> {}

/**
 * The applier refused at least one derived operation, so accepting the result
 * would not reproduce the target. Reported instead of returning a package that
 * silently under-represents the difference.
 */
export class CompareDocxApplyError extends TaggedError("CompareDocxApplyError")<{
  message: string;
  skipped: readonly FolioAIEditSkippedOperation[];
}> {}

/**
 * The generated tracked changes do not accept back to the target. Raised by
 * the self-check {@link compareDocx} runs before returning: a redline that
 * quietly loses part of the difference is worse than no redline, so the
 * mismatch is surfaced instead of the document.
 */
export class CompareDocxRoundTripError extends TaggedError("CompareDocxRoundTripError")<{
  message: string;
  story: FolioDocumentStoryHandle;
  /** Block texts the accepted result holds where the target differs. */
  acceptedText: readonly string[];
  targetText: readonly string[];
}> {}

/** The difference needs more operations than the engine will generate. */
export class CompareDocxOperationLimitError extends TaggedError("CompareDocxOperationLimitError")<{
  message: string;
  limit: number;
}> {}

export class CompareDocxSerializeError extends TaggedError("CompareDocxSerializeError")<{
  message: string;
  cause: unknown;
}> {}

export type CompareDocxError =
  | CompareDocxApplyError
  | CompareDocxOperationLimitError
  | CompareDocxParseError
  | CompareDocxRoundTripError
  | CompareDocxSerializeError
  | InvalidCompareDocxOptionsError;
