/**
 * Public vocabulary of {@link ./compare.compareDocx}: the change list, the
 * options that pin its output, and the failures it reports.
 */

import { TaggedError } from "better-result";

import type { FolioDocumentStoryHandle, FolioNumberingLevel } from "../ai-edits/headless";
import type { WordDiffGranularity } from "./text-diff";
import type { FolioAIBlockParagraphProperties, FolioAIBlockTableLocation } from "../ai-edits/types";
import type { FolioContentInlineFormattingChange } from "./content-types";
import type { FolioContentComparisonError } from "./content";
import type { CompareTableFormatDetails } from "./table-format-properties";
import type {
  CompareVerification,
  CompareVerificationCause,
  CompareVerificationFailure,
  CompareVerificationInvariant,
  CompareVerificationScope,
  FinalParagraphMarkRevision,
} from "./verification";

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
  /** Strict refuses any unsupported or unverified difference; it is the default. */
  mode?: "strict" | "bestEffort";
  /**
   * Token size a changed paragraph's redline is cut at: `"word"` (default)
   * marks whole words, `"character"` marks the changed letters inside one.
   *
   * Case and whitespace normalization are not options here, though
   * `diffWordSegments` offers them: a comparison that leaves a difference
   * unmarked does not accept back to the target, and the round trip is the
   * one thing this call promises.
   */
  granularity?: WordDiffGranularity;
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
  /** Exact authored and effective property deltas for this paired range. */
  formatting: FolioContentInlineFormattingChange;
};

export type {
  CompareTableCellCoordinate,
  CompareTableCellFormattingPropertyChange,
  CompareTableCellFormattingPropertyName,
  CompareTableCoordinate,
  CompareTableFormatDetails,
  CompareTableFormattingPropertyChange,
  CompareTableFormattingPropertyName,
  CompareTableRowCoordinate,
  CompareTableRowFormattingPropertyChange,
  CompareTableRowFormattingPropertyName,
} from "./table-format-properties";

export type CompareTableFormatChange = CompareTableFormatDetails & {
  readonly kind: "table-format";
  readonly location: CompareChangeLocation;
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
  /**
   * One paragraph became two: a paragraph mark was inserted and no words
   * changed. Reported as its own kind so a reader is not told the tail was
   * newly written.
   */
  | {
      kind: "split";
      location: CompareChangeLocation;
      baseBlockId: string;
      /** The two blocks the base block became, in target order. */
      targetBlockIds: readonly string[];
      text: string;
    }
  /** Two paragraphs became one: a paragraph mark was deleted. */
  | {
      kind: "merge";
      location: CompareChangeLocation;
      /** The two blocks that became one, in base order. */
      baseBlockIds: readonly string[];
      targetBlockId: string;
      text: string;
    }
  /**
   * A paragraph property moved and no words did: a list item demoted a level,
   * a paragraph restyled. Written as `w:pPrChange`, so rejecting restores the
   * whole previous property set the way Word does.
   */
  | {
      kind: "paragraph-format";
      location: CompareChangeLocation;
      baseBlockId: string;
      targetBlockId: string;
      /** Only the properties that differ, set to the target document's value. */
      properties: FolioAIBlockParagraphProperties;
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
  /**
   * A numbering definition that differs. It carries no `location`: numbering
   * lives in the package, not in a story, and one definition governs every
   * list that references it.
   *
   * Reported and not represented. A renumbering that FOLLOWS from an edit —
   * an item inserted, so the ones below it count on — is already shown
   * as-if-accepted, because labels are rendered from these definitions rather
   * than stored on the paragraphs. A definition that itself changed is a
   * different thing, and OOXML has no tracked-change grammar for it: Word
   * does not track `numbering.xml` either.
   */
  | {
      kind: "numbering";
      numId: number;
      level: number;
      /** `null` when the target added this level. */
      before: FolioNumberingLevel | null;
      /** `null` when the target dropped it. */
      after: FolioNumberingLevel | null;
    }
  /** A whole table the target added. */
  | {
      kind: "table-insert";
      location: CompareChangeLocation;
      tableIndex: number;
      /** Cell texts row by row, in physical cell order. */
      rows: readonly (readonly string[])[];
      targetBlockIds: readonly string[];
    }
  /** A whole table the target dropped. */
  | {
      kind: "table-delete";
      location: CompareChangeLocation;
      tableIndex: number;
      rows: readonly (readonly string[])[];
      baseBlockIds: readonly string[];
    }
  | {
      kind: "table-row-delete";
      location: CompareChangeLocation;
      tableIndex: number;
      /** Row index in the base document's table. */
      rowIndex: number;
      cells: readonly string[];
      baseBlockIds: readonly string[];
    }
  | {
      kind: "table-column-insert";
      location: CompareChangeLocation;
      tableIndex: number;
      /** Grid-column index in the target table. */
      columnIndex: number;
      /** Newly created physical-cell text in row order. */
      cells: readonly string[];
      targetBlockIds: readonly string[];
    }
  | {
      kind: "table-column-delete";
      location: CompareChangeLocation;
      tableIndex: number;
      /** Grid-column index in the base table. */
      columnIndex: number;
      cells: readonly string[];
      baseBlockIds: readonly string[];
    }
  /** A tracked table, row, or cell property change. */
  | CompareTableFormatChange;

/** Why a part of the package is absent from `changes`. */
export const COMPARE_UNSUPPORTED_REASONS = Object.freeze([
  /** The story exists only in the target package; creating a part is not a text edit. */
  "story-missing-in-base",
  /** The story exists only in the base package. */
  "story-missing-in-target",
  /** The story is present on both sides but carries no editable state. */
  "story-not-editable",
  /** A modeled block property has no tracked-document encoding. */
  "block-semantics",
  /** The block changed structural container without a lossless relocation encoding. */
  "container-change",
  /** A zero-width inline structural boundary cannot be transported losslessly. */
  "structural-boundary-change",
  /** Resolved paragraph presentation changed without an authored property change. */
  "effective-paragraph-formatting",
  /** Resolved run presentation changed without an authored property change. */
  "effective-inline-formatting",
  /** A complete target table cannot be copied losslessly into the base package. */
  "nonportable-table-template",
  /** A referenced numbering definition changed but has no tracked-change grammar. */
  "numbering-definition",
  /** A canonical event could not be resolved into a proved tracked-document instruction. */
  "transport-preflight",
] as const);

export type CompareUnsupportedReason = (typeof COMPARE_UNSUPPORTED_REASONS)[number];

/**
 * A package part the comparison did not cover. Reported rather than dropped so
 * a caller can tell "no differences" from "not looked at".
 */
type CompareUnsupportedStoryReason = Extract<
  CompareUnsupportedReason,
  "story-missing-in-base" | "story-missing-in-target" | "story-not-editable"
>;

type CompareUnsupportedContentReason = Exclude<
  CompareUnsupportedReason,
  CompareUnsupportedStoryReason | "numbering-definition" | "transport-preflight"
>;

/** Every bounded preflight disposition for a canonical DOCX instruction. */
export const COMPARE_DOCX_PREFLIGHT_REASONS = Object.freeze([
  "missing-block",
  "changed-block",
  "source-expectation-mismatch",
  "missing-anchor",
  "pending-paragraph-change",
  "pending-run-change",
  "source-formatting-mismatch",
  "unrepresentable-text-range",
  "unrepresentable-paragraph-boundary",
  "unrepresentable-table-geometry",
  "unrepresentable-table-structure",
  /** A sibling required by the same semantic change failed preflight. */
  "semantic-group-incomplete",
] as const);

export type CompareDocxPreflightReason = (typeof COMPARE_DOCX_PREFLIGHT_REASONS)[number];

export type CompareUnsupportedPart =
  | {
      readonly reason: CompareUnsupportedStoryReason;
      readonly baseStory: FolioDocumentStoryHandle | null;
      readonly targetStory: FolioDocumentStoryHandle | null;
    }
  | {
      readonly reason: CompareUnsupportedContentReason;
      readonly story: FolioDocumentStoryHandle;
      readonly eventType:
        | "unchanged"
        | "modified"
        | "formatting"
        | "inserted"
        | "deleted"
        | "moved"
        | "split"
        | "merge"
        | "tableReplacement"
        | "structural";
      readonly field?: string;
      readonly baseBlockId?: string;
      readonly targetBlockId?: string;
      readonly tableIndex?: number;
    }
  | {
      readonly reason: "numbering-definition";
      readonly numId: number;
      readonly level: number;
    }
  | {
      readonly reason: "transport-preflight";
      readonly story: FolioDocumentStoryHandle;
      readonly instructionIndex: number;
      readonly detail: CompareDocxPreflightReason;
      readonly blockId?: string;
    };

export type CompareResult = {
  /** The base package carrying the generated tracked changes. */
  buffer: ArrayBuffer;
  changes: readonly CompareChange[];
  /**
   * Whether the round trip was proven. Always `verified` in strict mode;
   * best-effort mode may return an unverified but explicitly bounded result.
   */
  verification: CompareVerification;
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

/** A paired story exceeded or violated the neutral comparison contract. */
export class CompareDocxContentComparisonError extends TaggedError(
  "CompareDocxContentComparisonError",
)<{
  message: string;
  story: FolioDocumentStoryHandle;
  cause: FolioContentComparisonError;
}> {}

export const COMPARE_DOCX_EXECUTION_REASONS = Object.freeze([
  "stale-preflight",
  "invalid-revision-stamp",
  "table-geometry-execution",
] as const);

export type CompareDocxExecutionReason = (typeof COMPARE_DOCX_EXECUTION_REASONS)[number];

/** A preflighted story could not complete its atomic comparison transaction. */
export class CompareDocxApplyError extends TaggedError("CompareDocxApplyError")<{
  message: string;
  story: FolioDocumentStoryHandle;
  reason: CompareDocxExecutionReason;
}> {}

/**
 * The generated tracked changes do not accept back to the target. Raised by
 * the self-check {@link compareDocx} runs before returning: a redline that
 * quietly loses part of the difference is worse than no redline, so the
 * mismatch is surfaced instead of the document.
 */
export class CompareDocxRoundTripError extends TaggedError("CompareDocxRoundTripError")<{
  message: string;
  scope: CompareVerificationScope;
  /** The invariant that did not hold, and what diverged under it. */
  invariant: CompareVerificationInvariant;
  cause: CompareVerificationCause;
  /**
   * Every invariant that failed, not only the one named above: a redline that
   * loses a difference usually loses it in both directions, and a caller
   * deciding what to do next wants the whole list. Each detail is structural,
   * so it is safe to log or quote.
   */
  failures: readonly CompareVerificationFailure[];
}> {}

/** The difference needs more operations than the engine will generate. */
export class CompareDocxOperationLimitError extends TaggedError("CompareDocxOperationLimitError")<{
  message: string;
  limit: number;
}> {}

/** Strict mode found differences with no proved tracked-document lowering. */
export class CompareDocxUnsupportedError extends TaggedError("CompareDocxUnsupportedError")<{
  message: string;
  unsupported: readonly CompareUnsupportedPart[];
}> {}

export const COMPARE_DOCX_LOWERING_REASONS = Object.freeze([
  "block-semantics",
  "container-change",
  "missing-insertion-anchor",
  "missing-removal-boundary",
  "nonportable-table-template",
  "structural-boundary-change",
  "table-row-anchor",
] as const);

export type CompareDocxLoweringReason = (typeof COMPARE_DOCX_LOWERING_REASONS)[number];

/** A canonical content event has no lossless tracked-document encoding. */
export class CompareDocxLoweringError extends TaggedError("CompareDocxLoweringError")<{
  message: string;
  reason: CompareDocxLoweringReason;
  story: FolioDocumentStoryHandle;
  baseBlockId?: string;
  targetBlockId?: string;
  tableIndex?: number;
}> {}

export class CompareDocxSerializeError extends TaggedError("CompareDocxSerializeError")<{
  message: string;
  cause: unknown;
}> {}

/**
 * A container's final paragraph mark carries a revision, so the package would
 * not open, or would open carrying one no reader can resolve.
 *
 * A deleted paragraph mark means "merge this paragraph into the following
 * one", an inserted one means the break was added and rejecting it closes the
 * paragraph back over the next one, and a container's last paragraph has no
 * following one either way. Checked before the package is written, and fatal
 * in either comparison mode: there is no redline to emit when a
 * consumer refuses the file.
 */
export class CompareDocxFinalParagraphMarkError extends TaggedError(
  "CompareDocxFinalParagraphMarkError",
)<{
  message: string;
  /** Every container that carries one, each named structurally. */
  revisions: readonly FinalParagraphMarkRevision[];
}> {}

export type CompareDocxError =
  | CompareDocxApplyError
  | CompareDocxContentComparisonError
  | CompareDocxFinalParagraphMarkError
  | CompareDocxLoweringError
  | CompareDocxOperationLimitError
  | CompareDocxParseError
  | CompareDocxRoundTripError
  | CompareDocxSerializeError
  | CompareDocxUnsupportedError
  | InvalidCompareDocxOptionsError;
