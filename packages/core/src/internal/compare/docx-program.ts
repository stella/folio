import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import type {
  FolioAIBlockParagraphProperties,
  FolioAIBlockTableLocation,
} from "../../ai-edits/types";
import type { TextFormatting } from "../../types/document";
import type {
  FolioContentFormattingChange,
  FolioContentRangePairRelation,
  FolioContentStructuralChange,
  FolioContentTextSegment,
  FolioContentWholePairRelation,
} from "../../compare/content";
import type {
  FolioContentBlock,
  FolioContentParagraphInsertionBoundary,
} from "../../compare/content-types";
import type { CompareChange } from "../../compare/types";
import { groupFolioContentTableRows } from "../../compare/content-alignment";
import {
  docxParagraphChangedProperties,
  docxParagraphPropertiesEqual,
  docxParagraphPropertiesFromBlock,
  docxTableLocationFromContent,
} from "./docx-paragraph-transport";
import {
  resolvedDocxAuthoredRunsForBlock,
  resolvedDocxAuthoredRunsForRange,
  resolvedDocxSourceOperand,
  resolvedDocxSourceOperandBlock,
  resolvedDocxSourceOperandSnapshot,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  resolvedDocxDeletedEventOperandPayload,
  resolvedDocxFormattingRangeOperand,
  resolvedDocxFormattingRangeOperandPayload,
  resolvedDocxInsertedEventOperandPayload,
  resolvedDocxMergeEventOperandPayload,
  resolvedDocxMoveEventOperandPayload,
  resolvedDocxPairedEventOperandPayload,
  resolvedDocxPairRangeOperand,
  resolvedDocxPairRangeOperandRelation,
  resolvedDocxReplacementRangeOperandPayload,
  resolvedDocxSeparatorOperandRelation,
  resolvedDocxSeparatorOperand,
  resolvedDocxSplitEventOperandPayload,
  resolvedDocxStoryComparisonPayload,
  resolvedDocxTerminalReplacementOperandPayload,
  resolvedDocxTableFormatOperandPayload,
  resolvedDocxTableStructureOperandPayload,
  resolvedDocxTableStructureReportOwners,
  resolvedDocxTargetBlockOperand,
  resolvedDocxTargetBlockOperandBlock,
  resolvedDocxTrailingDeletionOperandPayload,
  resolvedDocxWholeBlockReplacementOperand,
  type ResolvedDocxDeletedEventOperand,
  type ResolvedDocxFormattingRangeOperand,
  type ResolvedDocxInsertedEventOperand,
  type ResolvedDocxMergeEventOperand,
  type ResolvedDocxMoveEventOperand,
  type ResolvedDocxPairedEventOperand,
  type ResolvedDocxPairRangeOperand,
  type ResolvedDocxSeparatorOperand,
  type ResolvedDocxStoryComparison,
  type ResolvedDocxSplitEventOperand,
  type ResolvedDocxTerminalReplacementOperand,
  type ResolvedDocxTableFormatOperand,
  type ResolvedDocxTableStructureOperand,
  type ResolvedDocxTargetBlockOperand,
  type ResolvedDocxTrailingDeletionOperand,
  type ResolvedDocxWholeBlockReplacementOperand,
} from "./resolved-docx-story-comparison";

const MAX_DOCX_COMPARISON_INSTRUCTIONS = 10_000;

export type DocxAuthoredRun = {
  readonly startOffset: number;
  readonly endOffset: number;
  readonly formatting: Readonly<TextFormatting>;
};

export type DocxAuthoredChangeRange = {
  readonly baseStart: number;
  readonly baseEnd: number;
  readonly revisedStart: number;
  readonly revisedEnd: number;
  /** Canonical authored-property changes; effective-only changes are absent. */
  readonly properties: readonly string[];
};

type DocxComparisonRangePlanInput = {
  readonly sourceText: string;
  readonly targetText: string;
  readonly segments: readonly FolioContentTextSegment[];
  readonly sourceRuns: readonly DocxAuthoredRun[];
  readonly targetRuns: readonly DocxAuthoredRun[];
  readonly authoredChanges: readonly DocxAuthoredChangeRange[];
};

export type DocxComparisonSourceOperand = ResolvedDocxSourceOperand;

export type DocxComparisonSourceOperandGroup = readonly [
  DocxComparisonSourceOperand,
  ...DocxComparisonSourceOperand[],
];

export type DocxComparisonParagraphInsertionBoundary = {
  readonly type: "afterParagraph" | "beforeParagraph";
  readonly paragraph: DocxComparisonSourceOperand;
};

export type DocxComparisonParagraphRemovalBoundary =
  | {
      readonly type: "successorParagraph";
      readonly successor: DocxComparisonSourceOperand;
    }
  | {
      readonly type: "successorTable";
      readonly firstBlock: DocxComparisonSourceOperand;
    };

/**
 * Closed transport vocabulary for one DOCX story. These are not generic edit
 * requests: each branch owns its source expectation and complete accepted
 * payload, so execution never joins an operation to a second semantic record.
 */
export type DocxComparisonInstructionInput =
  | {
      readonly type: "replaceText";
      readonly range: ResolvedDocxPairRangeOperand | ResolvedDocxWholeBlockReplacementOperand;
    }
  | {
      readonly type: "formatText";
      readonly range: ResolvedDocxFormattingRangeOperand;
    }
  | {
      readonly type: "insertParagraph";
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: ResolvedDocxTargetBlockOperand;
    }
  | {
      readonly type: "deleteParagraph";
      readonly source: DocxComparisonSourceOperand;
    }
  | {
      readonly type: "moveParagraph";
      readonly source: DocxComparisonSourceOperand;
      readonly removalBoundary: DocxComparisonParagraphRemovalBoundary;
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: ResolvedDocxTargetBlockOperand;
    }
  | {
      readonly type: "moveTerminalParagraph";
      readonly predecessor: DocxComparisonSourceOperand;
      readonly source: DocxComparisonSourceOperand;
      readonly carrierTarget: ResolvedDocxTargetBlockOperand;
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: ResolvedDocxTargetBlockOperand;
    }
  | {
      readonly type: "splitParagraph";
      readonly first: ResolvedDocxPairRangeOperand;
      readonly second: ResolvedDocxPairRangeOperand;
      readonly separator: ResolvedDocxSeparatorOperand;
    }
  | {
      readonly type: "mergeParagraphs";
      readonly first: ResolvedDocxPairRangeOperand;
      readonly second: ResolvedDocxPairRangeOperand;
      readonly separator: ResolvedDocxSeparatorOperand;
    }
  | {
      readonly type: "deleteTrailingParagraphs";
      readonly chainStart: DocxComparisonSourceOperand;
      /** Paragraphs removed after the chain start, in base document order. */
      readonly deleted: DocxComparisonSourceOperandGroup;
    }
  | {
      readonly type: "setParagraphProperties";
      readonly source: DocxComparisonSourceOperand;
      readonly target: ResolvedDocxTargetBlockOperand;
    }
  | {
      readonly type: "tableStructure";
      readonly operation: ResolvedDocxTableStructureOperand;
    }
  | {
      readonly type: "tableFormat";
      readonly operation: ResolvedDocxTableFormatOperand;
    };

export type DocxComparisonReportInput = {
  /** Canonical comparison-stream order, independent of execution scheduling. */
  readonly sequence: number;
  readonly change: CompareChange;
};

type NonEmptyReadonlyArray<Value> = readonly [Value, ...Value[]];

/**
 * Closed semantic inputs for one DOCX comparison. Non-table branches own an
 * exact canonical comparison event; callers cannot supply reports or rebuild
 * an operation from independently selected blocks and ranges.
 */
export type DocxComparisonOperationInput =
  | {
      readonly type: "pairedBlock";
      readonly event: ResolvedDocxPairedEventOperand;
    }
  | {
      readonly type: "insertParagraph";
      readonly event: ResolvedDocxInsertedEventOperand;
    }
  | {
      readonly type: "deleteParagraph";
      readonly event: ResolvedDocxDeletedEventOperand;
    }
  | {
      readonly type: "moveParagraph";
      readonly event: ResolvedDocxMoveEventOperand;
    }
  | {
      readonly type: "splitParagraph";
      readonly event: ResolvedDocxSplitEventOperand;
    }
  | {
      readonly type: "mergeParagraphs";
      readonly event: ResolvedDocxMergeEventOperand;
    }
  | {
      readonly type: "deleteTrailingParagraphs";
      readonly operation: ResolvedDocxTrailingDeletionOperand;
    }
  | {
      readonly type: "replaceTerminalParagraph";
      readonly operation: ResolvedDocxTerminalReplacementOperand;
    }
  | {
      readonly type: "tableStructure";
      readonly operation: ResolvedDocxTableStructureOperand;
    }
  | {
      readonly type: "tableFormat";
      readonly operation: ResolvedDocxTableFormatOperand;
    };

export type DocxComparisonEqualFragment = {
  readonly type: "equal";
  readonly text: string;
  readonly baseStart: number;
  readonly baseEnd: number;
  readonly revisedStart: number;
  readonly revisedEnd: number;
  readonly sourceFormatting: Readonly<TextFormatting>;
  readonly targetFormatting: Readonly<TextFormatting>;
  /** Empty means the canonical comparison says authorship is unchanged. */
  readonly changedProperties: readonly string[];
};

export type DocxComparisonDeletedFragment = {
  readonly type: "del";
  readonly text: string;
  readonly baseStart: number;
  readonly baseEnd: number;
  readonly revisedStart: number;
  readonly revisedEnd: number;
  readonly sourceFormatting: Readonly<TextFormatting>;
};

export type DocxComparisonInsertedFragment = {
  readonly type: "ins";
  readonly text: string;
  readonly baseStart: number;
  readonly baseEnd: number;
  readonly revisedStart: number;
  readonly revisedEnd: number;
  readonly targetFormatting: Readonly<TextFormatting>;
};

export type DocxComparisonTextFragment =
  | DocxComparisonEqualFragment
  | DocxComparisonDeletedFragment
  | DocxComparisonInsertedFragment;

export type DocxComparisonRangePlan = {
  readonly sourceText: string;
  readonly targetText: string;
  readonly fragments: readonly DocxComparisonTextFragment[];
};

export type DocxComparisonParagraphTarget = {
  readonly text: string;
  readonly runs: readonly DocxAuthoredRun[];
  readonly properties: Readonly<FolioAIBlockParagraphProperties>;
  readonly table?: Readonly<FolioAIBlockTableLocation>;
};

type DocxComparisonInstructionPayload =
  | {
      readonly type: "replaceText";
      readonly source: DocxComparisonSourceOperand;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlan;
    }
  | {
      readonly type: "formatText";
      readonly source: DocxComparisonSourceOperand;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlan;
    }
  | {
      readonly type: "insertParagraph";
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "deleteParagraph";
      readonly source: DocxComparisonSourceOperand;
    }
  | {
      readonly type: "moveParagraph";
      readonly source: DocxComparisonSourceOperand;
      readonly removalBoundary: DocxComparisonParagraphRemovalBoundary;
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "moveTerminalParagraph";
      readonly predecessor: DocxComparisonSourceOperand;
      readonly source: DocxComparisonSourceOperand;
      readonly carrierTargetProperties: Readonly<FolioAIBlockParagraphProperties>;
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "splitParagraph";
      readonly source: DocxComparisonSourceOperand;
      readonly offset: number;
      readonly first: DocxComparisonRangePlan;
      readonly second: DocxComparisonRangePlan;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
      readonly firstTarget: DocxComparisonParagraphTarget;
      readonly secondTarget: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "mergeParagraphs";
      readonly firstSource: DocxComparisonSourceOperand;
      readonly secondSource: DocxComparisonSourceOperand;
      readonly first: DocxComparisonRangePlan;
      readonly second: DocxComparisonRangePlan;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
      readonly target: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "deleteTrailingParagraphs";
      readonly chainStart: DocxComparisonSourceOperand;
      readonly deleted: DocxComparisonSourceOperandGroup;
    }
  | {
      readonly type: "setParagraphProperties";
      readonly source: DocxComparisonSourceOperand;
      readonly targetProperties: Readonly<FolioAIBlockParagraphProperties>;
    }
  | {
      readonly type: "tableStructure";
      readonly operation: ResolvedDocxTableStructureOperand;
    }
  | {
      readonly type: "tableFormat";
      readonly operation: ResolvedDocxTableFormatOperand;
    };

type WithSemanticGroup<Instruction> = Instruction extends unknown
  ? Instruction & { readonly semanticGroupIndex: number }
  : never;

export type DocxComparisonInstruction = WithSemanticGroup<DocxComparisonInstructionPayload>;

export type DocxComparisonSemanticGroup = {
  readonly reports: readonly DocxComparisonReportInput[];
};

/** Exact operands and immutable instructions transferred to preflight once. */
export type ConsumedDocxComparisonProgram = {
  readonly comparison: ResolvedDocxStoryComparison;
  readonly sourceSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
  readonly semanticGroups: readonly DocxComparisonSemanticGroup[];
  readonly instructions: readonly DocxComparisonInstruction[];
};

const freezeRecursively = (value: unknown): void => {
  if (typeof value !== "object" || value === null || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freezeRecursively(child);
  Object.freeze(value);
};

const ownFormatting = (formatting: Readonly<TextFormatting>): Readonly<TextFormatting> => {
  const owned = structuredClone(formatting);
  freezeRecursively(owned);
  return owned;
};

const sameValue = (left: unknown, right: unknown): boolean => {
  if (left === right) return true;
  if (typeof left !== "object" || left === null || typeof right !== "object" || right === null) {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((value, index) => sameValue(value, right[index]));
  }
  const leftEntries = Object.entries(left).filter(([, value]) => value !== undefined);
  const rightEntries = Object.entries(right).filter(([, value]) => value !== undefined);
  if (leftEntries.length !== rightEntries.length) return false;
  return leftEntries.every(
    ([key, value]) => Object.hasOwn(right, key) && sameValue(value, Reflect.get(right, key)),
  );
};

const sameFormatting = (left: Readonly<TextFormatting>, right: Readonly<TextFormatting>): boolean =>
  Object.values(TEXT_FORMATTING_PROPERTY_DESCRIPTORS).every(({ field }) =>
    sameValue(left[field], right[field]),
  );

const ownRuns = (text: string, runs: readonly DocxAuthoredRun[]): readonly DocxAuthoredRun[] => {
  const owned: DocxAuthoredRun[] = [];
  let offset = 0;
  for (const run of runs) {
    if (
      run.startOffset !== offset ||
      run.endOffset < run.startOffset ||
      run.endOffset > text.length ||
      (run.endOffset === run.startOffset && text.length > 0)
    ) {
      return panic("A DOCX comparison run plan is not contiguous", { offset, run });
    }
    owned.push(
      Object.freeze({
        startOffset: run.startOffset,
        endOffset: run.endOffset,
        formatting: ownFormatting(run.formatting),
      }),
    );
    offset = run.endOffset;
  }
  if (offset !== text.length || (text.length === 0 && runs.length > 1)) {
    return panic("A DOCX comparison run plan does not reconstruct its text", {
      expected: text.length,
      actual: offset,
    });
  }
  return Object.freeze(owned);
};

type RunCursor = { readonly runs: readonly DocxAuthoredRun[]; index: number };

const runAt = (cursor: RunCursor, offset: number, textLength: number): DocxAuthoredRun => {
  while (
    cursor.index + 1 < cursor.runs.length &&
    (cursor.runs[cursor.index]?.endOffset ?? 0) <= offset
  ) {
    cursor.index++;
  }
  const run = cursor.runs[cursor.index];
  if (
    !run ||
    (offset < textLength && (offset < run.startOffset || offset >= run.endOffset)) ||
    (offset === textLength && run.endOffset !== textLength)
  ) {
    return panic("A DOCX comparison range lost its authored run", { offset, textLength });
  }
  return run;
};

const ownSegments = (
  sourceText: string,
  targetText: string,
  segments: readonly FolioContentTextSegment[],
): readonly FolioContentTextSegment[] => {
  const owned: FolioContentTextSegment[] = [];
  let baseOffset = 0;
  let revisedOffset = 0;
  for (const segment of segments) {
    if (
      segment.baseStart !== baseOffset ||
      segment.revisedStart !== revisedOffset ||
      segment.baseEnd < segment.baseStart ||
      segment.revisedEnd < segment.revisedStart
    ) {
      return panic("A DOCX comparison segment plan is not contiguous", { segment });
    }
    const baseSlice = sourceText.slice(segment.baseStart, segment.baseEnd);
    const revisedSlice = targetText.slice(segment.revisedStart, segment.revisedEnd);
    if (
      (segment.type !== "ins" && baseSlice !== segment.text) ||
      (segment.type !== "del" && revisedSlice !== segment.text) ||
      (segment.type === "equal" && baseSlice !== revisedSlice)
    ) {
      return panic("A DOCX comparison segment does not name its exact source and target text", {
        segment,
      });
    }
    owned.push(Object.freeze({ ...segment }));
    baseOffset = segment.baseEnd;
    revisedOffset = segment.revisedEnd;
  }
  if (baseOffset !== sourceText.length || revisedOffset !== targetText.length) {
    return panic("DOCX comparison segments do not reconstruct both texts", {
      baseOffset,
      revisedOffset,
      sourceLength: sourceText.length,
      targetLength: targetText.length,
    });
  }
  return Object.freeze(owned);
};

const knownFormattingProperties = new Set<string>(
  Object.values(TEXT_FORMATTING_PROPERTY_DESCRIPTORS).map(({ field }) => field),
);

const ownAuthoredChanges = (
  sourceText: string,
  targetText: string,
  ranges: readonly DocxAuthoredChangeRange[],
): readonly DocxAuthoredChangeRange[] => {
  const owned: DocxAuthoredChangeRange[] = [];
  let previousBaseEnd = 0;
  let previousRevisedEnd = 0;
  for (const range of ranges) {
    if (
      range.baseStart < previousBaseEnd ||
      range.revisedStart < previousRevisedEnd ||
      range.baseStart < 0 ||
      range.revisedStart < 0 ||
      range.baseEnd <= range.baseStart ||
      range.revisedEnd <= range.revisedStart ||
      range.baseEnd > sourceText.length ||
      range.revisedEnd > targetText.length ||
      range.baseEnd - range.baseStart !== range.revisedEnd - range.revisedStart ||
      sourceText.slice(range.baseStart, range.baseEnd) !==
        targetText.slice(range.revisedStart, range.revisedEnd)
    ) {
      return panic("A DOCX authored-formatting change range is invalid", { range });
    }
    const unique = new Set(range.properties);
    if (
      range.properties.length === 0 ||
      unique.size !== range.properties.length ||
      range.properties.some((property) => !knownFormattingProperties.has(property))
    ) {
      return panic("A DOCX authored-formatting change names invalid properties", { range });
    }
    owned.push(
      Object.freeze({
        baseStart: range.baseStart,
        baseEnd: range.baseEnd,
        revisedStart: range.revisedStart,
        revisedEnd: range.revisedEnd,
        properties: Object.freeze([...range.properties]),
      }),
    );
    previousBaseEnd = range.baseEnd;
    previousRevisedEnd = range.revisedEnd;
  }
  return Object.freeze(owned);
};

const changedRangeAt = (
  ranges: readonly DocxAuthoredChangeRange[],
  index: number,
  baseOffset: number,
  revisedOffset: number,
): DocxAuthoredChangeRange | null => {
  const range = ranges[index];
  return range &&
    baseOffset >= range.baseStart &&
    baseOffset < range.baseEnd &&
    revisedOffset >= range.revisedStart &&
    revisedOffset < range.revisedEnd
    ? range
    : null;
};

type AuthoredProjectionRun = {
  readonly text: string;
  readonly formatting: Readonly<TextFormatting>;
};

const appendProjectionRun = (
  projection: AuthoredProjectionRun[],
  text: string,
  formatting: Readonly<TextFormatting>,
): void => {
  if (text.length === 0) return;
  const previous = projection.at(-1);
  if (previous && sameFormatting(previous.formatting, formatting)) {
    projection[projection.length - 1] = Object.freeze({
      text: `${previous.text}${text}`,
      formatting: previous.formatting,
    });
    return;
  }
  projection.push(Object.freeze({ text, formatting }));
};

const projectionFromRuns = (
  text: string,
  runs: readonly DocxAuthoredRun[],
): readonly AuthoredProjectionRun[] => {
  const projection: AuthoredProjectionRun[] = [];
  for (const run of runs) {
    appendProjectionRun(projection, text.slice(run.startOffset, run.endOffset), run.formatting);
  }
  return projection;
};

const sameProjection = (
  left: readonly AuthoredProjectionRun[],
  right: readonly AuthoredProjectionRun[],
): boolean =>
  left.length === right.length &&
  left.every((run, index) => {
    const counterpart = right[index];
    return (
      counterpart !== undefined &&
      run.text === counterpart.text &&
      sameFormatting(run.formatting, counterpart.formatting)
    );
  });

const compileRange = (input: DocxComparisonRangePlanInput): DocxComparisonRangePlan => {
  const segments = ownSegments(input.sourceText, input.targetText, input.segments);
  const sourceRuns = ownRuns(input.sourceText, input.sourceRuns);
  const targetRuns = ownRuns(input.targetText, input.targetRuns);
  const authoredChanges = ownAuthoredChanges(
    input.sourceText,
    input.targetText,
    input.authoredChanges,
  );
  const sourceCursor: RunCursor = { runs: sourceRuns, index: 0 };
  const targetCursor: RunCursor = { runs: targetRuns, index: 0 };
  const fragments: DocxComparisonTextFragment[] = [];
  let changeIndex = 0;

  for (const segment of segments) {
    let baseOffset = segment.baseStart;
    let revisedOffset = segment.revisedStart;
    while (baseOffset < segment.baseEnd || revisedOffset < segment.revisedEnd) {
      while (
        authoredChanges[changeIndex] &&
        (authoredChanges[changeIndex]?.baseEnd ?? 0) <= baseOffset &&
        (authoredChanges[changeIndex]?.revisedEnd ?? 0) <= revisedOffset
      ) {
        changeIndex++;
      }
      if (segment.type === "del") {
        const sourceRun = runAt(sourceCursor, baseOffset, input.sourceText.length);
        const end = Math.min(segment.baseEnd, sourceRun.endOffset);
        fragments.push(
          Object.freeze({
            type: "del",
            text: input.sourceText.slice(baseOffset, end),
            baseStart: baseOffset,
            baseEnd: end,
            revisedStart: revisedOffset,
            revisedEnd: revisedOffset,
            sourceFormatting: sourceRun.formatting,
          }),
        );
        baseOffset = end;
        continue;
      }
      if (segment.type === "ins") {
        const targetRun = runAt(targetCursor, revisedOffset, input.targetText.length);
        const end = Math.min(segment.revisedEnd, targetRun.endOffset);
        fragments.push(
          Object.freeze({
            type: "ins",
            text: input.targetText.slice(revisedOffset, end),
            baseStart: baseOffset,
            baseEnd: baseOffset,
            revisedStart: revisedOffset,
            revisedEnd: end,
            targetFormatting: targetRun.formatting,
          }),
        );
        revisedOffset = end;
        continue;
      }

      const sourceRun = runAt(sourceCursor, baseOffset, input.sourceText.length);
      const targetRun = runAt(targetCursor, revisedOffset, input.targetText.length);
      const activeChange = changedRangeAt(authoredChanges, changeIndex, baseOffset, revisedOffset);
      const nextChange = authoredChanges[changeIndex];
      const changeBaseBoundary = activeChange
        ? activeChange.baseEnd
        : Math.min(nextChange?.baseStart ?? segment.baseEnd, segment.baseEnd);
      const changeRevisedBoundary = activeChange
        ? activeChange.revisedEnd
        : Math.min(nextChange?.revisedStart ?? segment.revisedEnd, segment.revisedEnd);
      const length = Math.min(
        segment.baseEnd - baseOffset,
        segment.revisedEnd - revisedOffset,
        sourceRun.endOffset - baseOffset,
        targetRun.endOffset - revisedOffset,
        changeBaseBoundary - baseOffset,
        changeRevisedBoundary - revisedOffset,
      );
      if (length <= 0) {
        return panic("A DOCX authored-formatting range does not align with an equal text segment", {
          segment,
          change: nextChange,
        });
      }
      const changedProperties = Object.freeze([...(activeChange?.properties ?? [])]);
      const actualChangedProperties = Object.values(TEXT_FORMATTING_PROPERTY_DESCRIPTORS)
        .filter(({ field }) => !sameValue(sourceRun.formatting[field], targetRun.formatting[field]))
        .map(({ field }) => field);
      if (
        actualChangedProperties.length !== changedProperties.length ||
        actualChangedProperties.some((property) => !changedProperties.includes(property))
      ) {
        return panic("Canonical authored-formatting changes do not reconstruct the target run", {
          actualChangedProperties,
          changedProperties,
        });
      }
      fragments.push(
        Object.freeze({
          type: "equal",
          text: input.sourceText.slice(baseOffset, baseOffset + length),
          baseStart: baseOffset,
          baseEnd: baseOffset + length,
          revisedStart: revisedOffset,
          revisedEnd: revisedOffset + length,
          sourceFormatting: sourceRun.formatting,
          targetFormatting: targetRun.formatting,
          changedProperties,
        }),
      );
      baseOffset += length;
      revisedOffset += length;
    }
  }

  const rejected: AuthoredProjectionRun[] = [];
  const accepted: AuthoredProjectionRun[] = [];
  for (const fragment of fragments) {
    if (fragment.type !== "ins") {
      appendProjectionRun(rejected, fragment.text, fragment.sourceFormatting);
    }
    if (fragment.type !== "del") {
      appendProjectionRun(accepted, fragment.text, fragment.targetFormatting);
    }
  }
  if (
    !sameProjection(rejected, projectionFromRuns(input.sourceText, sourceRuns)) ||
    !sameProjection(accepted, projectionFromRuns(input.targetText, targetRuns))
  ) {
    return panic("A DOCX comparison plan does not reconstruct accepted and rejected authored runs");
  }

  return Object.freeze({
    sourceText: input.sourceText,
    targetText: input.targetText,
    fragments: Object.freeze(fragments),
  });
};

const ownParagraphProperties = (
  properties: Readonly<FolioAIBlockParagraphProperties>,
): Readonly<FolioAIBlockParagraphProperties> => {
  const owned = structuredClone(properties);
  freezeRecursively(owned);
  return owned;
};

const ownTableLocation = (
  table: Readonly<FolioAIBlockTableLocation>,
): Readonly<FolioAIBlockTableLocation> => {
  const owned = structuredClone(table);
  freezeRecursively(owned);
  return owned;
};

const ownSourceOperand = (
  source: DocxComparisonSourceOperand,
  snapshot: ResolvedDocxStorySnapshot,
): DocxComparisonSourceOperand => {
  if (resolvedDocxSourceOperandSnapshot(source) !== snapshot) {
    return panic("A DOCX comparison program cannot mix source story snapshots");
  }
  resolvedDocxSourceOperandBlock(source, snapshot);
  return source;
};

const ownParagraphInsertionBoundary = (
  boundary: DocxComparisonParagraphInsertionBoundary,
  snapshot: ResolvedDocxStorySnapshot,
): DocxComparisonParagraphInsertionBoundary => {
  switch (boundary.type) {
    case "afterParagraph":
    case "beforeParagraph":
      return Object.freeze({
        type: boundary.type,
        paragraph: ownSourceOperand(boundary.paragraph, snapshot),
      });
    default: {
      const unreachable: never = boundary.type;
      return panic("A DOCX paragraph boundary has an invalid type", { type: unreachable });
    }
  }
};

const ownParagraphRemovalBoundary = (
  boundary: DocxComparisonParagraphRemovalBoundary,
  snapshot: ResolvedDocxStorySnapshot,
): DocxComparisonParagraphRemovalBoundary => {
  switch (boundary.type) {
    case "successorParagraph":
      return Object.freeze({
        type: boundary.type,
        successor: ownSourceOperand(boundary.successor, snapshot),
      });
    case "successorTable":
      return Object.freeze({
        type: boundary.type,
        firstBlock: ownSourceOperand(boundary.firstBlock, snapshot),
      });
    default: {
      const unreachable: never = boundary;
      return panic("A DOCX paragraph-removal boundary has an invalid type", {
        boundary: unreachable,
      });
    }
  }
};

const canonicalParagraphInsertionBoundary = (
  boundary: FolioContentParagraphInsertionBoundary,
  snapshot: ResolvedDocxStorySnapshot,
): DocxComparisonParagraphInsertionBoundary => {
  switch (boundary.type) {
    case "afterParagraph":
    case "beforeParagraph":
      return Object.freeze({
        type: boundary.type,
        paragraph: resolvedDocxSourceOperand(snapshot, boundary.paragraph),
      });
    case "unanchoredContainer":
      return panic("An unanchored insertion reached the DOCX semantic compiler");
    default: {
      const unreachable: never = boundary;
      return panic("Unhandled canonical paragraph insertion boundary", { boundary: unreachable });
    }
  }
};

const ownSourceOperandGroup = (
  sources: DocxComparisonSourceOperandGroup,
  snapshot: ResolvedDocxStorySnapshot,
): DocxComparisonSourceOperandGroup => {
  if (sources.length === 0 || sources.length > MAX_DOCX_COMPARISON_INSTRUCTIONS) {
    return panic("A trailing paragraph run has an invalid member count", {
      maximum: MAX_DOCX_COMPARISON_INSTRUCTIONS,
      actual: sources.length,
    });
  }
  const owned = sources.map((source) => ownSourceOperand(source, snapshot));
  const first = owned.at(0) ?? panic("A trailing paragraph run lost its first member");
  const identities = new Set(
    owned.map((source) => resolvedDocxSourceOperandBlock(source, snapshot).identity.id),
  );
  if (identities.size !== owned.length) {
    return panic("A trailing paragraph run repeats a source identity");
  }
  return Object.freeze([first, ...owned.slice(1)]);
};

const ownParagraphTargetBlock = (
  block: FolioContentBlock,
  targetSnapshot: ResolvedDocxStorySnapshot,
): DocxComparisonParagraphTarget =>
  Object.freeze({
    text: block.text,
    runs: ownRuns(block.text, resolvedDocxAuthoredRunsForBlock(targetSnapshot, block)),
    properties: ownParagraphProperties(docxParagraphPropertiesFromBlock(block)),
    ...(block.table !== undefined && {
      table: ownTableLocation(docxTableLocationFromContent(block.table)),
    }),
  });

const ownParagraphTarget = (
  target: ResolvedDocxTargetBlockOperand,
  comparison: ResolvedDocxStoryComparison,
  targetSnapshot: ResolvedDocxStorySnapshot,
): DocxComparisonParagraphTarget =>
  ownParagraphTargetBlock(resolvedDocxTargetBlockOperandBlock(target, comparison), targetSnapshot);

const authoredChangesForRelation = (
  relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
  baseStart: number,
  revisedStart: number,
): readonly DocxAuthoredChangeRange[] =>
  Object.freeze(
    relation.formatting?.ranges
      .filter(({ formatting }) => formatting.authored.length > 0)
      .map(
        ({
          baseStart: rangeBaseStart,
          baseEnd,
          revisedStart: rangeRevisedStart,
          revisedEnd,
          formatting,
        }) =>
          Object.freeze({
            baseStart: rangeBaseStart - baseStart,
            baseEnd: baseEnd - baseStart,
            revisedStart: rangeRevisedStart - revisedStart,
            revisedEnd: revisedEnd - revisedStart,
            properties: Object.freeze(formatting.authored.map(({ key }) => key)),
          }),
      ) ?? [],
  );

const rangeInputForRelation = (
  relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): DocxComparisonRangePlanInput => {
  const baseStart = relation.base.startOffset;
  const revisedStart = relation.revised.startOffset;
  const baseBlock = relation.base.block;
  const targetBlock = relation.revised.block;
  return {
    sourceText: baseBlock.text.slice(baseStart, relation.base.endOffset),
    targetText: targetBlock.text.slice(revisedStart, relation.revised.endOffset),
    segments: relation.segments.map((segment) =>
      Object.freeze({
        ...segment,
        baseStart: segment.baseStart - baseStart,
        baseEnd: segment.baseEnd - baseStart,
        revisedStart: segment.revisedStart - revisedStart,
        revisedEnd: segment.revisedEnd - revisedStart,
      }),
    ),
    sourceRuns: resolvedDocxAuthoredRunsForRange(
      sourceSnapshot,
      baseBlock,
      relation.base.startOffset,
      relation.base.endOffset,
    ),
    targetRuns: resolvedDocxAuthoredRunsForRange(
      targetSnapshot,
      targetBlock,
      relation.revised.startOffset,
      relation.revised.endOffset,
    ),
    authoredChanges: authoredChangesForRelation(relation, baseStart, revisedStart),
  };
};

type CompiledOperandRange = {
  readonly source: DocxComparisonSourceOperand;
  readonly sourceStartOffset: number;
  readonly range: DocxComparisonRangePlan;
};

const compilePairRelation = (
  relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): CompiledOperandRange => {
  const source = resolvedDocxSourceOperand(sourceSnapshot, relation.base.block);
  const range = compileRange(rangeInputForRelation(relation, sourceSnapshot, targetSnapshot));
  assertRangeMatchesSource(source, sourceSnapshot, relation.base.startOffset, range);
  return Object.freeze({ source, sourceStartOffset: relation.base.startOffset, range });
};

const compilePairRangeOperand = (
  operand: ResolvedDocxPairRangeOperand,
  comparison: ResolvedDocxStoryComparison,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): CompiledOperandRange =>
  compilePairRelation(
    resolvedDocxPairRangeOperandRelation(operand, comparison),
    sourceSnapshot,
    targetSnapshot,
  );

const compileFormattingRangeOperand = (
  operand: ResolvedDocxFormattingRangeOperand,
  comparison: ResolvedDocxStoryComparison,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): CompiledOperandRange => {
  const { relation, range: formattingRange } = resolvedDocxFormattingRangeOperandPayload(
    operand,
    comparison,
  );
  const sourceText = relation.base.block.text.slice(
    formattingRange.baseStart,
    formattingRange.baseEnd,
  );
  const targetText = relation.revised.block.text.slice(
    formattingRange.revisedStart,
    formattingRange.revisedEnd,
  );
  const source = resolvedDocxSourceOperand(sourceSnapshot, relation.base.block);
  const range = compileRange({
    sourceText,
    targetText,
    segments: Object.freeze([
      Object.freeze({
        type: "equal",
        text: sourceText,
        baseStart: 0,
        baseEnd: sourceText.length,
        revisedStart: 0,
        revisedEnd: targetText.length,
      }),
    ]),
    sourceRuns: resolvedDocxAuthoredRunsForRange(
      sourceSnapshot,
      relation.base.block,
      formattingRange.baseStart,
      formattingRange.baseEnd,
    ),
    targetRuns: resolvedDocxAuthoredRunsForRange(
      targetSnapshot,
      relation.revised.block,
      formattingRange.revisedStart,
      formattingRange.revisedEnd,
    ),
    authoredChanges: Object.freeze([
      Object.freeze({
        baseStart: 0,
        baseEnd: sourceText.length,
        revisedStart: 0,
        revisedEnd: targetText.length,
        properties: Object.freeze(formattingRange.formatting.authored.map(({ key }) => key)),
      }),
    ]),
  });
  assertRangeMatchesSource(source, sourceSnapshot, formattingRange.baseStart, range);
  return Object.freeze({ source, sourceStartOffset: formattingRange.baseStart, range });
};

const wholeReplacementSegments = (
  sourceText: string,
  targetText: string,
): readonly FolioContentTextSegment[] =>
  Object.freeze([
    ...(sourceText.length > 0
      ? [
          Object.freeze({
            type: "del" as const,
            text: sourceText,
            baseStart: 0,
            baseEnd: sourceText.length,
            revisedStart: 0,
            revisedEnd: 0,
          }),
        ]
      : []),
    ...(targetText.length > 0
      ? [
          Object.freeze({
            type: "ins" as const,
            text: targetText,
            baseStart: sourceText.length,
            baseEnd: sourceText.length,
            revisedStart: 0,
            revisedEnd: targetText.length,
          }),
        ]
      : []),
  ]);

const compileWholeBlockReplacement = (
  {
    baseBlock,
    targetBlock,
  }: {
    readonly baseBlock: FolioContentBlock;
    readonly targetBlock: FolioContentBlock;
  },
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): CompiledOperandRange => {
  const source = resolvedDocxSourceOperand(sourceSnapshot, baseBlock);
  const range = compileRange({
    sourceText: baseBlock.text,
    targetText: targetBlock.text,
    segments: wholeReplacementSegments(baseBlock.text, targetBlock.text),
    sourceRuns: resolvedDocxAuthoredRunsForBlock(sourceSnapshot, baseBlock),
    targetRuns: resolvedDocxAuthoredRunsForBlock(targetSnapshot, targetBlock),
    authoredChanges: Object.freeze([]),
  });
  assertRangeMatchesSource(source, sourceSnapshot, 0, range);
  return Object.freeze({ source, sourceStartOffset: 0, range });
};

const compileReplacementRangeOperand = (
  operand: ResolvedDocxPairRangeOperand | ResolvedDocxWholeBlockReplacementOperand,
  comparison: ResolvedDocxStoryComparison,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): CompiledOperandRange => {
  const payload = resolvedDocxReplacementRangeOperandPayload(operand, comparison);
  switch (payload.type) {
    case "pair":
      return compilePairRelation(payload.relation, sourceSnapshot, targetSnapshot);
    case "whole-block":
      return compileWholeBlockReplacement(payload, sourceSnapshot, targetSnapshot);
    default: {
      const unreachable: never = payload;
      return panic("Unhandled DOCX replacement-range operand", { payload: unreachable });
    }
  }
};

const assertRangeMatchesSource = (
  source: DocxComparisonSourceOperand,
  snapshot: ResolvedDocxStorySnapshot,
  sourceStartOffset: number,
  range: DocxComparisonRangePlan,
): void => {
  const block = resolvedDocxSourceOperandBlock(source, snapshot);
  if (
    !Number.isSafeInteger(sourceStartOffset) ||
    sourceStartOffset < 0 ||
    block.text.slice(sourceStartOffset, sourceStartOffset + range.sourceText.length) !==
      range.sourceText
  ) {
    return panic("A DOCX comparison range does not name its exact source block", {
      blockId: block.identity.id,
      sourceStartOffset,
    });
  }
};

const assertRunsEqual = (
  text: string,
  left: readonly DocxAuthoredRun[],
  right: readonly DocxAuthoredRun[],
): void => {
  if (!sameProjection(projectionFromRuns(text, left), projectionFromRuns(text, right))) {
    return panic("A DOCX comparison instruction has conflicting target run projections");
  }
};

const concatenateRuns = (
  parts: readonly { readonly text: string; readonly runs: readonly DocxAuthoredRun[] }[],
): readonly DocxAuthoredRun[] => {
  const runs: DocxAuthoredRun[] = [];
  let offset = 0;
  for (const part of parts) {
    for (const run of part.runs) {
      runs.push(
        Object.freeze({
          startOffset: offset + run.startOffset,
          endOffset: offset + run.endOffset,
          formatting: run.formatting,
        }),
      );
    }
    offset += part.text.length;
  }
  return Object.freeze(runs);
};

const compileInstruction = (
  input: DocxComparisonInstructionInput,
  comparison: ResolvedDocxStoryComparison,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): DocxComparisonInstructionPayload => {
  switch (input.type) {
    case "replaceText": {
      const owned = compileReplacementRangeOperand(
        input.range,
        comparison,
        sourceSnapshot,
        targetSnapshot,
      );
      const { source, sourceStartOffset, range } = owned;
      if (
        range.sourceText === range.targetText &&
        range.fragments.every(
          (fragment) => fragment.type === "equal" && fragment.changedProperties.length === 0,
        )
      ) {
        return panic("A DOCX comparison replacement instruction contains no semantic change");
      }
      return Object.freeze({
        type: "replaceText",
        source,
        sourceStartOffset,
        range,
      });
    }
    case "formatText": {
      const { source, sourceStartOffset, range } = compileFormattingRangeOperand(
        input.range,
        comparison,
        sourceSnapshot,
        targetSnapshot,
      );
      if (
        range.sourceText !== range.targetText ||
        range.fragments.some((fragment) => fragment.type !== "equal") ||
        !range.fragments.some(
          (fragment) => fragment.type === "equal" && fragment.changedProperties.length > 0,
        )
      ) {
        return panic(
          "A DOCX comparison formatting instruction must contain only formatting changes",
        );
      }
      return Object.freeze({
        type: "formatText",
        source,
        sourceStartOffset,
        range,
      });
    }
    case "insertParagraph":
      return Object.freeze({
        type: "insertParagraph",
        boundary: ownParagraphInsertionBoundary(input.boundary, sourceSnapshot),
        target: ownParagraphTarget(input.target, comparison, targetSnapshot),
      });
    case "deleteParagraph":
      return Object.freeze({
        type: "deleteParagraph",
        source: ownSourceOperand(input.source, sourceSnapshot),
      });
    case "moveParagraph":
      return Object.freeze({
        type: "moveParagraph",
        source: ownSourceOperand(input.source, sourceSnapshot),
        removalBoundary: ownParagraphRemovalBoundary(input.removalBoundary, sourceSnapshot),
        boundary: ownParagraphInsertionBoundary(input.boundary, sourceSnapshot),
        target: ownParagraphTarget(input.target, comparison, targetSnapshot),
      });
    case "moveTerminalParagraph":
      return Object.freeze({
        type: "moveTerminalParagraph",
        predecessor: ownSourceOperand(input.predecessor, sourceSnapshot),
        source: ownSourceOperand(input.source, sourceSnapshot),
        carrierTargetProperties: ownParagraphProperties(
          docxParagraphPropertiesFromBlock(
            resolvedDocxTargetBlockOperandBlock(input.carrierTarget, comparison),
          ),
        ),
        boundary: ownParagraphInsertionBoundary(input.boundary, sourceSnapshot),
        target: ownParagraphTarget(input.target, comparison, targetSnapshot),
      });
    case "splitParagraph": {
      const firstCompiled = compilePairRangeOperand(
        input.first,
        comparison,
        sourceSnapshot,
        targetSnapshot,
      );
      const secondCompiled = compilePairRangeOperand(
        input.second,
        comparison,
        sourceSnapshot,
        targetSnapshot,
      );
      const firstRelation = resolvedDocxPairRangeOperandRelation(input.first, comparison);
      const secondRelation = resolvedDocxPairRangeOperandRelation(input.second, comparison);
      const separatorRelation = resolvedDocxSeparatorOperandRelation(input.separator, comparison);
      const source = firstCompiled.source;
      const sourceBlock = resolvedDocxSourceOperandBlock(source, sourceSnapshot);
      const first = firstCompiled.range;
      const second = secondCompiled.range;
      const separatorText = separatorRelation.base.block.text.slice(
        separatorRelation.base.startOffset,
        separatorRelation.base.endOffset,
      );
      const separatorRuns = ownRuns(
        separatorText,
        resolvedDocxAuthoredRunsForRange(
          sourceSnapshot,
          separatorRelation.base.block,
          separatorRelation.base.startOffset,
          separatorRelation.base.endOffset,
        ),
      );
      const firstTarget = ownParagraphTargetBlock(firstRelation.revised.block, targetSnapshot);
      const secondTarget = ownParagraphTargetBlock(secondRelation.revised.block, targetSnapshot);
      if (
        secondCompiled.source !== source ||
        separatorRelation.base.block !== sourceBlock ||
        firstCompiled.sourceStartOffset !== 0 ||
        firstRelation.base.endOffset !== separatorRelation.base.startOffset ||
        separatorRelation.base.endOffset !== secondRelation.base.startOffset ||
        secondRelation.base.endOffset !== sourceBlock.text.length ||
        firstRelation.revised.startOffset !== 0 ||
        firstRelation.revised.endOffset !== firstRelation.revised.block.text.length ||
        secondRelation.revised.startOffset !== 0 ||
        secondRelation.revised.endOffset !== secondRelation.revised.block.text.length ||
        firstRelation.revised.block === secondRelation.revised.block ||
        sourceBlock.text !== `${first.sourceText}${separatorText}${second.sourceText}` ||
        first.targetText !== firstTarget.text ||
        second.targetText !== secondTarget.text
      ) {
        return panic("A DOCX comparison split does not partition its source and targets exactly");
      }
      assertRunsEqual(
        first.targetText,
        firstTarget.runs,
        resolvedDocxAuthoredRunsForRange(
          targetSnapshot,
          firstRelation.revised.block,
          firstRelation.revised.startOffset,
          firstRelation.revised.endOffset,
        ),
      );
      assertRunsEqual(
        second.targetText,
        secondTarget.runs,
        resolvedDocxAuthoredRunsForRange(
          targetSnapshot,
          secondRelation.revised.block,
          secondRelation.revised.startOffset,
          secondRelation.revised.endOffset,
        ),
      );
      return Object.freeze({
        type: "splitParagraph",
        source,
        offset: first.sourceText.length,
        first,
        second,
        separatorText,
        separatorRuns,
        firstTarget,
        secondTarget,
      });
    }
    case "mergeParagraphs": {
      const firstCompiled = compilePairRangeOperand(
        input.first,
        comparison,
        sourceSnapshot,
        targetSnapshot,
      );
      const secondCompiled = compilePairRangeOperand(
        input.second,
        comparison,
        sourceSnapshot,
        targetSnapshot,
      );
      const firstRelation = resolvedDocxPairRangeOperandRelation(input.first, comparison);
      const secondRelation = resolvedDocxPairRangeOperandRelation(input.second, comparison);
      const separatorRelation = resolvedDocxSeparatorOperandRelation(input.separator, comparison);
      const firstSource = firstCompiled.source;
      const secondSource = secondCompiled.source;
      const firstSourceBlock = resolvedDocxSourceOperandBlock(firstSource, sourceSnapshot);
      const secondSourceBlock = resolvedDocxSourceOperandBlock(secondSource, sourceSnapshot);
      const first = firstCompiled.range;
      const second = secondCompiled.range;
      const separatorText = separatorRelation.revised.block.text.slice(
        separatorRelation.revised.startOffset,
        separatorRelation.revised.endOffset,
      );
      const separatorRuns = ownRuns(
        separatorText,
        resolvedDocxAuthoredRunsForRange(
          targetSnapshot,
          separatorRelation.revised.block,
          separatorRelation.revised.startOffset,
          separatorRelation.revised.endOffset,
        ),
      );
      const target = ownParagraphTargetBlock(firstRelation.revised.block, targetSnapshot);
      const targetText = `${first.targetText}${separatorText}${second.targetText}`;
      if (
        firstRelation.revised.block !== secondRelation.revised.block ||
        separatorRelation.revised.block !== firstRelation.revised.block ||
        firstRelation.revised.startOffset !== 0 ||
        firstRelation.revised.endOffset !== separatorRelation.revised.startOffset ||
        separatorRelation.revised.endOffset !== secondRelation.revised.startOffset ||
        secondRelation.revised.endOffset !== target.text.length ||
        firstSourceBlock.text !== first.sourceText ||
        secondSourceBlock.text !== second.sourceText ||
        target.text !== targetText
      ) {
        return panic("A DOCX comparison merge does not reconstruct its sources and target exactly");
      }
      assertRunsEqual(
        targetText,
        target.runs,
        concatenateRuns([
          {
            text: first.targetText,
            runs: resolvedDocxAuthoredRunsForRange(
              targetSnapshot,
              firstRelation.revised.block,
              firstRelation.revised.startOffset,
              firstRelation.revised.endOffset,
            ),
          },
          { text: separatorText, runs: separatorRuns },
          {
            text: second.targetText,
            runs: resolvedDocxAuthoredRunsForRange(
              targetSnapshot,
              secondRelation.revised.block,
              secondRelation.revised.startOffset,
              secondRelation.revised.endOffset,
            ),
          },
        ]),
      );
      return Object.freeze({
        type: "mergeParagraphs",
        firstSource,
        secondSource,
        first,
        second,
        separatorText,
        separatorRuns,
        target,
      });
    }
    case "deleteTrailingParagraphs": {
      const chainStart = ownSourceOperand(input.chainStart, sourceSnapshot);
      const deleted = ownSourceOperandGroup(input.deleted, sourceSnapshot);
      const chainStartId = resolvedDocxSourceOperandBlock(chainStart, sourceSnapshot).identity.id;
      if (
        deleted.some(
          (source) =>
            resolvedDocxSourceOperandBlock(source, sourceSnapshot).identity.id === chainStartId,
        )
      ) {
        return panic("A trailing paragraph run includes its surviving chain start");
      }
      return Object.freeze({ type: "deleteTrailingParagraphs", chainStart, deleted });
    }
    case "setParagraphProperties": {
      const targetProperties = docxParagraphPropertiesFromBlock(
        resolvedDocxTargetBlockOperandBlock(input.target, comparison),
      );
      if (
        docxParagraphPropertiesEqual(
          docxParagraphPropertiesFromBlock(
            resolvedDocxSourceOperandBlock(input.source, sourceSnapshot),
          ),
          targetProperties,
        )
      ) {
        return panic("A DOCX comparison paragraph-property instruction contains no change");
      }
      return Object.freeze({
        type: "setParagraphProperties",
        source: ownSourceOperand(input.source, sourceSnapshot),
        targetProperties: ownParagraphProperties(targetProperties),
      });
    }
    case "tableStructure":
      resolvedDocxTableStructureOperandPayload(input.operation, comparison);
      return input;
    case "tableFormat":
      resolvedDocxTableFormatOperandPayload(input.operation, comparison);
      return input;
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX comparison instruction", { instruction: unreachable });
    }
  }
};

const REPORTS_PER_EVENT = 16;

const reportSequence = (eventSequence: number, withinEvent: number): number => {
  if (withinEvent < 0 || withinEvent >= REPORTS_PER_EVENT) {
    return panic("A DOCX semantic operation exceeded its per-event report budget", {
      eventSequence,
      withinEvent,
    });
  }
  return eventSequence * REPORTS_PER_EVENT + withinEvent;
};

const reportLocation = (
  story: ReturnType<typeof resolvedDocxStoryComparisonPayload>["baseStory"],
  block: FolioContentBlock,
) => (block.table ? { story, cell: docxTableLocationFromContent(block.table) } : { story });

const representedFormattingChange = (
  formatting: FolioContentFormattingChange["ranges"][number]["formatting"],
) => {
  const authoredKeys = new Set(formatting.authored.map(({ key }) => key));
  return Object.freeze({
    authored: formatting.authored,
    effective: Object.freeze(formatting.effective.filter(({ key }) => authoredKeys.has(key))),
  });
};

const formatReport = ({
  relation,
  story,
  sequence,
}: {
  readonly relation: FolioContentWholePairRelation | FolioContentRangePairRelation;
  readonly story: ReturnType<typeof resolvedDocxStoryComparisonPayload>["baseStory"];
  readonly sequence: number;
}): DocxComparisonReportInput | null => {
  const ranges =
    relation.formatting?.ranges
      .filter(({ formatting }) => formatting.authored.length > 0)
      .map(({ baseStart, baseEnd, formatting }) =>
        Object.freeze({
          startOffset: baseStart,
          endOffset: baseEnd,
          formatting: representedFormattingChange(formatting),
        }),
      ) ?? [];
  if (ranges.length === 0) return null;
  return Object.freeze({
    sequence,
    change: Object.freeze({
      kind: "format" as const,
      location: reportLocation(story, relation.base.block),
      baseBlockId: relation.base.block.identity.id,
      targetBlockId: relation.revised.block.identity.id,
      text: relation.base.block.text,
      ranges: Object.freeze(ranges),
    }),
  });
};

const paragraphFormatReport = ({
  relation,
  story,
  sequence,
}: {
  readonly relation: FolioContentWholePairRelation | FolioContentRangePairRelation;
  readonly story: ReturnType<typeof resolvedDocxStoryComparisonPayload>["baseStory"];
  readonly sequence: number;
}): DocxComparisonReportInput | null => {
  const authored = relation.formatting?.paragraph.authored ?? [];
  const properties = docxParagraphChangedProperties({
    changes: authored,
    target: relation.revised.block,
  });
  if (Object.keys(properties).length === 0) return null;
  return Object.freeze({
    sequence,
    change: Object.freeze({
      kind: "paragraph-format" as const,
      location: reportLocation(story, relation.base.block),
      baseBlockId: relation.base.block.identity.id,
      targetBlockId: relation.revised.block.identity.id,
      properties: ownParagraphProperties(properties),
    }),
  });
};

type CompiledSemanticOperationInput = {
  readonly reports: readonly DocxComparisonReportInput[];
  readonly instructions: NonEmptyReadonlyArray<DocxComparisonInstructionInput>;
};

const compiledSemanticOperation = (
  input: CompiledSemanticOperationInput,
): CompiledSemanticOperationInput =>
  Object.freeze({
    reports: Object.freeze([...input.reports]),
    instructions: Object.freeze(input.instructions),
  });

const ownSemanticReport = ({
  sequence,
  change,
}: DocxComparisonReportInput): DocxComparisonReportInput => {
  if (!Number.isSafeInteger(sequence) || sequence < 0) {
    return panic("A DOCX comparison report has an invalid canonical sequence", { sequence });
  }
  const ownedChange = structuredClone(change);
  freezeRecursively(ownedChange);
  return Object.freeze({ sequence, change: ownedChange });
};

const tableRowCellTexts = (blocks: readonly FolioContentBlock[]): string[] => {
  const byCell: (string | undefined)[] = [];
  for (const block of blocks) {
    const cellIndex = block.table?.cellIndex ?? 0;
    const existing = byCell[cellIndex];
    byCell[cellIndex] = existing === undefined ? block.text : `${existing}\n${block.text}`;
  }
  return Array.from(byCell, (text) => text ?? "");
};

const tableColumnCellTexts = (blocks: readonly FolioContentBlock[]): string[] => {
  const byCell = new Map<string, string>();
  for (const block of blocks) {
    const table = block.table;
    if (!table) continue;
    const key = `${String(table.rowIndex)}:${String(table.cellIndex)}`;
    const existing = byCell.get(key);
    byCell.set(key, existing === undefined ? block.text : `${existing}\n${block.text}`);
  }
  return [...byCell.values()];
};

const tableCellTexts = (blocks: readonly FolioContentBlock[]): string[][] => {
  const rows = groupFolioContentTableRows(blocks).map((row) => tableRowCellTexts(row));
  let width = 0;
  for (const row of rows) width = Math.max(width, row.length);
  for (const row of rows) {
    while (row.length < width) row.push("");
  }
  return rows;
};

const tableStructuralChangeReport = ({
  change,
  story,
}: {
  readonly change: FolioContentStructuralChange;
  readonly story: ReturnType<typeof resolvedDocxStoryComparisonPayload>["baseStory"];
}): CompareChange => {
  const firstBlock = change.blocks[0];
  const table = firstBlock.table;
  if (!table) return panic("A canonical table change has no table location");
  const location = { story, cell: docxTableLocationFromContent(table) };
  switch (change.type) {
    case "table-delete":
      return {
        kind: "table-delete",
        location,
        tableIndex: change.tableIndex,
        rows: tableCellTexts(change.blocks),
        baseBlockIds: change.blocks.map(({ identity }) => identity.id),
      };
    case "table-insert":
      return {
        kind: "table-insert",
        location,
        tableIndex: change.tableIndex,
        rows: tableCellTexts(change.blocks),
        targetBlockIds: change.blocks.map(({ identity }) => identity.id),
      };
    case "table-row-delete":
      return {
        kind: "table-row-delete",
        location,
        tableIndex: change.tableIndex,
        rowIndex: change.rowIndex,
        cells: tableRowCellTexts(change.blocks),
        baseBlockIds: change.blocks.map(({ identity }) => identity.id),
      };
    case "table-row-insert":
      return {
        kind: "table-row-insert",
        location,
        tableIndex: change.tableIndex,
        rowIndex: change.rowIndex,
        cells: tableRowCellTexts(change.blocks),
        targetBlockIds: change.blocks.map(({ identity }) => identity.id),
      };
    case "table-column-delete":
      return {
        kind: "table-column-delete",
        location,
        tableIndex: change.tableIndex,
        columnIndex: change.columnIndex,
        cells: tableColumnCellTexts(change.blocks),
        baseBlockIds: change.blocks.map(({ identity }) => identity.id),
      };
    case "table-column-insert":
      return {
        kind: "table-column-insert",
        location,
        tableIndex: change.tableIndex,
        columnIndex: change.columnIndex,
        cells: tableColumnCellTexts(change.blocks),
        targetBlockIds: change.blocks.map(({ identity }) => identity.id),
      };
    default: {
      const unreachable: never = change;
      return panic("Unhandled table structural report", { change: unreachable });
    }
  }
};

const compileTableStructureOperation = (
  operation: ResolvedDocxTableStructureOperand,
  comparison: ResolvedDocxStoryComparison,
  story: ReturnType<typeof resolvedDocxStoryComparisonPayload>["baseStory"],
): CompiledSemanticOperationInput => {
  const reports: DocxComparisonReportInput[] = [];
  for (const owner of resolvedDocxTableStructureReportOwners(operation, comparison)) {
    switch (owner.type) {
      case "structural":
        reports.push(
          ownSemanticReport({
            sequence: reportSequence(owner.sequence, 0),
            change: tableStructuralChangeReport({ change: owner.change, story }),
          }),
        );
        break;
      case "replacement": {
        const base = owner.replacement.baseBlocks[0];
        const target = owner.replacement.revisedBlocks[0];
        if (!base.table || !target.table) {
          return panic("A canonical table replacement has no table location");
        }
        reports.push(
          ownSemanticReport({
            sequence: reportSequence(owner.sequence, 0),
            change: {
              kind: "table-delete",
              location: { story, cell: docxTableLocationFromContent(base.table) },
              tableIndex: owner.replacement.baseTableIndex,
              rows: tableCellTexts(owner.replacement.baseBlocks),
              baseBlockIds: owner.replacement.baseBlocks.map(({ identity }) => identity.id),
            },
          }),
          ownSemanticReport({
            sequence: reportSequence(owner.sequence, 1),
            change: {
              kind: "table-insert",
              location: { story, cell: docxTableLocationFromContent(target.table) },
              tableIndex: owner.replacement.revisedTableIndex,
              rows: tableCellTexts(owner.replacement.revisedBlocks),
              targetBlockIds: owner.replacement.revisedBlocks.map(({ identity }) => identity.id),
            },
          }),
        );
        break;
      }
      case "deletedCarrier":
        reports.push(
          ownSemanticReport({
            sequence: reportSequence(owner.sequence, 0),
            change: {
              kind: "delete",
              location: reportLocation(story, owner.event.block),
              baseBlockId: owner.event.block.identity.id,
              before: owner.event.block.text,
            },
          }),
        );
        break;
      default: {
        const unreachable: never = owner;
        return panic("Unhandled table report owner", { owner: unreachable });
      }
    }
  }
  return compiledSemanticOperation({
    reports: Object.freeze(reports),
    instructions: Object.freeze([{ type: "tableStructure", operation }]),
  });
};

const compileTableFormatOperation = (
  operation: ResolvedDocxTableFormatOperand,
  comparison: ResolvedDocxStoryComparison,
  story: ReturnType<typeof resolvedDocxStoryComparisonPayload>["baseStory"],
): CompiledSemanticOperationInput => {
  const payload = resolvedDocxTableFormatOperandPayload(operation, comparison);
  const reports: DocxComparisonReportInput[] = [];
  if (payload.status === "ready") {
    const withinEvent = new Map<number, number>();
    for (const change of payload.changes) {
      const offset = withinEvent.get(change.sequence) ?? 8;
      withinEvent.set(change.sequence, offset + 1);
      reports.push(
        ownSemanticReport({
          sequence: reportSequence(change.sequence, offset),
          change: {
            kind: "table-format",
            location: { story },
            scope: change.scope,
            base: change.base,
            target: change.target,
          },
        }),
      );
    }
  }
  return compiledSemanticOperation({
    reports: Object.freeze(reports),
    instructions: Object.freeze([{ type: "tableFormat", operation }]),
  });
};

const compileSemanticOperationInput = (
  input: DocxComparisonOperationInput,
  comparison: ResolvedDocxStoryComparison,
  sourceSnapshot: ResolvedDocxStorySnapshot,
  targetSnapshot: ResolvedDocxStorySnapshot,
): CompiledSemanticOperationInput => {
  const { baseStory: story } = resolvedDocxStoryComparisonPayload(comparison);
  switch (input.type) {
    case "pairedBlock": {
      const { event, sequence } = resolvedDocxPairedEventOperandPayload(input.event, comparison);
      const { relation } = event;
      const reports: DocxComparisonReportInput[] = [];
      const instructions: DocxComparisonInstructionInput[] = [];
      const changesText = relation.segments.some(({ type }) => type !== "equal");
      if (changesText) {
        reports.push({
          sequence: reportSequence(sequence, 0),
          change: {
            kind: "replace",
            location: reportLocation(story, relation.base.block),
            baseBlockId: relation.base.block.identity.id,
            targetBlockId: relation.revised.block.identity.id,
            before: relation.base.block.text,
            after: relation.revised.block.text,
          },
        });
        instructions.push({
          type: "replaceText",
          range: resolvedDocxPairRangeOperand(comparison, relation),
        });
      }
      const inlineReport = formatReport({
        relation,
        story,
        sequence: reportSequence(sequence, 1),
      });
      if (inlineReport) {
        reports.push(inlineReport);
        if (!changesText) {
          for (const range of relation.formatting?.ranges ?? []) {
            if (range.formatting.authored.length === 0) continue;
            instructions.push({
              type: "formatText",
              range: resolvedDocxFormattingRangeOperand(comparison, relation, range),
            });
          }
        }
      }
      const paragraphReport = paragraphFormatReport({
        relation,
        story,
        sequence: reportSequence(sequence, 2),
      });
      if (paragraphReport) {
        reports.push(paragraphReport);
        instructions.push({
          type: "setParagraphProperties",
          source: resolvedDocxSourceOperand(sourceSnapshot, relation.base.block),
          target: resolvedDocxTargetBlockOperand(comparison, relation.revised.block),
        });
      }
      const firstInstruction = instructions.at(0);
      if (!firstInstruction) {
        return panic("A paired DOCX semantic operation contains no representable change");
      }
      return compiledSemanticOperation({
        reports: Object.freeze(reports.map(ownSemanticReport)),
        instructions: Object.freeze([firstInstruction, ...instructions.slice(1)]),
      });
    }
    case "insertParagraph": {
      const { event, sequence } = resolvedDocxInsertedEventOperandPayload(input.event, comparison);
      const boundary = canonicalParagraphInsertionBoundary(event.boundary, sourceSnapshot);
      return compiledSemanticOperation({
        reports: Object.freeze([
          ownSemanticReport({
            sequence: reportSequence(sequence, 0),
            change: {
              kind: "insert",
              location: reportLocation(story, event.block),
              targetBlockId: event.block.identity.id,
              after: event.block.text,
            },
          }),
        ]),
        instructions: Object.freeze([
          {
            type: "insertParagraph",
            boundary,
            target: resolvedDocxTargetBlockOperand(comparison, event.block),
          },
        ]),
      });
    }
    case "deleteParagraph": {
      const { event, sequence } = resolvedDocxDeletedEventOperandPayload(input.event, comparison);
      return compiledSemanticOperation({
        reports: Object.freeze([
          ownSemanticReport({
            sequence: reportSequence(sequence, 0),
            change: {
              kind: "delete",
              location: reportLocation(story, event.block),
              baseBlockId: event.block.identity.id,
              before: event.block.text,
            },
          }),
        ]),
        instructions: Object.freeze([
          {
            type: "deleteParagraph",
            source: resolvedDocxSourceOperand(sourceSnapshot, event.block),
          },
        ]),
      });
    }
    case "moveParagraph": {
      const { event, sequence } = resolvedDocxMoveEventOperandPayload(input.event, comparison);
      const { relation, sourceRemovalBoundary } = event.move;
      const boundary = canonicalParagraphInsertionBoundary(
        event.move.destinationBoundary,
        sourceSnapshot,
      );
      const reports: DocxComparisonReportInput[] = [
        ownSemanticReport({
          sequence: reportSequence(sequence, 0),
          change: {
            kind: "move",
            location: reportLocation(story, relation.revised.block),
            baseBlockId: relation.base.block.identity.id,
            targetBlockId: relation.revised.block.identity.id,
            text: relation.revised.block.text,
          },
        }),
      ];
      const inlineReport = formatReport({
        relation,
        story,
        sequence: reportSequence(sequence, 1),
      });
      if (inlineReport) reports.push(ownSemanticReport(inlineReport));
      const paragraphReport = paragraphFormatReport({
        relation,
        story,
        sequence: reportSequence(sequence, 2),
      });
      if (paragraphReport) reports.push(ownSemanticReport(paragraphReport));
      switch (sourceRemovalBoundary.type) {
        case "successorParagraph":
          return compiledSemanticOperation({
            reports: Object.freeze(reports),
            instructions: Object.freeze([
              {
                type: "moveParagraph",
                source: resolvedDocxSourceOperand(sourceSnapshot, relation.base.block),
                removalBoundary: Object.freeze({
                  type: "successorParagraph",
                  successor: resolvedDocxSourceOperand(
                    sourceSnapshot,
                    sourceRemovalBoundary.successor,
                  ),
                }),
                boundary,
                target: resolvedDocxTargetBlockOperand(comparison, relation.revised.block),
              },
            ]),
          });
        case "successorTable":
          return compiledSemanticOperation({
            reports: Object.freeze(reports),
            instructions: Object.freeze([
              {
                type: "moveParagraph",
                source: resolvedDocxSourceOperand(sourceSnapshot, relation.base.block),
                removalBoundary: Object.freeze({
                  type: "successorTable",
                  firstBlock: resolvedDocxSourceOperand(
                    sourceSnapshot,
                    sourceRemovalBoundary.firstBlock,
                  ),
                }),
                boundary,
                target: resolvedDocxTargetBlockOperand(comparison, relation.revised.block),
              },
            ]),
          });
        case "terminalPredecessor":
          return compiledSemanticOperation({
            reports: Object.freeze(reports),
            instructions: Object.freeze([
              {
                type: "moveTerminalParagraph",
                predecessor: resolvedDocxSourceOperand(
                  sourceSnapshot,
                  sourceRemovalBoundary.predecessor,
                ),
                source: resolvedDocxSourceOperand(sourceSnapshot, relation.base.block),
                carrierTarget: resolvedDocxTargetBlockOperand(
                  comparison,
                  sourceRemovalBoundary.targetCarrier,
                ),
                boundary,
                target: resolvedDocxTargetBlockOperand(comparison, relation.revised.block),
              },
            ]),
          });
        case "unanchoredContainer":
          return panic("An unanchored move reached the DOCX semantic compiler");
        default: {
          const unreachable: never = sourceRemovalBoundary;
          return panic("Unhandled DOCX move removal boundary", { boundary: unreachable });
        }
      }
    }
    case "splitParagraph": {
      const { event, sequence } = resolvedDocxSplitEventOperandPayload(input.event, comparison);
      const [first, second] = event.relations;
      const reports: DocxComparisonReportInput[] = [
        ownSemanticReport({
          sequence: reportSequence(sequence, 0),
          change: {
            kind: "split",
            location: reportLocation(story, first.base.block),
            baseBlockId: first.base.block.identity.id,
            targetBlockIds: [first.revised.block.identity.id, second.revised.block.identity.id],
            text: first.base.block.text,
          },
        }),
      ];
      const companions = [
        formatReport({ relation: first, story, sequence: reportSequence(sequence, 1) }),
        formatReport({ relation: second, story, sequence: reportSequence(sequence, 2) }),
        paragraphFormatReport({ relation: first, story, sequence: reportSequence(sequence, 3) }),
        paragraphFormatReport({ relation: second, story, sequence: reportSequence(sequence, 4) }),
      ];
      for (const companion of companions) {
        if (companion) reports.push(ownSemanticReport(companion));
      }
      return compiledSemanticOperation({
        reports: Object.freeze(reports),
        instructions: Object.freeze([
          {
            type: "splitParagraph",
            first: resolvedDocxPairRangeOperand(comparison, first),
            second: resolvedDocxPairRangeOperand(comparison, second),
            separator: resolvedDocxSeparatorOperand(comparison, event.separator),
          },
        ]),
      });
    }
    case "mergeParagraphs": {
      const { event, sequence } = resolvedDocxMergeEventOperandPayload(input.event, comparison);
      const [first, second] = event.relations;
      const reports: DocxComparisonReportInput[] = [
        ownSemanticReport({
          sequence: reportSequence(sequence, 0),
          change: {
            kind: "merge",
            location: reportLocation(story, first.base.block),
            baseBlockIds: [first.base.block.identity.id, second.base.block.identity.id],
            targetBlockId: first.revised.block.identity.id,
            text: first.revised.block.text,
          },
        }),
      ];
      const companions = [
        formatReport({ relation: first, story, sequence: reportSequence(sequence, 1) }),
        formatReport({ relation: second, story, sequence: reportSequence(sequence, 2) }),
        paragraphFormatReport({ relation: first, story, sequence: reportSequence(sequence, 3) }),
      ];
      for (const companion of companions) {
        if (companion) reports.push(ownSemanticReport(companion));
      }
      return compiledSemanticOperation({
        reports: Object.freeze(reports),
        instructions: Object.freeze([
          {
            type: "mergeParagraphs",
            first: resolvedDocxPairRangeOperand(comparison, first),
            second: resolvedDocxPairRangeOperand(comparison, second),
            separator: resolvedDocxSeparatorOperand(comparison, event.separator),
          },
        ]),
      });
    }
    case "deleteTrailingParagraphs": {
      const operation = resolvedDocxTrailingDeletionOperandPayload(input.operation, comparison);
      const deleted = operation.events;
      const firstDeleted = deleted.at(0);
      if (!firstDeleted) return panic("A trailing deletion operation has no deleted event");
      const reports = deleted.map(({ event, sequence }) =>
        ownSemanticReport({
          sequence: reportSequence(sequence, 0),
          change: {
            kind: "delete",
            location: reportLocation(story, event.block),
            baseBlockId: event.block.identity.id,
            before: event.block.text,
          },
        }),
      );
      const instructions: DocxComparisonInstructionInput[] = [
        {
          type: "deleteTrailingParagraphs",
          chainStart: operation.chainStart,
          deleted: [
            resolvedDocxSourceOperand(sourceSnapshot, firstDeleted.event.block),
            ...deleted
              .slice(1)
              .map(({ event }) => resolvedDocxSourceOperand(sourceSnapshot, event.block)),
          ],
        },
      ];
      if (operation.targetCarrier) {
        const carrier = deleted.at(-1)?.event.block;
        if (!carrier) return panic("A trailing deletion operation lost its carrier");
        const target = operation.targetCarrier;
        if (
          !docxParagraphPropertiesEqual(
            docxParagraphPropertiesFromBlock(carrier),
            docxParagraphPropertiesFromBlock(target),
          )
        ) {
          instructions.push({
            type: "setParagraphProperties",
            source: resolvedDocxSourceOperand(sourceSnapshot, carrier),
            target: resolvedDocxTargetBlockOperand(comparison, target),
          });
        }
      }
      return compiledSemanticOperation({
        reports: Object.freeze(reports),
        instructions: Object.freeze([
          instructions[0] ?? panic("A trailing deletion operation lost its instruction"),
          ...instructions.slice(1),
        ]),
      });
    }
    case "replaceTerminalParagraph": {
      const { deleted, inserted } = resolvedDocxTerminalReplacementOperandPayload(
        input.operation,
        comparison,
      );
      const baseBlock = deleted.event.block;
      const targetBlock = inserted.event.block;
      const reports = Object.freeze([
        ownSemanticReport({
          sequence: reportSequence(deleted.sequence, 0),
          change: {
            kind: "delete",
            location: reportLocation(story, baseBlock),
            baseBlockId: baseBlock.identity.id,
            before: baseBlock.text,
          },
        }),
        ownSemanticReport({
          sequence: reportSequence(inserted.sequence, 0),
          change: {
            kind: "insert",
            location: reportLocation(story, targetBlock),
            targetBlockId: targetBlock.identity.id,
            after: targetBlock.text,
          },
        }),
      ]);
      const instructions: DocxComparisonInstructionInput[] = [];
      const baseRuns = resolvedDocxAuthoredRunsForBlock(sourceSnapshot, baseBlock);
      const targetRuns = resolvedDocxAuthoredRunsForBlock(targetSnapshot, targetBlock);
      if (
        baseBlock.text !== targetBlock.text ||
        !sameProjection(
          projectionFromRuns(baseBlock.text, baseRuns),
          projectionFromRuns(targetBlock.text, targetRuns),
        )
      ) {
        instructions.push({
          type: "replaceText",
          range: resolvedDocxWholeBlockReplacementOperand(comparison, baseBlock, targetBlock),
        });
      }
      if (
        !docxParagraphPropertiesEqual(
          docxParagraphPropertiesFromBlock(baseBlock),
          docxParagraphPropertiesFromBlock(targetBlock),
        )
      ) {
        instructions.push({
          type: "setParagraphProperties",
          source: resolvedDocxSourceOperand(sourceSnapshot, baseBlock),
          target: resolvedDocxTargetBlockOperand(comparison, targetBlock),
        });
      }
      const firstInstruction = instructions.at(0);
      if (!firstInstruction) {
        return panic("A terminal replacement operation contains no semantic change");
      }
      return compiledSemanticOperation({
        reports,
        instructions: Object.freeze([firstInstruction, ...instructions.slice(1)]),
      });
    }
    case "tableStructure":
      return compileTableStructureOperation(input.operation, comparison, story);
    case "tableFormat":
      return compileTableFormatOperation(input.operation, comparison, story);
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX semantic operation", { operation: unreachable });
    }
  }
};

type CompiledDocxComparisonProgram = {
  readonly instructions: readonly DocxComparisonInstruction[];
  readonly semanticGroups: readonly DocxComparisonSemanticGroup[];
};

type CompileDocxComparisonProgramOptions = {
  readonly comparison: ResolvedDocxStoryComparison;
  readonly inputs: readonly DocxComparisonOperationInput[];
  readonly sourceSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
};

/** Every source paragraph mark is one edge and has exactly one instruction owner. */
const assertUniqueParagraphMarkOwnership = (
  instructions: readonly DocxComparisonInstruction[],
  sourceSnapshot: ResolvedDocxStorySnapshot,
): void => {
  const ownerByBlock = new Map<
    FolioContentBlock,
    { readonly instructionIndex: number; readonly instructionType: string }
  >();
  for (const [instructionIndex, instruction] of instructions.entries()) {
    let owners: readonly DocxComparisonSourceOperand[];
    switch (instruction.type) {
      case "deleteParagraph":
      case "moveParagraph":
      case "splitParagraph":
        owners = [instruction.source];
        break;
      case "moveTerminalParagraph":
        owners = [instruction.predecessor];
        break;
      case "mergeParagraphs":
        owners = [instruction.firstSource];
        break;
      case "deleteTrailingParagraphs":
        owners = [instruction.chainStart, ...instruction.deleted.slice(0, -1)];
        break;
      case "replaceText":
      case "formatText":
      case "insertParagraph":
      case "setParagraphProperties":
      case "tableStructure":
      case "tableFormat":
        owners = [];
        break;
      default: {
        const unreachable: never = instruction;
        return panic("Unhandled instruction while proving paragraph-mark ownership", {
          instruction: unreachable,
        });
      }
    }
    for (const owner of owners) {
      const block = resolvedDocxSourceOperandBlock(owner, sourceSnapshot);
      const existing = ownerByBlock.get(block);
      if (existing !== undefined) {
        return panic("A source paragraph edge has more than one comparison instruction owner", {
          blockId: block.identity.id,
          firstInstructionIndex: existing.instructionIndex,
          firstInstructionType: existing.instructionType,
          secondInstructionIndex: instructionIndex,
          secondInstructionType: instruction.type,
        });
      }
      ownerByBlock.set(block, { instructionIndex, instructionType: instruction.type });
    }
  }
};

const compileDocxComparisonProgram = ({
  comparison,
  inputs,
  sourceSnapshot,
  targetSnapshot,
}: CompileDocxComparisonProgramOptions): CompiledDocxComparisonProgram => {
  const semanticGroups: DocxComparisonSemanticGroup[] = [];
  const instructions: DocxComparisonInstruction[] = [];
  const reportSequences = new Set<number>();
  for (const input of inputs) {
    const compiled = compileSemanticOperationInput(
      input,
      comparison,
      sourceSnapshot,
      targetSnapshot,
    );
    for (const { sequence } of compiled.reports) {
      if (reportSequences.has(sequence)) {
        return panic("A DOCX comparison report has an invalid canonical sequence", {
          sequence,
        });
      }
      reportSequences.add(sequence);
    }
    const semanticGroupIndex = semanticGroups.length;
    semanticGroups.push(Object.freeze({ reports: compiled.reports }));
    for (const instruction of compiled.instructions) {
      instructions.push(
        Object.freeze({
          ...compileInstruction(instruction, comparison, sourceSnapshot, targetSnapshot),
          semanticGroupIndex,
        }),
      );
    }
    if (instructions.length > MAX_DOCX_COMPARISON_INSTRUCTIONS) {
      return panic("A DOCX comparison program exceeds its instruction limit", {
        limit: MAX_DOCX_COMPARISON_INSTRUCTIONS,
        actual: instructions.length,
      });
    }
  }
  assertUniqueParagraphMarkOwnership(instructions, sourceSnapshot);
  return Object.freeze({
    instructions: Object.freeze(instructions),
    semanticGroups: Object.freeze(semanticGroups),
  });
};

/**
 * An owned one-shot transport plan. Construction validates the complete
 * instruction graph; consumption transfers its immutable semantics to the
 * dedicated comparison executor exactly once.
 */
export class DocxComparisonProgram {
  readonly #comparison: ResolvedDocxStoryComparison;
  readonly #instructions: readonly DocxComparisonInstruction[];
  readonly #semanticGroups: readonly DocxComparisonSemanticGroup[];
  readonly #sourceSnapshot: ResolvedDocxStorySnapshot;
  readonly #targetSnapshot: ResolvedDocxStorySnapshot;
  #state: "ready" | "consumed" = "ready";

  private constructor(
    comparison: ResolvedDocxStoryComparison,
    inputs: readonly DocxComparisonOperationInput[],
  ) {
    const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
    const compiled = compileDocxComparisonProgram({
      comparison,
      inputs,
      sourceSnapshot: baseSnapshot,
      targetSnapshot,
    });
    this.#comparison = comparison;
    this.#sourceSnapshot = baseSnapshot;
    this.#targetSnapshot = targetSnapshot;
    this.#instructions = compiled.instructions;
    this.#semanticGroups = compiled.semanticGroups;
  }

  static create(
    comparison: ResolvedDocxStoryComparison,
    inputs: readonly DocxComparisonOperationInput[],
  ): DocxComparisonProgram {
    if (inputs.length > MAX_DOCX_COMPARISON_INSTRUCTIONS) {
      return panic("A DOCX comparison program exceeds its instruction limit", {
        limit: MAX_DOCX_COMPARISON_INSTRUCTIONS,
        actual: inputs.length,
      });
    }
    return new DocxComparisonProgram(comparison, inputs);
  }

  get size(): number {
    return this.#instructions.length;
  }

  consume(): ConsumedDocxComparisonProgram {
    if (this.#state !== "ready") {
      return panic("A DOCX comparison program was consumed more than once");
    }
    this.#state = "consumed";
    return Object.freeze({
      comparison: this.#comparison,
      sourceSnapshot: this.#sourceSnapshot,
      targetSnapshot: this.#targetSnapshot,
      semanticGroups: this.#semanticGroups,
      instructions: this.#instructions,
    });
  }
}
