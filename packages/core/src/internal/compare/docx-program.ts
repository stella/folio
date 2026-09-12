import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import type {
  FolioAIBlockParagraphProperties,
  FolioAIBlockTableLocation,
} from "../../ai-edits/types";
import type { TableGeometryPairing } from "./table-geometry-program";
import type { TextFormatting } from "../../types/document";
import type {
  FolioContentRangePairRelation,
  FolioContentTextSegment,
  FolioContentWholePairRelation,
} from "../../compare/content";
import type { FolioContentBlock, FolioContentTableLocation } from "../../compare/content-types";
import {
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
  resolvedDocxTableNodes,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  resolvedDocxFormattingRangeOperandPayload,
  resolvedDocxPairRangeOperandRelation,
  resolvedDocxPairedBaseTableIndexes,
  resolvedDocxReplacementRangeOperandPayload,
  resolvedDocxSeparatorOperandRelation,
  resolvedDocxStoryComparisonPayload,
  resolvedDocxTableGeometryPairings,
  resolvedDocxTargetBlockOperandBlock,
  resolvedDocxTargetColumnOperandChange,
  resolvedDocxTargetRowOperandChange,
  resolvedDocxTargetTableOperandOwner,
  type ResolvedDocxFormattingRangeOperand,
  type ResolvedDocxPairRangeOperand,
  type ResolvedDocxSeparatorOperand,
  type ResolvedDocxStoryComparison,
  type ResolvedDocxTargetBlockOperand,
  type ResolvedDocxTargetColumnOperand,
  type ResolvedDocxTargetRowOperand,
  type ResolvedDocxTargetTableOperand,
  type ResolvedDocxWholeBlockReplacementOperand,
} from "./resolved-docx-story-comparison";

const MAX_DOCX_COMPARISON_INSTRUCTIONS = 10_000;
const MAX_DOCX_COMPARISON_GEOMETRY_PAIRINGS = 10_000;

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

type DocxComparisonStructuralInsertionOperand = {
  readonly source: DocxComparisonSourceOperand;
  readonly position: "after" | "before";
};

export type DocxComparisonStructuralInsertionAnchor = {
  readonly blockId: string;
  readonly position: "after" | "before";
};

export type DocxComparisonParagraphInsertionBoundary = {
  readonly type: "afterParagraph" | "beforeParagraph";
  readonly paragraph: DocxComparisonSourceOperand;
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
      readonly successor: DocxComparisonSourceOperand;
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
      readonly type: "insertTable";
      readonly anchor: DocxComparisonStructuralInsertionOperand;
      readonly target: ResolvedDocxTargetTableOperand;
    }
  | {
      readonly type: "deleteTable";
      readonly source: DocxComparisonSourceOperand;
    }
  | {
      readonly type: "replaceTable";
      readonly source: DocxComparisonSourceOperand;
      readonly target: ResolvedDocxTargetTableOperand;
    }
  | {
      readonly type: "insertTableRow";
      readonly anchor: DocxComparisonStructuralInsertionOperand;
      readonly target: ResolvedDocxTargetRowOperand;
    }
  | {
      readonly type: "deleteTableRow";
      readonly source: DocxComparisonSourceOperand;
    }
  | {
      readonly type: "insertTableColumn";
      readonly anchor: DocxComparisonStructuralInsertionOperand;
      readonly target: ResolvedDocxTargetColumnOperand;
    }
  | {
      readonly type: "deleteTableColumn";
      readonly source: DocxComparisonSourceOperand;
    }
  | {
      readonly type: "matchTableGeometry";
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

export type DocxComparisonInstruction =
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
      readonly successor: DocxComparisonSourceOperand;
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
      readonly type: "insertTable";
      readonly anchor: DocxComparisonStructuralInsertionAnchor;
      readonly targetTableIndex: number;
    }
  | {
      readonly type: "deleteTable";
      readonly source: DocxComparisonSourceOperand;
      readonly baseTableIndex: number;
    }
  | {
      readonly type: "replaceTable";
      readonly source: DocxComparisonSourceOperand;
      readonly baseTableIndex: number;
      readonly targetTableIndex: number;
    }
  | {
      readonly type: "insertTableRow";
      readonly anchor: DocxComparisonStructuralInsertionAnchor;
      readonly targetTableIndex: number;
      readonly targetRowIndex: number;
    }
  | {
      readonly type: "deleteTableRow";
      readonly source: DocxComparisonSourceOperand;
      readonly baseTableIndex: number;
      readonly baseRowIndex: number;
    }
  | {
      readonly type: "insertTableColumn";
      readonly anchor: DocxComparisonStructuralInsertionAnchor;
      readonly targetTableIndex: number;
      readonly targetColumnIndex: number;
      readonly cellTexts: readonly string[];
    }
  | {
      readonly type: "deleteTableColumn";
      readonly source: DocxComparisonSourceOperand;
      readonly baseTableIndex: number;
      readonly baseColumnIndex: number;
    }
  | {
      readonly type: "matchTableGeometry";
      readonly pairings: readonly TableGeometryPairing[];
    };

/** Exact operands and immutable instructions transferred to preflight once. */
export type ConsumedDocxComparisonProgram = {
  readonly comparison: ResolvedDocxStoryComparison;
  readonly sourceSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
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

const ownStructuralInsertionAnchor = (
  anchor: DocxComparisonStructuralInsertionOperand,
  snapshot: ResolvedDocxStorySnapshot,
): DocxComparisonStructuralInsertionAnchor => {
  switch (anchor.position) {
    case "after":
    case "before":
      return Object.freeze({
        blockId: resolvedDocxSourceOperandBlock(ownSourceOperand(anchor.source, snapshot), snapshot)
          .identity.id,
        position: anchor.position,
      });
    default: {
      const unreachable: never = anchor.position;
      return panic("A DOCX comparison insertion anchor has an invalid position", {
        position: unreachable,
      });
    }
  }
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

const ownIndex = (name: string, value: number): number => {
  if (!Number.isSafeInteger(value) || value < 0) {
    return panic("A DOCX comparison structural index is invalid", { name, value });
  }
  return value;
};

const ownTableGeometryPairings = (
  pairings: readonly TableGeometryPairing[],
): readonly TableGeometryPairing[] => {
  if (pairings.length > MAX_DOCX_COMPARISON_GEOMETRY_PAIRINGS) {
    return panic("A DOCX comparison geometry instruction exceeds its pairing limit", {
      limit: MAX_DOCX_COMPARISON_GEOMETRY_PAIRINGS,
      actual: pairings.length,
    });
  }
  const owned = structuredClone(pairings);
  freezeRecursively(owned);
  return owned;
};

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

const tableLocationOf = (block: FolioContentBlock, side: "base" | "target") =>
  block.table ?? panic(`A DOCX ${side} table operand has no canonical table location`);

const targetTableNode = (targetSnapshot: ResolvedDocxStorySnapshot, tableIndex: number) =>
  resolvedDocxTableNodes(targetSnapshot).get(tableIndex) ??
  panic("A DOCX target table operand has no exact target table node", { tableIndex });

const targetTableFromOperand = (
  operand: ResolvedDocxTargetTableOperand,
  expected: "inserted" | "replacement",
  comparison: ResolvedDocxStoryComparison,
  targetSnapshot: ResolvedDocxStorySnapshot,
): number => {
  const owner = resolvedDocxTargetTableOperandOwner(operand, comparison);
  const inserted = "type" in owner;
  if ((expected === "inserted") !== inserted) {
    return panic("A DOCX target table operand belongs to another instruction kind", {
      expected,
    });
  }
  const tableIndex = ownIndex(
    "targetTableIndex",
    inserted ? owner.tableIndex : owner.revisedTableIndex,
  );
  const blocks = inserted ? owner.blocks : owner.revisedBlocks;
  if (!blocks.some((block) => block.table?.tableIndex === tableIndex)) {
    return panic("A DOCX target table operand does not own its canonical table coordinates", {
      tableIndex,
    });
  }
  targetTableNode(targetSnapshot, tableIndex);
  return tableIndex;
};

const targetRowFromOperand = (
  operand: ResolvedDocxTargetRowOperand,
  comparison: ResolvedDocxStoryComparison,
  targetSnapshot: ResolvedDocxStorySnapshot,
): { readonly tableIndex: number; readonly rowIndex: number } => {
  const change = resolvedDocxTargetRowOperandChange(operand, comparison);
  const tableIndex = ownIndex("targetTableIndex", change.tableIndex);
  const rowIndex = ownIndex("targetRowIndex", change.rowIndex);
  if (
    !change.blocks.some(
      (block) => block.table?.tableIndex === tableIndex && block.table.rowIndex === rowIndex,
    )
  ) {
    return panic("A DOCX target row operand does not own its canonical row coordinates", {
      tableIndex,
      rowIndex,
    });
  }
  const node = targetTableNode(targetSnapshot, tableIndex);
  if (rowIndex >= node.childCount) {
    return panic("A DOCX target row operand has no exact target row node", {
      tableIndex,
      rowIndex,
    });
  }
  node.child(rowIndex);
  return Object.freeze({ tableIndex, rowIndex });
};

const columnCellTexts = (blocks: readonly FolioContentBlock[]): readonly string[] => {
  const byCell = new Map<string, string>();
  for (const block of blocks) {
    const table = block.table;
    if (!table) return panic("A DOCX target column contains a block outside its table");
    const key = `${String(table.rowIndex)}:${String(table.cellIndex)}`;
    const existing = byCell.get(key);
    byCell.set(key, existing === undefined ? block.text : `${existing}\n${block.text}`);
  }
  return Object.freeze([...byCell.values()]);
};

const targetColumnFromOperand = (
  operand: ResolvedDocxTargetColumnOperand,
  comparison: ResolvedDocxStoryComparison,
  targetSnapshot: ResolvedDocxStorySnapshot,
): {
  readonly tableIndex: number;
  readonly columnIndex: number;
  readonly cellTexts: readonly string[];
} => {
  const change = resolvedDocxTargetColumnOperandChange(operand, comparison);
  const tableIndex = ownIndex("targetTableIndex", change.tableIndex);
  const columnIndex = ownIndex("targetColumnIndex", change.columnIndex);
  if (
    !change.blocks.every(
      (block) =>
        block.table?.tableIndex === tableIndex && block.table.gridColumnIndex === columnIndex,
    )
  ) {
    return panic("A DOCX target column operand does not own its canonical column coordinates", {
      tableIndex,
      columnIndex,
    });
  }
  targetTableNode(targetSnapshot, tableIndex);
  return Object.freeze({
    tableIndex,
    columnIndex,
    cellTexts: columnCellTexts(change.blocks),
  });
};

const sourceTableLocation = (
  source: DocxComparisonSourceOperand,
  sourceSnapshot: ResolvedDocxStorySnapshot,
): FolioContentTableLocation =>
  tableLocationOf(resolvedDocxSourceOperandBlock(source, sourceSnapshot), "base");

const assertStructuralAnchorPairsWithTargetTable = (
  anchor: DocxComparisonStructuralInsertionOperand,
  targetTableIndex: number,
  comparison: ResolvedDocxStoryComparison,
  sourceSnapshot: ResolvedDocxStorySnapshot,
): void => {
  const sourceLocation = sourceTableLocation(anchor.source, sourceSnapshot);
  if (
    !resolvedDocxPairedBaseTableIndexes(comparison, targetTableIndex).has(sourceLocation.tableIndex)
  ) {
    return panic("A DOCX structural insertion anchor belongs to another canonical table", {
      baseTableIndex: sourceLocation.tableIndex,
      targetTableIndex,
    });
  }
};

const tableGeometryPairings = (
  comparison: ResolvedDocxStoryComparison,
): readonly TableGeometryPairing[] =>
  ownTableGeometryPairings(resolvedDocxTableGeometryPairings(comparison));

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
): DocxComparisonInstruction => {
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
        successor: ownSourceOperand(input.successor, sourceSnapshot),
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
    case "insertTable": {
      const target = targetTableFromOperand(input.target, "inserted", comparison, targetSnapshot);
      return Object.freeze({
        type: "insertTable",
        anchor: ownStructuralInsertionAnchor(input.anchor, sourceSnapshot),
        targetTableIndex: target,
      });
    }
    case "deleteTable": {
      const source = ownSourceOperand(input.source, sourceSnapshot);
      return Object.freeze({
        type: "deleteTable",
        source,
        baseTableIndex: ownIndex(
          "baseTableIndex",
          sourceTableLocation(source, sourceSnapshot).tableIndex,
        ),
      });
    }
    case "replaceTable": {
      const source = ownSourceOperand(input.source, sourceSnapshot);
      const owner = resolvedDocxTargetTableOperandOwner(input.target, comparison);
      if (
        "type" in owner ||
        owner.baseBlocks.at(0) !== resolvedDocxSourceOperandBlock(source, sourceSnapshot)
      ) {
        return panic("A DOCX replacement table operand does not own its exact base table");
      }
      const target = targetTableFromOperand(
        input.target,
        "replacement",
        comparison,
        targetSnapshot,
      );
      return Object.freeze({
        type: "replaceTable",
        source,
        baseTableIndex: ownIndex(
          "baseTableIndex",
          sourceTableLocation(source, sourceSnapshot).tableIndex,
        ),
        targetTableIndex: target,
      });
    }
    case "insertTableRow": {
      const target = targetRowFromOperand(input.target, comparison, targetSnapshot);
      assertStructuralAnchorPairsWithTargetTable(
        input.anchor,
        target.tableIndex,
        comparison,
        sourceSnapshot,
      );
      return Object.freeze({
        type: "insertTableRow",
        anchor: ownStructuralInsertionAnchor(input.anchor, sourceSnapshot),
        targetTableIndex: target.tableIndex,
        targetRowIndex: target.rowIndex,
      });
    }
    case "deleteTableRow": {
      const source = ownSourceOperand(input.source, sourceSnapshot);
      const location = sourceTableLocation(source, sourceSnapshot);
      return Object.freeze({
        type: "deleteTableRow",
        source,
        baseTableIndex: ownIndex("baseTableIndex", location.tableIndex),
        baseRowIndex: ownIndex("baseRowIndex", location.rowIndex),
      });
    }
    case "insertTableColumn": {
      const target = targetColumnFromOperand(input.target, comparison, targetSnapshot);
      assertStructuralAnchorPairsWithTargetTable(
        input.anchor,
        target.tableIndex,
        comparison,
        sourceSnapshot,
      );
      return Object.freeze({
        type: "insertTableColumn",
        anchor: ownStructuralInsertionAnchor(input.anchor, sourceSnapshot),
        targetTableIndex: target.tableIndex,
        targetColumnIndex: target.columnIndex,
        cellTexts: target.cellTexts,
      });
    }
    case "deleteTableColumn": {
      const source = ownSourceOperand(input.source, sourceSnapshot);
      const location = sourceTableLocation(source, sourceSnapshot);
      return Object.freeze({
        type: "deleteTableColumn",
        source,
        baseTableIndex: ownIndex("baseTableIndex", location.tableIndex),
        baseColumnIndex: ownIndex("baseColumnIndex", location.gridColumnIndex),
      });
    }
    case "matchTableGeometry":
      return Object.freeze({
        type: "matchTableGeometry",
        pairings: tableGeometryPairings(comparison),
      });
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX comparison instruction", { instruction: unreachable });
    }
  }
};

/**
 * An owned one-shot transport plan. Construction validates the complete
 * instruction graph; consumption transfers its immutable semantics to the
 * dedicated comparison executor exactly once.
 */
export class DocxComparisonProgram {
  readonly #comparison: ResolvedDocxStoryComparison;
  readonly #instructions: readonly DocxComparisonInstruction[];
  readonly #sourceSnapshot: ResolvedDocxStorySnapshot;
  readonly #targetSnapshot: ResolvedDocxStorySnapshot;
  #state: "ready" | "consumed" = "ready";

  private constructor(
    comparison: ResolvedDocxStoryComparison,
    inputs: readonly DocxComparisonInstructionInput[],
  ) {
    if (inputs.length > MAX_DOCX_COMPARISON_INSTRUCTIONS) {
      panic("A DOCX comparison program exceeds its instruction limit", {
        limit: MAX_DOCX_COMPARISON_INSTRUCTIONS,
        actual: inputs.length,
      });
    }
    const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
    this.#comparison = comparison;
    this.#sourceSnapshot = baseSnapshot;
    this.#targetSnapshot = targetSnapshot;
    this.#instructions = Object.freeze(
      inputs.map((input) => compileInstruction(input, comparison, baseSnapshot, targetSnapshot)),
    );
  }

  static create(
    comparison: ResolvedDocxStoryComparison,
    inputs: readonly DocxComparisonInstructionInput[],
  ): DocxComparisonProgram {
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
      instructions: this.#instructions,
    });
  }
}
