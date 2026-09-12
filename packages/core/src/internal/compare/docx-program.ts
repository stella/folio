import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import type {
  FolioAIBlockParagraphProperties,
  FolioAIBlockTableLocation,
} from "../../ai-edits/types";
import type { TableGeometryPairing } from "./table-geometry-program";
import type { TextFormatting } from "../../types/document";
import type { FolioContentTextSegment } from "../../compare/content";
import {
  docxParagraphPropertiesEqual,
  docxParagraphPropertiesFromBlock,
} from "./docx-paragraph-transport";
import {
  resolvedDocxSourceOperandBlock,
  resolvedDocxSourceOperandSnapshot,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  resolvedDocxStoryComparisonPayload,
  type ResolvedDocxStoryComparison,
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

export type DocxComparisonRangePlanInput = {
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

export type DocxComparisonStructuralInsertionAnchor = {
  readonly blockId: string;
  readonly position: "after" | "before";
};

export type DocxComparisonParagraphInsertionBoundary = {
  readonly type: "afterParagraph" | "beforeParagraph";
  readonly paragraph: DocxComparisonSourceOperand;
};

export type DocxComparisonParagraphTargetInput = {
  readonly text: string;
  readonly runs: readonly DocxAuthoredRun[];
  readonly properties: Readonly<FolioAIBlockParagraphProperties>;
  readonly table?: Readonly<FolioAIBlockTableLocation>;
};

/**
 * Closed transport vocabulary for one DOCX story. These are not generic edit
 * requests: each branch owns its source expectation and complete accepted
 * payload, so execution never joins an operation to a second semantic record.
 */
export type DocxComparisonInstructionInput =
  | {
      readonly type: "replaceText";
      readonly source: DocxComparisonSourceOperand;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlanInput;
    }
  | {
      readonly type: "formatText";
      readonly source: DocxComparisonSourceOperand;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlanInput;
    }
  | {
      readonly type: "insertParagraph";
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: DocxComparisonParagraphTargetInput;
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
      readonly target: DocxComparisonParagraphTargetInput;
    }
  | {
      readonly type: "moveTerminalParagraph";
      readonly predecessor: DocxComparisonSourceOperand;
      readonly source: DocxComparisonSourceOperand;
      readonly carrierTargetProperties: Readonly<FolioAIBlockParagraphProperties>;
      readonly boundary: DocxComparisonParagraphInsertionBoundary;
      readonly target: DocxComparisonParagraphTargetInput;
    }
  | {
      readonly type: "splitParagraph";
      readonly source: DocxComparisonSourceOperand;
      readonly offset: number;
      readonly first: DocxComparisonRangePlanInput;
      readonly second: DocxComparisonRangePlanInput;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
      readonly firstTarget: DocxComparisonParagraphTargetInput;
      readonly secondTarget: DocxComparisonParagraphTargetInput;
    }
  | {
      readonly type: "mergeParagraphs";
      readonly firstSource: DocxComparisonSourceOperand;
      readonly secondSource: DocxComparisonSourceOperand;
      readonly first: DocxComparisonRangePlanInput;
      readonly second: DocxComparisonRangePlanInput;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
      readonly target: DocxComparisonParagraphTargetInput;
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
  | Extract<
      DocxComparisonInstructionInput,
      {
        readonly type:
          | "insertTable"
          | "deleteTable"
          | "replaceTable"
          | "insertTableRow"
          | "deleteTableRow"
          | "insertTableColumn"
          | "deleteTableColumn"
          | "matchTableGeometry";
      }
    >;

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
  table: Readonly<FolioAIBlockTableLocation> | undefined,
): Readonly<FolioAIBlockTableLocation> | undefined => {
  if (table === undefined) return undefined;
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
  anchor: DocxComparisonStructuralInsertionAnchor,
): DocxComparisonStructuralInsertionAnchor => {
  if (anchor.blockId.length === 0) {
    return panic("A DOCX comparison insertion anchor has no block identity");
  }
  switch (anchor.position) {
    case "after":
    case "before":
      return Object.freeze({ blockId: anchor.blockId, position: anchor.position });
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

const ownParagraphTarget = (
  target: DocxComparisonParagraphTargetInput,
): DocxComparisonParagraphTarget =>
  Object.freeze({
    text: target.text,
    runs: ownRuns(target.text, target.runs),
    properties: ownParagraphProperties(target.properties),
    ...(target.table !== undefined && { table: ownTableLocation(target.table) }),
  });

const ownIndex = (name: string, value: number): number => {
  if (!Number.isSafeInteger(value) || value < 0) {
    return panic("A DOCX comparison structural index is invalid", { name, value });
  }
  return value;
};

const ownCellTexts = (cellTexts: readonly string[]): readonly string[] => {
  const owned: string[] = [];
  for (const text of cellTexts) {
    if (typeof text !== "string") {
      return panic("A DOCX comparison table column contains non-text content");
    }
    owned.push(text);
  }
  return Object.freeze(owned);
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
  sourceSnapshot: ResolvedDocxStorySnapshot,
): DocxComparisonInstruction => {
  switch (input.type) {
    case "replaceText": {
      const source = ownSourceOperand(input.source, sourceSnapshot);
      const range = compileRange(input.range);
      assertRangeMatchesSource(source, sourceSnapshot, input.sourceStartOffset, range);
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
        sourceStartOffset: input.sourceStartOffset,
        range,
      });
    }
    case "formatText": {
      const source = ownSourceOperand(input.source, sourceSnapshot);
      const range = compileRange(input.range);
      assertRangeMatchesSource(source, sourceSnapshot, input.sourceStartOffset, range);
      if (
        range.sourceText !== range.targetText ||
        range.fragments.some((fragment) => fragment.type !== "equal") ||
        range.fragments.every(({ changedProperties }) => changedProperties.length === 0)
      ) {
        return panic(
          "A DOCX comparison formatting instruction must contain only formatting changes",
        );
      }
      return Object.freeze({
        type: "formatText",
        source,
        sourceStartOffset: input.sourceStartOffset,
        range,
      });
    }
    case "insertParagraph":
      return Object.freeze({
        type: "insertParagraph",
        boundary: ownParagraphInsertionBoundary(input.boundary, sourceSnapshot),
        target: ownParagraphTarget(input.target),
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
        target: ownParagraphTarget(input.target),
      });
    case "moveTerminalParagraph":
      return Object.freeze({
        type: "moveTerminalParagraph",
        predecessor: ownSourceOperand(input.predecessor, sourceSnapshot),
        source: ownSourceOperand(input.source, sourceSnapshot),
        carrierTargetProperties: ownParagraphProperties(input.carrierTargetProperties),
        boundary: ownParagraphInsertionBoundary(input.boundary, sourceSnapshot),
        target: ownParagraphTarget(input.target),
      });
    case "splitParagraph": {
      const source = ownSourceOperand(input.source, sourceSnapshot);
      const sourceBlock = resolvedDocxSourceOperandBlock(source, sourceSnapshot);
      const first = compileRange(input.first);
      const second = compileRange(input.second);
      const separatorRuns = ownRuns(input.separatorText, input.separatorRuns);
      const firstTarget = ownParagraphTarget(input.firstTarget);
      const secondTarget = ownParagraphTarget(input.secondTarget);
      if (
        input.offset !== first.sourceText.length ||
        sourceBlock.text !== `${first.sourceText}${input.separatorText}${second.sourceText}` ||
        first.targetText !== firstTarget.text ||
        second.targetText !== secondTarget.text
      ) {
        return panic("A DOCX comparison split does not partition its source and targets exactly");
      }
      assertRunsEqual(
        first.targetText,
        firstTarget.runs,
        ownRuns(first.targetText, input.first.targetRuns),
      );
      assertRunsEqual(
        second.targetText,
        secondTarget.runs,
        ownRuns(second.targetText, input.second.targetRuns),
      );
      return Object.freeze({
        type: "splitParagraph",
        source,
        offset: input.offset,
        first,
        second,
        separatorText: input.separatorText,
        separatorRuns,
        firstTarget,
        secondTarget,
      });
    }
    case "mergeParagraphs": {
      const firstSource = ownSourceOperand(input.firstSource, sourceSnapshot);
      const secondSource = ownSourceOperand(input.secondSource, sourceSnapshot);
      const firstSourceBlock = resolvedDocxSourceOperandBlock(firstSource, sourceSnapshot);
      const secondSourceBlock = resolvedDocxSourceOperandBlock(secondSource, sourceSnapshot);
      const first = compileRange(input.first);
      const second = compileRange(input.second);
      const separatorRuns = ownRuns(input.separatorText, input.separatorRuns);
      const target = ownParagraphTarget(input.target);
      const targetText = `${first.targetText}${input.separatorText}${second.targetText}`;
      if (
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
          { text: first.targetText, runs: ownRuns(first.targetText, input.first.targetRuns) },
          { text: input.separatorText, runs: separatorRuns },
          { text: second.targetText, runs: ownRuns(second.targetText, input.second.targetRuns) },
        ]),
      );
      return Object.freeze({
        type: "mergeParagraphs",
        firstSource,
        secondSource,
        first,
        second,
        separatorText: input.separatorText,
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
    case "setParagraphProperties":
      if (
        docxParagraphPropertiesEqual(
          docxParagraphPropertiesFromBlock(
            resolvedDocxSourceOperandBlock(input.source, sourceSnapshot),
          ),
          input.targetProperties,
        )
      ) {
        return panic("A DOCX comparison paragraph-property instruction contains no change");
      }
      return Object.freeze({
        type: "setParagraphProperties",
        source: ownSourceOperand(input.source, sourceSnapshot),
        targetProperties: ownParagraphProperties(input.targetProperties),
      });
    case "insertTable":
      return Object.freeze({
        type: "insertTable",
        anchor: ownStructuralInsertionAnchor(input.anchor),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
      });
    case "deleteTable":
      return Object.freeze({
        type: "deleteTable",
        source: ownSourceOperand(input.source, sourceSnapshot),
        baseTableIndex: ownIndex("baseTableIndex", input.baseTableIndex),
      });
    case "replaceTable":
      return Object.freeze({
        type: "replaceTable",
        source: ownSourceOperand(input.source, sourceSnapshot),
        baseTableIndex: ownIndex("baseTableIndex", input.baseTableIndex),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
      });
    case "insertTableRow":
      return Object.freeze({
        type: "insertTableRow",
        anchor: ownStructuralInsertionAnchor(input.anchor),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
        targetRowIndex: ownIndex("targetRowIndex", input.targetRowIndex),
      });
    case "deleteTableRow":
      return Object.freeze({
        type: "deleteTableRow",
        source: ownSourceOperand(input.source, sourceSnapshot),
        baseTableIndex: ownIndex("baseTableIndex", input.baseTableIndex),
        baseRowIndex: ownIndex("baseRowIndex", input.baseRowIndex),
      });
    case "insertTableColumn":
      return Object.freeze({
        type: "insertTableColumn",
        anchor: ownStructuralInsertionAnchor(input.anchor),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
        targetColumnIndex: ownIndex("targetColumnIndex", input.targetColumnIndex),
        cellTexts: ownCellTexts(input.cellTexts),
      });
    case "deleteTableColumn":
      return Object.freeze({
        type: "deleteTableColumn",
        source: ownSourceOperand(input.source, sourceSnapshot),
        baseTableIndex: ownIndex("baseTableIndex", input.baseTableIndex),
        baseColumnIndex: ownIndex("baseColumnIndex", input.baseColumnIndex),
      });
    case "matchTableGeometry":
      return Object.freeze({
        type: "matchTableGeometry",
        pairings: ownTableGeometryPairings(input.pairings),
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
      return panic("A DOCX comparison program exceeds its instruction limit", {
        limit: MAX_DOCX_COMPARISON_INSTRUCTIONS,
        actual: inputs.length,
      });
    }
    const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
    this.#comparison = comparison;
    this.#sourceSnapshot = baseSnapshot;
    this.#targetSnapshot = targetSnapshot;
    this.#instructions = Object.freeze(
      inputs.map((input) => compileInstruction(input, baseSnapshot)),
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
