import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import type {
  FolioAIBlockParagraphProperties,
  FolioAIBlockTableLocation,
} from "../../ai-edits/types";
import type { TableGeometryPairing } from "../../ai-edits/table-geometry";
import type { TextFormatting } from "../../types/document";
import type { FolioContentTextSegment } from "../../compare/content";
import type { FolioContentStructuralBoundary } from "../../compare/content-types";

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

export type DocxComparisonBlockExpectation = {
  readonly blockId: string;
  readonly kind: string;
  readonly text: string;
  readonly paragraphProperties: Readonly<FolioAIBlockParagraphProperties>;
  readonly structuralBoundaries: readonly FolioContentStructuralBoundary[];
  readonly containerPath: readonly { readonly kind: string; readonly id: string }[];
  readonly table?: Readonly<FolioAIBlockTableLocation>;
};

export type DocxComparisonInsertionAnchor = {
  readonly blockId: string;
  readonly position: "after" | "before";
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
      readonly source: DocxComparisonBlockExpectation;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlanInput;
    }
  | {
      readonly type: "formatText";
      readonly source: DocxComparisonBlockExpectation;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlanInput;
    }
  | {
      readonly type: "insertParagraph";
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly target: DocxComparisonParagraphTargetInput;
    }
  | {
      readonly type: "deleteParagraph";
      readonly source: DocxComparisonBlockExpectation;
    }
  | {
      readonly type: "moveParagraph";
      readonly source: DocxComparisonBlockExpectation;
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly target: DocxComparisonParagraphTargetInput;
    }
  | {
      readonly type: "splitParagraph";
      readonly source: DocxComparisonBlockExpectation;
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
      readonly firstSource: DocxComparisonBlockExpectation;
      readonly secondSource: DocxComparisonBlockExpectation;
      readonly first: DocxComparisonRangePlanInput;
      readonly second: DocxComparisonRangePlanInput;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
      readonly target: DocxComparisonParagraphTargetInput;
    }
  | {
      readonly type: "mergeTerminalCarrier";
      readonly source: DocxComparisonBlockExpectation;
    }
  | {
      readonly type: "setParagraphProperties";
      readonly source: DocxComparisonBlockExpectation;
      readonly targetProperties: Readonly<FolioAIBlockParagraphProperties>;
    }
  | {
      readonly type: "insertTable";
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly targetTableIndex: number;
    }
  | {
      readonly type: "deleteTable";
      readonly source: DocxComparisonBlockExpectation;
      readonly baseTableIndex: number;
    }
  | {
      readonly type: "replaceTable";
      readonly source: DocxComparisonBlockExpectation;
      readonly baseTableIndex: number;
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly targetTableIndex: number;
    }
  | {
      readonly type: "insertTableRow";
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly targetTableIndex: number;
      readonly targetRowIndex: number;
    }
  | {
      readonly type: "deleteTableRow";
      readonly source: DocxComparisonBlockExpectation;
      readonly baseTableIndex: number;
      readonly baseRowIndex: number;
    }
  | {
      readonly type: "insertTableColumn";
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly targetTableIndex: number;
      readonly targetColumnIndex: number;
      readonly cellTexts: readonly string[];
    }
  | {
      readonly type: "deleteTableColumn";
      readonly source: DocxComparisonBlockExpectation;
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
      readonly source: DocxComparisonBlockExpectation;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlan;
    }
  | {
      readonly type: "formatText";
      readonly source: DocxComparisonBlockExpectation;
      readonly sourceStartOffset: number;
      readonly range: DocxComparisonRangePlan;
    }
  | {
      readonly type: "insertParagraph";
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly target: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "deleteParagraph";
      readonly source: DocxComparisonBlockExpectation;
    }
  | {
      readonly type: "moveParagraph";
      readonly source: DocxComparisonBlockExpectation;
      readonly anchor: DocxComparisonInsertionAnchor;
      readonly target: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "splitParagraph";
      readonly source: DocxComparisonBlockExpectation;
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
      readonly firstSource: DocxComparisonBlockExpectation;
      readonly secondSource: DocxComparisonBlockExpectation;
      readonly first: DocxComparisonRangePlan;
      readonly second: DocxComparisonRangePlan;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
      readonly target: DocxComparisonParagraphTarget;
    }
  | {
      readonly type: "mergeTerminalCarrier";
      readonly source: DocxComparisonBlockExpectation;
    }
  | {
      readonly type: "setParagraphProperties";
      readonly source: DocxComparisonBlockExpectation;
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

const sameFormatting = (
  left: Readonly<TextFormatting>,
  right: Readonly<TextFormatting>,
): boolean =>
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
      const activeChange = changedRangeAt(
        authoredChanges,
        changeIndex,
        baseOffset,
        revisedOffset,
      );
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
        .filter(
          ({ field }) =>
            !sameValue(sourceRun.formatting[field], targetRun.formatting[field]),
        )
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

const ownBlockExpectation = (
  source: DocxComparisonBlockExpectation,
): DocxComparisonBlockExpectation => {
  if (source.blockId.length === 0 || source.kind.length === 0) {
    return panic("A DOCX comparison source expectation has no block identity");
  }
  const structuralBoundaries = structuredClone(source.structuralBoundaries);
  const containerPath = structuredClone(source.containerPath);
  freezeRecursively(structuralBoundaries);
  freezeRecursively(containerPath);
  return Object.freeze({
    blockId: source.blockId,
    kind: source.kind,
    text: source.text,
    paragraphProperties: ownParagraphProperties(source.paragraphProperties),
    structuralBoundaries,
    containerPath,
    ...(source.table !== undefined && { table: ownTableLocation(source.table) }),
  });
};

const ownInsertionAnchor = (
  anchor: DocxComparisonInsertionAnchor,
): DocxComparisonInsertionAnchor => {
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
  source: DocxComparisonBlockExpectation,
  sourceStartOffset: number,
  range: DocxComparisonRangePlan,
): void => {
  if (
    !Number.isSafeInteger(sourceStartOffset) ||
    sourceStartOffset < 0 ||
    source.text.slice(sourceStartOffset, sourceStartOffset + range.sourceText.length) !==
      range.sourceText
  ) {
    return panic("A DOCX comparison range does not name its exact source block", {
      blockId: source.blockId,
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
): DocxComparisonInstruction => {
  switch (input.type) {
    case "replaceText": {
      const source = ownBlockExpectation(input.source);
      const range = compileRange(input.range);
      assertRangeMatchesSource(source, input.sourceStartOffset, range);
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
      const source = ownBlockExpectation(input.source);
      const range = compileRange(input.range);
      assertRangeMatchesSource(source, input.sourceStartOffset, range);
      if (
        range.sourceText !== range.targetText ||
        range.fragments.some((fragment) => fragment.type !== "equal") ||
        range.fragments.every(({ changedProperties }) => changedProperties.length === 0)
      ) {
        return panic("A DOCX comparison formatting instruction must contain only formatting changes");
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
        anchor: ownInsertionAnchor(input.anchor),
        target: ownParagraphTarget(input.target),
      });
    case "deleteParagraph":
      return Object.freeze({
        type: "deleteParagraph",
        source: ownBlockExpectation(input.source),
      });
    case "moveParagraph":
      return Object.freeze({
        type: "moveParagraph",
        source: ownBlockExpectation(input.source),
        anchor: ownInsertionAnchor(input.anchor),
        target: ownParagraphTarget(input.target),
      });
    case "splitParagraph": {
      const source = ownBlockExpectation(input.source);
      const first = compileRange(input.first);
      const second = compileRange(input.second);
      const separatorRuns = ownRuns(input.separatorText, input.separatorRuns);
      const firstTarget = ownParagraphTarget(input.firstTarget);
      const secondTarget = ownParagraphTarget(input.secondTarget);
      if (
        input.offset !== first.sourceText.length ||
        source.text !== `${first.sourceText}${input.separatorText}${second.sourceText}` ||
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
      const firstSource = ownBlockExpectation(input.firstSource);
      const secondSource = ownBlockExpectation(input.secondSource);
      const first = compileRange(input.first);
      const second = compileRange(input.second);
      const separatorRuns = ownRuns(input.separatorText, input.separatorRuns);
      const target = ownParagraphTarget(input.target);
      const targetText = `${first.targetText}${input.separatorText}${second.targetText}`;
      if (
        firstSource.text !== first.sourceText ||
        secondSource.text !== second.sourceText ||
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
    case "mergeTerminalCarrier":
      return Object.freeze({
        type: "mergeTerminalCarrier",
        source: ownBlockExpectation(input.source),
      });
    case "setParagraphProperties":
      if (sameValue(input.source.paragraphProperties, input.targetProperties)) {
        return panic("A DOCX comparison paragraph-property instruction contains no change");
      }
      return Object.freeze({
        type: "setParagraphProperties",
        source: ownBlockExpectation(input.source),
        targetProperties: ownParagraphProperties(input.targetProperties),
      });
    case "insertTable":
      return Object.freeze({
        type: "insertTable",
        anchor: ownInsertionAnchor(input.anchor),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
      });
    case "deleteTable":
      return Object.freeze({
        type: "deleteTable",
        source: ownBlockExpectation(input.source),
        baseTableIndex: ownIndex("baseTableIndex", input.baseTableIndex),
      });
    case "replaceTable":
      return Object.freeze({
        type: "replaceTable",
        source: ownBlockExpectation(input.source),
        baseTableIndex: ownIndex("baseTableIndex", input.baseTableIndex),
        anchor: ownInsertionAnchor(input.anchor),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
      });
    case "insertTableRow":
      return Object.freeze({
        type: "insertTableRow",
        anchor: ownInsertionAnchor(input.anchor),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
        targetRowIndex: ownIndex("targetRowIndex", input.targetRowIndex),
      });
    case "deleteTableRow":
      return Object.freeze({
        type: "deleteTableRow",
        source: ownBlockExpectation(input.source),
        baseTableIndex: ownIndex("baseTableIndex", input.baseTableIndex),
        baseRowIndex: ownIndex("baseRowIndex", input.baseRowIndex),
      });
    case "insertTableColumn":
      return Object.freeze({
        type: "insertTableColumn",
        anchor: ownInsertionAnchor(input.anchor),
        targetTableIndex: ownIndex("targetTableIndex", input.targetTableIndex),
        targetColumnIndex: ownIndex("targetColumnIndex", input.targetColumnIndex),
        cellTexts: ownCellTexts(input.cellTexts),
      });
    case "deleteTableColumn":
      return Object.freeze({
        type: "deleteTableColumn",
        source: ownBlockExpectation(input.source),
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
  readonly #instructions: readonly DocxComparisonInstruction[];
  #state: "ready" | "consumed" = "ready";

  private constructor(inputs: readonly DocxComparisonInstructionInput[]) {
    if (inputs.length > MAX_DOCX_COMPARISON_INSTRUCTIONS) {
      return panic("A DOCX comparison program exceeds its instruction limit", {
        limit: MAX_DOCX_COMPARISON_INSTRUCTIONS,
        actual: inputs.length,
      });
    }
    this.#instructions = Object.freeze(inputs.map(compileInstruction));
  }

  static create(inputs: readonly DocxComparisonInstructionInput[]): DocxComparisonProgram {
    return new DocxComparisonProgram(inputs);
  }

  get size(): number {
    return this.#instructions.length;
  }

  consume(): readonly DocxComparisonInstruction[] {
    if (this.#state !== "ready") {
      return panic("A DOCX comparison program was consumed more than once");
    }
    this.#state = "consumed";
    return this.#instructions;
  }
}
