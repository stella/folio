import { TEXT_FORMATTING_PROPERTY_DESCRIPTORS } from "@stll/docx-core/model";
import { panic } from "better-result";

import type { FolioAIEditOperation } from "../ai-edits/types";
import type { TextFormatting } from "../types/document";
import type { FolioContentTextSegment } from "./content";

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

type DocxComparisonReplacementPlanInput = {
  readonly type: "replacement";
  readonly range: DocxComparisonRangePlanInput;
};

type DocxComparisonFormatPlanInput = {
  readonly type: "format";
  readonly range: DocxComparisonRangePlanInput;
};

type DocxComparisonInsertionPlanInput = {
  readonly type: "insertion";
  readonly targetText: string;
  readonly targetRuns: readonly DocxAuthoredRun[];
};

type DocxComparisonSplitPlanInput = {
  readonly type: "split";
  readonly first: DocxComparisonRangePlanInput;
  readonly second: DocxComparisonRangePlanInput;
  readonly separatorText: string;
  readonly separatorRuns: readonly DocxAuthoredRun[];
};

type DocxComparisonSemanticMergePlanInput = {
  readonly type: "merge";
  readonly encoding: "semantic";
  readonly first: DocxComparisonRangePlanInput;
  readonly second: DocxComparisonRangePlanInput;
  readonly separatorText: string;
  readonly separatorRuns: readonly DocxAuthoredRun[];
};

type DocxComparisonTerminalMergePlanInput = {
  readonly type: "merge";
  readonly encoding: "terminal-deletion-carrier";
};

type DocxComparisonNoPlanInput = {
  readonly type: "none";
};

type DocxComparisonTableTemplatePlanInput = {
  readonly type: "table-template";
  readonly targetTableIndex: number;
  readonly targetRowIndex?: never;
};

type DocxComparisonTableRowTemplatePlanInput = {
  readonly type: "table-template";
  readonly targetTableIndex: number;
  readonly targetRowIndex: number;
};

type DocxComparisonPlainColumnPlanInput = {
  readonly type: "plain-column";
};

type DocxComparisonExactOperationPlanInput =
  | DocxComparisonReplacementPlanInput
  | DocxComparisonFormatPlanInput
  | DocxComparisonInsertionPlanInput
  | DocxComparisonSplitPlanInput
  | DocxComparisonSemanticMergePlanInput
  | DocxComparisonTerminalMergePlanInput;

type DocxComparisonOperationPlanByType = {
  readonly replaceInBlock: DocxComparisonReplacementPlanInput;
  readonly replaceRange: DocxComparisonReplacementPlanInput;
  readonly commentOnRange: never;
  readonly formatRange: DocxComparisonFormatPlanInput;
  readonly insertAfterBlock: DocxComparisonInsertionPlanInput;
  readonly insertBeforeBlock: DocxComparisonInsertionPlanInput;
  readonly replaceBlock: DocxComparisonReplacementPlanInput;
  readonly deleteBlock: DocxComparisonNoPlanInput;
  readonly splitBlock: DocxComparisonSplitPlanInput;
  readonly mergeBlockWithNext:
    | DocxComparisonSemanticMergePlanInput
    | DocxComparisonTerminalMergePlanInput;
  readonly setBlockParagraphProperties: DocxComparisonNoPlanInput;
  readonly insertTable: DocxComparisonTableTemplatePlanInput;
  readonly deleteTable: DocxComparisonNoPlanInput;
  readonly commentOnBlock: never;
  readonly insertSignatureTable: never;
  readonly insertTableRow: DocxComparisonTableRowTemplatePlanInput;
  readonly deleteTableRow: DocxComparisonNoPlanInput;
  readonly insertTableColumn: DocxComparisonPlainColumnPlanInput;
  readonly deleteTableColumn: DocxComparisonNoPlanInput;
  readonly mergeTableCells: never;
  readonly splitTableCell: never;
};

type FolioAIEditOperationType = FolioAIEditOperation["type"];

/**
 * One comparison operation and the only semantic plan valid for its exact
 * operation discriminant. This is the planner's single owner: public
 * operations, exact text/run application, and table templates are derived
 * from this record rather than joined by ids later.
 */
export type DocxComparisonOperationInput = {
  readonly [Type in FolioAIEditOperationType]: DocxComparisonOperationPlanByType[Type] extends never
    ? never
    : {
        readonly operation: Extract<FolioAIEditOperation, { readonly type: Type }>;
        readonly plan: DocxComparisonOperationPlanByType[Type];
      };
}[FolioAIEditOperationType];

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

export type DocxComparisonOperationPlan =
  | { readonly type: "replacement"; readonly range: DocxComparisonRangePlan }
  | { readonly type: "format"; readonly range: DocxComparisonRangePlan }
  | {
      readonly type: "insertion";
      readonly targetText: string;
      readonly targetRuns: readonly DocxAuthoredRun[];
    }
  | {
      readonly type: "split";
      readonly first: DocxComparisonRangePlan;
      readonly second: DocxComparisonRangePlan;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
    }
  | {
      readonly type: "merge";
      readonly encoding: "semantic";
      readonly first: DocxComparisonRangePlan;
      readonly second: DocxComparisonRangePlan;
      readonly separatorText: string;
      readonly separatorRuns: readonly DocxAuthoredRun[];
    }
  | { readonly type: "merge"; readonly encoding: "terminal-deletion-carrier" }
  | DocxComparisonNoPlanInput
  | DocxComparisonTableTemplatePlanInput
  | DocxComparisonTableRowTemplatePlanInput
  | DocxComparisonPlainColumnPlanInput;

export type DocxComparisonOperationClaim = DocxComparisonOperationPlan;

export type BorrowedDocxComparisonOperation = {
  readonly semantics: DocxComparisonOperationClaim;
  /** Consume the proof only after the operation produced its intended edit. */
  readonly commitApplied: () => void;
};

type PlanRequirement = DocxComparisonOperationPlan["type"] | "forbidden";

const PLAN_REQUIREMENT_BY_OPERATION = Object.freeze({
  replaceInBlock: "replacement",
  replaceRange: "replacement",
  commentOnRange: "forbidden",
  formatRange: "format",
  insertAfterBlock: "insertion",
  insertBeforeBlock: "insertion",
  replaceBlock: "replacement",
  deleteBlock: "none",
  splitBlock: "split",
  mergeBlockWithNext: "merge",
  setBlockParagraphProperties: "none",
  insertTable: "table-template",
  deleteTable: "none",
  commentOnBlock: "forbidden",
  insertSignatureTable: "forbidden",
  insertTableRow: "table-template",
  deleteTableRow: "none",
  insertTableColumn: "plain-column",
  deleteTableColumn: "none",
  mergeTableCells: "forbidden",
  splitTableCell: "forbidden",
} as const satisfies Readonly<Record<FolioAIEditOperation["type"], PlanRequirement>>);

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

type DocxComparisonOperationPlanInput = DocxComparisonOperationInput["plan"];

const compilePlan = (input: DocxComparisonOperationPlanInput): DocxComparisonOperationPlan => {
  switch (input.type) {
    case "replacement":
      return Object.freeze({ type: "replacement", range: compileRange(input.range) });
    case "format": {
      const range = compileRange(input.range);
      if (range.fragments.some((fragment) => fragment.type !== "equal")) {
        return panic("A DOCX formatting plan may not change text");
      }
      return Object.freeze({ type: "format", range });
    }
    case "insertion":
      return Object.freeze({
        type: "insertion",
        targetText: input.targetText,
        targetRuns: ownRuns(input.targetText, input.targetRuns),
      });
    case "split":
      return Object.freeze({
        type: "split",
        first: compileRange(input.first),
        second: compileRange(input.second),
        separatorText: input.separatorText,
        separatorRuns: ownRuns(input.separatorText, input.separatorRuns),
      });
    case "merge":
      return input.encoding === "semantic"
        ? Object.freeze({
            type: "merge",
            encoding: "semantic",
            first: compileRange(input.first),
            second: compileRange(input.second),
            separatorText: input.separatorText,
            separatorRuns: ownRuns(input.separatorText, input.separatorRuns),
          })
        : Object.freeze({ type: "merge", encoding: "terminal-deletion-carrier" });
    case "none":
      return Object.freeze({ type: "none" });
    case "table-template":
      return Object.freeze({
        type: "table-template",
        targetTableIndex: input.targetTableIndex,
        ...(input.targetRowIndex !== undefined && { targetRowIndex: input.targetRowIndex }),
      });
    case "plain-column":
      return Object.freeze({ type: "plain-column" });
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX comparison operation plan", { plan: unreachable });
    }
  }
};

type PendingPlan = {
  readonly operationType: FolioAIEditOperation["type"];
  readonly plan: DocxComparisonOperationPlan;
};

export type DocxComparisonTableTemplateRequest = {
  readonly operationId: string;
  readonly targetTableIndex: number;
  readonly targetRowIndex?: number;
};

/**
 * Proof-carrying, one-use operation semantics available only to the DOCX
 * comparison applier. Construction proves every operation has exactly the
 * plan its discriminant requires; claiming consumes that proof.
 */
export class DocxComparisonOperationBatch {
  readonly #pending: Map<string, PendingPlan>;
  readonly #borrowed = new Set<string>();
  readonly operations: readonly FolioAIEditOperation[];
  readonly tableTemplateRequests: readonly DocxComparisonTableTemplateRequest[];
  #state: "open" | "abandoned" | "finalized" = "open";

  private constructor(inputs: readonly DocxComparisonOperationInput[]) {
    const pending = new Map<string, PendingPlan>();
    const operations: FolioAIEditOperation[] = [];
    const tableTemplateRequests: DocxComparisonTableTemplateRequest[] = [];
    for (const input of inputs) {
      const operation = structuredClone(input.operation);
      freezeRecursively(operation);
      if (pending.has(operation.id)) {
        return panic("A DOCX comparison plan contains duplicate operation ids", {
          operationId: operation.id,
        });
      }
      const requirement = PLAN_REQUIREMENT_BY_OPERATION[operation.type];
      if (requirement === "forbidden" || input.plan.type !== requirement) {
        return panic("A DOCX comparison operation is missing its exact plan", {
          operationId: operation.id,
          operationType: operation.type,
          requirement,
          received: input.plan.type,
        });
      }
      const plan = compilePlan(input.plan);
      if (operation.type === "insertTable") {
        if (plan.type !== "table-template" || plan.targetRowIndex !== undefined) {
          return panic("A table insertion received a row template plan", {
            operationId: operation.id,
          });
        }
      }
      if (operation.type === "insertTableRow") {
        if (plan.type !== "table-template" || plan.targetRowIndex === undefined) {
          return panic("A table-row insertion received a whole-table template plan", {
            operationId: operation.id,
          });
        }
      }
      if (plan.type === "table-template") {
        tableTemplateRequests.push(
          Object.freeze({
            operationId: operation.id,
            targetTableIndex: plan.targetTableIndex,
            ...(plan.targetRowIndex !== undefined && { targetRowIndex: plan.targetRowIndex }),
          }),
        );
      }
      operations.push(operation);
      pending.set(
        operation.id,
        Object.freeze({ operationType: operation.type, plan }),
      );
    }
    this.#pending = pending;
    this.operations = Object.freeze(operations);
    this.tableTemplateRequests = Object.freeze(tableTemplateRequests);
  }

  static create(inputs: readonly DocxComparisonOperationInput[]): DocxComparisonOperationBatch {
    return new DocxComparisonOperationBatch(inputs);
  }

  borrow(operation: FolioAIEditOperation): BorrowedDocxComparisonOperation {
    if (this.#state !== "open") {
      return panic("A closed DOCX comparison operation batch was reused", {
        state: this.#state,
      });
    }
    const pending = this.#pending.get(operation.id);
    if (!pending || pending.operationType !== operation.type || this.#borrowed.has(operation.id)) {
      return panic("A DOCX comparison operation did not match its one-use plan", {
        operationId: operation.id,
        operationType: operation.type,
        expectedType: pending?.operationType,
        borrowed: this.#borrowed.has(operation.id),
      });
    }
    this.#borrowed.add(operation.id);
    let state: "borrowed" | "committed" = "borrowed";
    return Object.freeze({
      semantics: pending.plan,
      commitApplied: () => {
        if (state !== "borrowed" || this.#state !== "open") {
          return panic("A DOCX comparison operation proof was committed more than once", {
            operationId: operation.id,
            batchState: this.#state,
          });
        }
        state = "committed";
        this.#borrowed.delete(operation.id);
        this.#pending.delete(operation.id);
      },
    });
  }

  /** Close a batch whose application returned a typed refusal before dispatch. */
  abandon(): void {
    if (this.#state !== "open") {
      return panic("A DOCX comparison operation batch was closed more than once", {
        state: this.#state,
      });
    }
    this.#state = "abandoned";
  }

  finalize(): void {
    if (this.#state !== "open") {
      return panic("A DOCX comparison operation batch was finalized after it closed", {
        state: this.#state,
      });
    }
    const remaining = this.#pending.keys().next();
    if (!remaining.done) {
      return panic("A DOCX comparison operation plan was not consumed completely", {
        operationId: remaining.value,
        borrowed: this.#borrowed.has(remaining.value),
      });
    }
    this.#state = "finalized";
  }
}
