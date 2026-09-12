import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { canJoin, canSplit } from "prosemirror-transform";

import { applyBlockParagraphProperties, withRotatedAddedFinalBreaks } from "../../ai-edits/apply";
import { buildCleanBlockText } from "../../ai-edits/clean-text";
import { FolioStableBlockResolver } from "./stable-block-resolution";
import { storyTablesOf } from "../../ai-edits/snapshot";
import { expectParagraphAttrs } from "../../prosemirror/attrs";
import { hasSerializableParagraphPropertyChange } from "../../prosemirror/commands/propertyChangeScope";
import { paragraphEndsItsContainer } from "../../prosemirror/containerFinalParagraph";
import { markStructuralChange } from "../../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { requestDeterministicParaIds } from "../../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { getDocumentNumbering } from "../../prosemirror/plugins/documentNumbering";
import { getDocumentStyleResolver } from "../../prosemirror/plugins/documentStyles";
import {
  COMPARE_DOCX_PREFLIGHT_REASONS,
  type CompareDocxPreflightReason,
} from "../../compare/types";
import { stripBlockIdentityAttrs } from "../../ai-edits/block-identity";
import type { FolioRevisionStamp } from "../../ai-edits/apply";
import {
  DocxComparisonProgram,
  type DocxComparisonInstruction,
  type DocxComparisonParagraphTarget,
  type DocxComparisonSourceOperand,
} from "./docx-program";
import {
  applyExactDirectFormatting,
  applyPreflightedDocxTextRange,
  preflightDocxTextRange,
  type DocxTextRangePreflight,
} from "./docx-text-executor";
import {
  executeTableGeometryProgram,
  preflightTableGeometry,
  type TableGeometryExecutionIssue,
  type TableGeometryExecutionReceipt,
  type TableGeometryPreflightResult,
  type TableGeometryUnsupportedIssue,
} from "./table-geometry-program";
import {
  resolvedDocxOperationSnapshot,
  resolvedDocxSourceDocument,
  resolvedDocxSourceOperandBlock,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";

export const DOCX_COMPARISON_PREFLIGHT_REASONS = COMPARE_DOCX_PREFLIGHT_REASONS;

export type DocxComparisonPreflightReason = CompareDocxPreflightReason;

export type DocxComparisonPreflightIssue = {
  readonly instructionIndex: number;
  readonly instructionType: DocxComparisonInstruction["type"];
  readonly reason: DocxComparisonPreflightReason;
  readonly blockId?: string;
  readonly tableGeometry?: TableGeometryUnsupportedIssue;
};

type ResolvedBlock = {
  readonly node: PMNode;
  readonly from: number;
  readonly to: number;
};

type ResolvedParagraphBoundary = ResolvedBlock & {
  readonly insertionPosition: number;
};

type ReadyTextRange = Extract<DocxTextRangePreflight, { readonly type: "ready" }>;
type ReadyTableGeometryProgram = Extract<
  TableGeometryPreflightResult,
  { readonly status: "ready" }
>["program"];

type PreparedSourceSchedule = {
  readonly phase: "source";
  readonly from: number;
  readonly to: number;
};

type PreparedInsertionSchedule = {
  readonly phase: "insertion";
  readonly position: number;
};

type PreparedInstruction =
  | {
      readonly type: "replaceText" | "formatText";
      readonly source: ResolvedBlock;
      readonly sourceStartOffset: number;
      readonly range: ReadyTextRange;
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "replaceText" | "formatText" }
      >;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "insertParagraph";
      readonly boundary: ResolvedParagraphBoundary;
      readonly target: DocxComparisonParagraphTarget;
      readonly originalIndex: number;
      readonly schedule: PreparedInsertionSchedule;
    }
  | {
      readonly type: "deleteParagraph";
      readonly source: ResolvedBlock;
      readonly semantic: Extract<DocxComparisonInstruction, { readonly type: "deleteParagraph" }>;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "deleteTrailingParagraphs";
      readonly chainStart: ResolvedBlock;
      readonly deleted: readonly [ResolvedBlock, ...ResolvedBlock[]];
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "deleteTrailingParagraphs" }
      >;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "moveParagraph";
      readonly source: ResolvedBlock;
      readonly successor: ResolvedBlock;
      readonly boundary: ResolvedParagraphBoundary;
      readonly target: DocxComparisonParagraphTarget;
      readonly originalIndex: number;
      readonly sourceSchedule: PreparedSourceSchedule;
      readonly destinationSchedule: PreparedInsertionSchedule;
    }
  | {
      readonly type: "moveTerminalParagraph";
      readonly predecessor: ResolvedBlock;
      readonly source: ResolvedBlock;
      readonly boundary: ResolvedParagraphBoundary;
      readonly target: DocxComparisonParagraphTarget;
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "moveTerminalParagraph" }
      >;
      readonly originalIndex: number;
      readonly sourceSchedule: PreparedSourceSchedule;
      readonly destinationSchedule: PreparedInsertionSchedule;
    }
  | {
      readonly type: "splitParagraph";
      readonly source: ResolvedBlock;
      readonly splitPosition: number;
      readonly first: ReadyTextRange;
      readonly second: ReadyTextRange;
      readonly semantic: Extract<DocxComparisonInstruction, { readonly type: "splitParagraph" }>;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "mergeParagraphs";
      readonly firstSource: ResolvedBlock;
      readonly secondSource: ResolvedBlock;
      readonly first: ReadyTextRange;
      readonly second: ReadyTextRange;
      readonly semantic: Extract<DocxComparisonInstruction, { readonly type: "mergeParagraphs" }>;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "setParagraphProperties";
      readonly source: ResolvedBlock;
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "setParagraphProperties" }
      >;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "matchTableGeometry";
      readonly program: ReadyTableGeometryProgram;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "geometry" };
    };

type PreparedSourceInstruction = Extract<
  PreparedInstruction,
  { readonly schedule: { readonly phase: "source" } }
>;
type PreparedInsertionInstruction = Extract<
  PreparedInstruction,
  { readonly type: "insertParagraph" }
>;
type PreparedMoveInstruction = Extract<
  PreparedInstruction,
  { readonly type: "moveParagraph" | "moveTerminalParagraph" }
>;
type PreparedGeometryInstruction = Extract<
  PreparedInstruction,
  { readonly schedule: { readonly phase: "geometry" } }
>;

type PreparedInsertionMember = {
  readonly instruction: PreparedInsertionInstruction | PreparedMoveInstruction;
};

type PreparedInsertionRun = {
  readonly type: "insertionRun";
  readonly position: number;
  readonly members: readonly [PreparedInsertionMember, ...PreparedInsertionMember[]];
};

type PreparedSourceTask =
  | {
      readonly type: "sourceInstruction";
      readonly position: number;
      readonly from: number;
      readonly instruction: PreparedSourceInstruction;
    }
  | {
      readonly type: "moveSource";
      readonly position: number;
      readonly from: number;
      readonly instruction: PreparedMoveInstruction;
    };

type PreparedExecutionTask = PreparedInsertionRun | PreparedSourceTask;

const executionTasks = (
  instructions: readonly PreparedInstruction[],
): readonly PreparedExecutionTask[] => {
  const sourceTasks: PreparedSourceTask[] = [];
  const insertionMembersByPosition = new Map<number, PreparedInsertionMember[]>();
  for (const instruction of instructions) {
    switch (instruction.type) {
      case "insertParagraph": {
        const members = insertionMembersByPosition.get(instruction.schedule.position) ?? [];
        members.push({ instruction });
        insertionMembersByPosition.set(instruction.schedule.position, members);
        break;
      }
      case "moveParagraph":
      case "moveTerminalParagraph": {
        sourceTasks.push({
          type: "moveSource",
          position: instruction.sourceSchedule.to,
          from: instruction.sourceSchedule.from,
          instruction,
        });
        const members =
          insertionMembersByPosition.get(instruction.destinationSchedule.position) ?? [];
        members.push({ instruction });
        insertionMembersByPosition.set(instruction.destinationSchedule.position, members);
        break;
      }
      case "matchTableGeometry":
        break;
      default:
        sourceTasks.push({
          type: "sourceInstruction",
          position: instruction.schedule.to,
          from: instruction.schedule.from,
          instruction,
        });
        break;
    }
  }
  const insertionRuns: PreparedInsertionRun[] = [];
  for (const [position, members] of insertionMembersByPosition) {
    const ordered = members.toSorted(
      (left, right) => left.instruction.originalIndex - right.instruction.originalIndex,
    );
    const first = ordered.at(0) ?? panic("A prepared insertion run has no member");
    insertionRuns.push({
      type: "insertionRun",
      position,
      members: Object.freeze([first, ...ordered.slice(1)]),
    });
  }
  return Object.freeze(
    [...sourceTasks, ...insertionRuns].toSorted((left, right) => {
      const byPosition = right.position - left.position;
      if (byPosition !== 0) return byPosition;
      // A boundary insertion at the exact end of a source range runs first:
      // right-to-left execution then leaves every source coordinate original.
      if (left.type === "insertionRun" || right.type === "insertionRun") {
        return left.type === "insertionRun" ? -1 : 1;
      }
      const byStart = right.from - left.from;
      if (byStart !== 0) return byStart;
      return right.instruction.originalIndex - left.instruction.originalIndex;
    }),
  );
};

type ResolveExpectedBlockOptions = {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly resolver: FolioStableBlockResolver;
  readonly source: DocxComparisonSourceOperand;
};

const resolveExpectedBlock = ({
  snapshot,
  resolver,
  source,
}: ResolveExpectedBlockOptions):
  | { readonly type: "ready"; readonly block: ResolvedBlock }
  | { readonly type: "unsupported"; readonly reason: DocxComparisonPreflightReason } => {
  const canonicalBlock = resolvedDocxSourceOperandBlock(source, snapshot);
  const resolved = resolver.resolve(canonicalBlock.identity.id);
  if (resolved.type === "unsupported") return resolved;
  return {
    type: "ready",
    block: Object.freeze({
      node: resolved.blockNode,
      from: resolved.blockFrom,
      to: resolved.blockTo,
    }),
  };
};

const resolveParagraphBoundary = (
  snapshot: ResolvedDocxStorySnapshot,
  resolver: FolioStableBlockResolver,
  boundary: Extract<DocxComparisonInstruction, { readonly type: "insertParagraph" }>["boundary"],
): ResolvedParagraphBoundary | null => {
  const resolved = resolveExpectedBlock({
    snapshot,
    resolver,
    source: boundary.paragraph,
  });
  if (resolved.type === "unsupported") return null;
  const insertionPosition =
    boundary.type === "afterParagraph" ? resolved.block.to : resolved.block.from;
  return Object.freeze({
    ...resolved.block,
    insertionPosition,
  });
};

const issue = (
  instruction: DocxComparisonInstruction,
  instructionIndex: number,
  reason: DocxComparisonPreflightReason,
  blockId?: string,
  tableGeometry?: TableGeometryUnsupportedIssue,
): DocxComparisonPreflightIssue =>
  Object.freeze({
    instructionIndex,
    instructionType: instruction.type,
    reason,
    ...(blockId !== undefined && { blockId }),
    ...(tableGeometry !== undefined && { tableGeometry }),
  });

const structuralRangeIssueReason = (
  ranges: readonly DocxTextRangePreflight[],
): DocxComparisonPreflightReason => {
  if (
    ranges.some((range) => range.type === "unsupported" && range.reason === "pending-run-change")
  ) {
    return "pending-run-change";
  }
  if (
    ranges.some(
      (range) => range.type === "unsupported" && range.reason === "source-formatting-mismatch",
    )
  ) {
    return "source-formatting-mismatch";
  }
  return "unrepresentable-paragraph-boundary";
};

const sourceSchedule = (source: ResolvedBlock): PreparedSourceSchedule =>
  Object.freeze({ phase: "source", from: source.from, to: source.to });

const textRangeSchedule = (
  source: ResolvedBlock,
  range: ReadyTextRange,
): PreparedSourceSchedule => {
  let from = Number.POSITIVE_INFINITY;
  let to = Number.NEGATIVE_INFINITY;
  for (const step of range.steps) {
    const stepFrom = step.type === "ins" ? step.at : step.from;
    const stepTo = step.type === "ins" ? step.at : step.to;
    from = Math.min(from, stepFrom);
    to = Math.max(to, stepTo);
  }
  // A compiled semantic change normally owns at least one concrete step. The
  // source span remains a safe deterministic schedule if a transport codec
  // represents a change without a positional mutation.
  return Number.isFinite(from) && Number.isFinite(to)
    ? Object.freeze({ phase: "source", from, to })
    : sourceSchedule(source);
};

const PREPARED_DOCX_COMPARISON_BRAND: unique symbol = Symbol("prepared-docx-comparison");

export type PreparedDocxComparison = {
  readonly [PREPARED_DOCX_COMPARISON_BRAND]: true;
  readonly issues: readonly DocxComparisonPreflightIssue[];
  readonly supportedInstructionCount: number;
  readonly totalInstructionCount: number;
};

type PreparedDocxComparisonState = {
  readonly state: EditorState;
  readonly instructions: readonly PreparedInstruction[];
  readonly totalInstructionCount: number;
  lifecycle: "ready" | "consumed";
};

const preparedDocxComparisons = new WeakMap<PreparedDocxComparison, PreparedDocxComparisonState>();

const ownPreparedDocxComparison = ({
  state,
  instructions,
  issues,
  totalInstructionCount,
}: {
  readonly state: EditorState;
  readonly instructions: readonly PreparedInstruction[];
  readonly issues: readonly DocxComparisonPreflightIssue[];
  readonly totalInstructionCount: number;
}): PreparedDocxComparison => {
  const ownedInstructions = Object.freeze([...instructions]);
  const ownedIssues = Object.freeze([...issues]);
  const dispositions = new Set([
    ...ownedInstructions.map(({ originalIndex }) => originalIndex),
    ...ownedIssues.map(({ instructionIndex }) => instructionIndex),
  ]);
  if (
    ownedInstructions.length + ownedIssues.length !== totalInstructionCount ||
    dispositions.size !== totalInstructionCount ||
    [...dispositions].some((index) => index < 0 || index >= totalInstructionCount)
  ) {
    return panic("DOCX comparison preflight did not disposition every instruction exactly once", {
      supported: ownedInstructions.length,
      unsupported: ownedIssues.length,
      totalInstructionCount,
    });
  }
  const prepared = Object.freeze({
    [PREPARED_DOCX_COMPARISON_BRAND]: true as const,
    issues: ownedIssues,
    supportedInstructionCount: ownedInstructions.length,
    totalInstructionCount,
  });
  preparedDocxComparisons.set(prepared, {
    state,
    instructions: ownedInstructions,
    totalInstructionCount,
    lifecycle: "ready",
  });
  return prepared;
};

/** Resolve every instruction against one immutable story before any transaction exists. */
export const preflightDocxComparisonProgram = ({
  state,
  snapshot,
  targetTables,
  program,
}: {
  readonly state: EditorState;
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly targetTables: ReadonlyMap<number, PMNode>;
  readonly program: DocxComparisonProgram;
}): PreparedDocxComparison => {
  const operationSnapshot = resolvedDocxOperationSnapshot(snapshot);
  const instructions = program.consume(snapshot);
  if (state.doc !== resolvedDocxSourceDocument(snapshot)) {
    return ownPreparedDocxComparison({
      state,
      instructions: [],
      issues: instructions.map((instruction, instructionIndex) =>
        issue(instruction, instructionIndex, "source-expectation-mismatch"),
      ),
      totalInstructionCount: instructions.length,
    });
  }
  const resolver = FolioStableBlockResolver.create(state.doc, operationSnapshot);
  const styleResolver = getDocumentStyleResolver(state);
  const prepared: PreparedInstruction[] = [];
  const issues: DocxComparisonPreflightIssue[] = [];
  for (const [instructionIndex, instruction] of instructions.entries()) {
    const resolveSource = (source: DocxComparisonSourceOperand) =>
      resolveExpectedBlock({ snapshot, resolver, source });
    const sourceBlockId = (source: DocxComparisonSourceOperand): string =>
      resolvedDocxSourceOperandBlock(source, snapshot).identity.id;
    switch (instruction.type) {
      case "replaceText":
      case "formatText": {
        const source = resolveSource(instruction.source);
        if (source.type === "unsupported") {
          issues.push(
            issue(instruction, instructionIndex, source.reason, sourceBlockId(instruction.source)),
          );
          break;
        }
        const range = preflightDocxTextRange(state.doc, {
          blockNode: source.block.node,
          blockFrom: source.block.from,
          sourceStartOffset: instruction.sourceStartOffset,
          range: instruction.range,
          styleResolver,
        });
        if (range.type === "unsupported") {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              range.reason === "pending-run-change"
                ? "pending-run-change"
                : range.reason === "source-formatting-mismatch"
                  ? "source-formatting-mismatch"
                  : "unrepresentable-text-range",
              sourceBlockId(instruction.source),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: instruction.type,
            source: source.block,
            sourceStartOffset: instruction.sourceStartOffset,
            range,
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: textRangeSchedule(source.block, range),
          }),
        );
        break;
      }
      case "insertParagraph": {
        const boundary = resolveParagraphBoundary(snapshot, resolver, instruction.boundary);
        if (!boundary) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "missing-anchor",
              sourceBlockId(instruction.boundary.paragraph),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "insertParagraph",
            boundary,
            target: instruction.target,
            originalIndex: instructionIndex,
            schedule: Object.freeze({ phase: "insertion", position: boundary.insertionPosition }),
          }),
        );
        break;
      }
      case "deleteParagraph": {
        const source = resolveSource(instruction.source);
        if (source.type === "unsupported") {
          issues.push(
            issue(instruction, instructionIndex, source.reason, sourceBlockId(instruction.source)),
          );
          break;
        }
        if (
          paragraphEndsItsContainer(
            state.doc.resolve(source.block.from),
            source.block.node.type.name,
          )
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "unrepresentable-paragraph-boundary",
              sourceBlockId(instruction.source),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: instruction.type,
            source: source.block,
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: sourceSchedule(source.block),
          }),
        );
        break;
      }
      case "deleteTrailingParagraphs": {
        const chainStart = resolveSource(instruction.chainStart);
        const deleted = instruction.deleted.map(resolveSource);
        const rejected =
          chainStart.type === "unsupported"
            ? { result: chainStart, blockId: sourceBlockId(instruction.chainStart) }
            : deleted
                .flatMap((result, index) =>
                  result.type === "unsupported"
                    ? [
                        {
                          result,
                          blockId:
                            instruction.deleted[index] === undefined
                              ? ""
                              : sourceBlockId(instruction.deleted[index]),
                        },
                      ]
                    : [],
                )
                .at(0);
        if (rejected !== undefined) {
          issues.push(
            issue(instruction, instructionIndex, rejected.result.reason, rejected.blockId),
          );
          break;
        }
        if (chainStart.type !== "ready") {
          return panic("A trailing deletion chain lost its resolved start");
        }
        const resolvedDeleted = deleted.map((result) => {
          if (result.type !== "ready") {
            return panic("A trailing deletion chain retained an unresolved member");
          }
          return result.block;
        });
        const last = resolvedDeleted.at(-1) ?? panic("A trailing deletion chain has no carrier");
        const sequence = [chainStart.block, ...resolvedDeleted];
        const isContiguous = sequence.every(
          (block, index) => index === 0 || sequence[index - 1]?.to === block.from,
        );
        if (
          !isContiguous ||
          !paragraphEndsItsContainer(state.doc.resolve(last.from), last.node.type.name) ||
          resolvedDeleted
            .slice(0, -1)
            .some((block) =>
              paragraphEndsItsContainer(state.doc.resolve(block.from), block.node.type.name),
            )
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "unrepresentable-paragraph-boundary",
              instruction.deleted.at(0) === undefined
                ? undefined
                : sourceBlockId(instruction.deleted[0]),
            ),
          );
          break;
        }
        const firstDeleted =
          resolvedDeleted.at(0) ?? panic("A trailing deletion chain lost its first member");
        prepared.push(
          Object.freeze({
            type: "deleteTrailingParagraphs",
            chainStart: chainStart.block,
            deleted: Object.freeze([firstDeleted, ...resolvedDeleted.slice(1)]),
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: Object.freeze({
              phase: "source",
              from: chainStart.block.from,
              to: last.to,
            }),
          }),
        );
        break;
      }
      case "moveParagraph": {
        const source = resolveSource(instruction.source);
        const successor = resolveSource(instruction.successor);
        const boundary = resolveParagraphBoundary(snapshot, resolver, instruction.boundary);
        const rejected =
          source.type === "unsupported"
            ? { result: source, blockId: sourceBlockId(instruction.source) }
            : successor.type === "unsupported"
              ? { result: successor, blockId: sourceBlockId(instruction.successor) }
              : null;
        if (rejected || !boundary) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              rejected?.result.reason ?? "missing-anchor",
              rejected?.blockId ?? sourceBlockId(instruction.boundary.paragraph),
            ),
          );
          break;
        }
        if (source.type !== "ready" || successor.type !== "ready") {
          return panic("A paragraph move retained an unresolved source boundary");
        }
        if (
          source.block.to !== successor.block.from ||
          paragraphEndsItsContainer(
            state.doc.resolve(source.block.from),
            source.block.node.type.name,
          )
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "unrepresentable-paragraph-boundary",
              sourceBlockId(instruction.source),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "moveParagraph",
            source: source.block,
            successor: successor.block,
            boundary,
            target: instruction.target,
            originalIndex: instructionIndex,
            sourceSchedule: sourceSchedule(source.block),
            destinationSchedule: Object.freeze({
              phase: "insertion",
              position: boundary.insertionPosition,
            }),
          }),
        );
        break;
      }
      case "moveTerminalParagraph": {
        const predecessor = resolveSource(instruction.predecessor);
        const source = resolveSource(instruction.source);
        const boundary = resolveParagraphBoundary(snapshot, resolver, instruction.boundary);
        const rejected =
          predecessor.type === "unsupported"
            ? { result: predecessor, blockId: sourceBlockId(instruction.predecessor) }
            : source.type === "unsupported"
              ? { result: source, blockId: sourceBlockId(instruction.source) }
              : null;
        if (rejected || !boundary) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              rejected?.result.reason ?? "missing-anchor",
              rejected?.blockId ?? sourceBlockId(instruction.boundary.paragraph),
            ),
          );
          break;
        }
        if (predecessor.type !== "ready" || source.type !== "ready") {
          return panic("A terminal paragraph move retained an unresolved source boundary");
        }
        if (
          predecessor.block.to !== source.block.from ||
          !canJoin(state.doc, predecessor.block.to) ||
          paragraphEndsItsContainer(
            state.doc.resolve(predecessor.block.from),
            predecessor.block.node.type.name,
          ) ||
          !paragraphEndsItsContainer(
            state.doc.resolve(source.block.from),
            source.block.node.type.name,
          )
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "unrepresentable-paragraph-boundary",
              sourceBlockId(instruction.source),
            ),
          );
          break;
        }
        const predecessorAttrs = expectParagraphAttrs(predecessor.block.node);
        const sourceAttrs = expectParagraphAttrs(source.block.node);
        if (
          predecessorAttrs.pPrMark != null ||
          sourceAttrs.pPrMark != null ||
          hasSerializableParagraphPropertyChange(sourceAttrs._propertyChanges)
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "pending-paragraph-change",
              sourceBlockId(instruction.source),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "moveTerminalParagraph",
            predecessor: predecessor.block,
            source: source.block,
            boundary,
            target: instruction.target,
            semantic: instruction,
            originalIndex: instructionIndex,
            sourceSchedule: Object.freeze({
              phase: "source",
              from: predecessor.block.from,
              to: source.block.to,
            }),
            destinationSchedule: Object.freeze({
              phase: "insertion",
              position: boundary.insertionPosition,
            }),
          }),
        );
        break;
      }
      case "splitParagraph": {
        const source = resolveSource(instruction.source);
        if (source.type === "unsupported") {
          issues.push(
            issue(instruction, instructionIndex, source.reason, sourceBlockId(instruction.source)),
          );
          break;
        }
        const clean = buildCleanBlockText(source.block.node, source.block.from);
        const splitPosition = clean.offsets[instruction.offset];
        const first = preflightDocxTextRange(state.doc, {
          blockNode: source.block.node,
          blockFrom: source.block.from,
          sourceStartOffset: 0,
          range: instruction.first,
          styleResolver,
        });
        const second = preflightDocxTextRange(state.doc, {
          blockNode: source.block.node,
          blockFrom: source.block.from,
          sourceStartOffset: instruction.offset + instruction.separatorText.length,
          range: instruction.second,
          styleResolver,
        });
        if (
          splitPosition === undefined ||
          !canSplit(state.doc, splitPosition) ||
          first.type === "unsupported" ||
          second.type === "unsupported"
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              structuralRangeIssueReason([first, second]),
              sourceBlockId(instruction.source),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "splitParagraph",
            source: source.block,
            splitPosition,
            first,
            second,
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: sourceSchedule(source.block),
          }),
        );
        break;
      }
      case "mergeParagraphs": {
        const firstSource = resolveSource(instruction.firstSource);
        const secondSource = resolveSource(instruction.secondSource);
        if (firstSource.type === "unsupported" || secondSource.type === "unsupported") {
          const rejected = firstSource.type === "unsupported" ? firstSource : secondSource;
          const blockId =
            firstSource.type === "unsupported"
              ? sourceBlockId(instruction.firstSource)
              : sourceBlockId(instruction.secondSource);
          issues.push(issue(instruction, instructionIndex, rejected.reason, blockId));
          break;
        }
        const first = preflightDocxTextRange(state.doc, {
          blockNode: firstSource.block.node,
          blockFrom: firstSource.block.from,
          sourceStartOffset: 0,
          range: instruction.first,
          styleResolver,
        });
        const second = preflightDocxTextRange(state.doc, {
          blockNode: secondSource.block.node,
          blockFrom: secondSource.block.from,
          sourceStartOffset: 0,
          range: instruction.second,
          styleResolver,
        });
        if (
          firstSource.block.to !== secondSource.block.from ||
          !canJoin(state.doc, firstSource.block.to) ||
          first.type === "unsupported" ||
          second.type === "unsupported"
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              structuralRangeIssueReason([first, second]),
              sourceBlockId(instruction.firstSource),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "mergeParagraphs",
            firstSource: firstSource.block,
            secondSource: secondSource.block,
            first,
            second,
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: Object.freeze({
              phase: "source",
              from: firstSource.block.from,
              to: secondSource.block.to,
            }),
          }),
        );
        break;
      }
      case "setParagraphProperties": {
        const source = resolveSource(instruction.source);
        if (source.type === "unsupported") {
          issues.push(
            issue(instruction, instructionIndex, source.reason, sourceBlockId(instruction.source)),
          );
          break;
        }
        if (
          hasSerializableParagraphPropertyChange(
            expectParagraphAttrs(source.block.node)._propertyChanges,
          )
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "pending-paragraph-change",
              sourceBlockId(instruction.source),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "setParagraphProperties",
            source: source.block,
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: sourceSchedule(source.block),
          }),
        );
        break;
      }
      case "insertTable":
      case "deleteTable":
      case "replaceTable":
      case "insertTableRow":
      case "deleteTableRow":
      case "insertTableColumn":
      case "deleteTableColumn":
        issues.push(issue(instruction, instructionIndex, "unresolved-table-instruction"));
        break;
      case "matchTableGeometry": {
        const geometry = preflightTableGeometry({
          baseTables: storyTablesOf(operationSnapshot),
          targetTables,
          pairings: instruction.pairings,
        });
        if (geometry.status === "unsupported") {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "unrepresentable-table-geometry",
              undefined,
              geometry.issue,
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "matchTableGeometry",
            program: geometry.program,
            originalIndex: instructionIndex,
            schedule: Object.freeze({ phase: "geometry" }),
          }),
        );
        break;
      }
      default: {
        const unreachable: never = instruction;
        return panic("Unhandled DOCX comparison instruction during preflight", {
          instruction: unreachable,
        });
      }
    }
  }
  return ownPreparedDocxComparison({
    state,
    instructions: prepared,
    issues,
    totalInstructionCount: instructions.length,
  });
};

const trackedParagraph = (
  tr: Transaction,
  boundary: ResolvedParagraphBoundary,
  target: DocxComparisonParagraphTarget,
  textRevisionId: number | null,
  paragraphRevisionId: number,
  author: string,
  date: string,
  kind: "ins" | "moveTo" = "ins",
): PMNode => {
  const insertionType =
    tr.doc.type.schema.marks["insertion"] ??
    panic("A preflighted DOCX comparison lost tracked-text schema support");
  const marks =
    textRevisionId === null
      ? []
      : [
          insertionType.create({
            revisionId: textRevisionId,
            author,
            date,
            ...(kind === "moveTo" && { moveKind: "moveTo" }),
          }),
        ];
  const content = target.text.length === 0 ? null : tr.doc.type.schema.text(target.text, marks);
  return boundary.node.type.create(
    {
      ...stripBlockIdentityAttrs(boundary.node.attrs),
      pPrMark: { kind, info: { id: paragraphRevisionId, author, date } },
      _propertyChanges: null,
    },
    content,
  );
};

const markParagraphTextDeletion = ({
  tr,
  source,
  revisionId,
  author,
  date,
  kind = "del",
}: {
  readonly tr: Transaction;
  readonly source: ResolvedBlock;
  readonly revisionId: number;
  readonly author: string;
  readonly date: string;
  readonly kind?: "del" | "moveFrom";
}): {
  readonly nextRevisionId: number;
  readonly revisionIds: readonly number[];
  readonly transaction: Transaction;
} => {
  const at = source.from;
  const node = tr.doc.nodeAt(at) ?? panic("A preflighted paragraph deletion lost its source");
  const clean = buildCleanBlockText(node, at);
  const deletionType =
    tr.doc.type.schema.marks["deletion"] ??
    panic("A preflighted DOCX comparison lost tracked-text schema support");
  let nextRevisionId = revisionId;
  const revisionIds: number[] = [];
  const from = clean.offsets[0];
  const to = clean.offsets[clean.text.length];
  if (from !== undefined && to !== undefined && from < to) {
    const textRevisionId = nextRevisionId++;
    tr.addMark(
      from,
      to,
      deletionType.create({
        revisionId: textRevisionId,
        author,
        date,
        ...(kind === "moveFrom" && { moveKind: "moveFrom" }),
      }),
    );
    revisionIds.push(textRevisionId);
  }
  return { transaction: tr, nextRevisionId, revisionIds: Object.freeze(revisionIds) };
};

const markParagraphDeletion = ({
  tr,
  source,
  revisionId,
  author,
  date,
  kind = "del",
}: {
  readonly tr: Transaction;
  readonly source: ResolvedBlock;
  readonly revisionId: number;
  readonly author: string;
  readonly date: string;
  readonly kind?: "del" | "moveFrom";
}): {
  readonly nextRevisionId: number;
  readonly revisionIds: readonly number[];
  readonly transaction: Transaction;
} => {
  const text = markParagraphTextDeletion({
    tr,
    source,
    revisionId,
    author,
    date,
    kind,
  });
  const node = text.transaction.doc.nodeAt(source.from) ??
    panic("A preflighted paragraph deletion lost its source");
  let nextRevisionId = text.nextRevisionId;
  const revisionIds = [...text.revisionIds];
  if (!paragraphEndsItsContainer(text.transaction.doc.resolve(source.from), node.type.name)) {
    const paragraphRevisionId = nextRevisionId++;
    text.transaction.setNodeAttribute(source.from, "pPrMark", {
      kind,
      info: { id: paragraphRevisionId, author, date },
    });
    revisionIds.push(paragraphRevisionId);
  }
  return {
    transaction: text.transaction,
    nextRevisionId,
    revisionIds: Object.freeze(revisionIds),
  };
};

const applyParagraphTargetProperties = ({
  tr,
  position,
  target,
  revisionId,
  author,
  date,
  styleResolver,
  numbering,
  tracked,
}: {
  readonly tr: Transaction;
  readonly position: number;
  readonly target: DocxComparisonParagraphTarget;
  readonly revisionId: number;
  readonly author: string;
  readonly date: string;
  readonly styleResolver: ReturnType<typeof getDocumentStyleResolver>;
  readonly numbering: ReturnType<typeof getDocumentNumbering>;
  readonly tracked: boolean;
}): {
  readonly transaction: Transaction;
  readonly nextRevisionId: number;
  readonly revisionIds: readonly number[];
} => {
  const node = tr.doc.nodeAt(position) ?? panic("A paragraph-property target was not produced");
  let nextRevisionId = revisionId;
  const revisionIds: number[] = [];
  const applied = applyBlockParagraphProperties({
    tr,
    position,
    node,
    properties: target.properties,
    styleResolver,
    numbering,
    ...(tracked
      ? {
          revisionInfo: () => {
            const id = nextRevisionId++;
            revisionIds.push(id);
            return { id, author, date };
          },
        }
      : {}),
  });
  return {
    transaction: applied.tr,
    nextRevisionId,
    revisionIds: Object.freeze(revisionIds),
  };
};

const applyLiveRange = ({
  tr,
  blockPosition,
  sourceStartOffset,
  semantic,
  application,
  revisionId,
  author,
  date,
  styleResolver,
}: {
  readonly tr: Transaction;
  readonly blockPosition: number;
  readonly sourceStartOffset: number;
  readonly semantic: Extract<
    DocxComparisonInstruction,
    { readonly type: "splitParagraph" | "mergeParagraphs" }
  >["first"];
  readonly application: "format" | "replace";
  readonly revisionId: number;
  readonly author: string;
  readonly date: string;
  readonly styleResolver: ReturnType<typeof getDocumentStyleResolver>;
}): {
  readonly transaction: Transaction;
  readonly nextRevisionId: number;
  readonly revisionIds: readonly number[];
} => {
  const node = tr.doc.nodeAt(blockPosition) ?? panic("A structural text target was not produced");
  const range = preflightDocxTextRange(tr.doc, {
    blockNode: node,
    blockFrom: blockPosition,
    sourceStartOffset,
    range: semantic,
    styleResolver,
  });
  if (range.type === "unsupported") {
    return panic("A preflighted structural text range became unrepresentable", {
      reason: range.reason,
    });
  }
  const applied = applyPreflightedDocxTextRange({
    tr,
    preflight: range,
    revisionIdSeed: revisionId,
    author,
    date,
    styleResolver,
    application,
  });
  return {
    transaction: applied.transaction,
    nextRevisionId: applied.nextRevisionId,
    revisionIds: applied.revisionIds,
  };
};

type AppliedInsertionMember = {
  readonly instruction: PreparedInsertionInstruction | PreparedMoveInstruction;
  readonly revisionIds: readonly number[];
};

const applyInsertionRun = ({
  tr,
  run,
  revisionId,
  author,
  date,
  styleResolver,
  numbering,
}: {
  readonly tr: Transaction;
  readonly run: PreparedInsertionRun;
  readonly revisionId: number;
  readonly author: string;
  readonly date: string;
  readonly styleResolver: ReturnType<typeof getDocumentStyleResolver>;
  readonly numbering: ReturnType<typeof getDocumentNumbering>;
}): {
  readonly transaction: Transaction;
  readonly nextRevisionId: number;
  readonly members: readonly AppliedInsertionMember[];
} => {
  let nextRevisionId = revisionId;
  const prepared: {
    readonly instruction: PreparedInsertionInstruction | PreparedMoveInstruction;
    readonly paragraph: PMNode;
    readonly revisionIds: readonly number[];
  }[] = [];
  for (const { instruction } of run.members) {
    const textRevisionId = instruction.target.text.length > 0 ? nextRevisionId++ : null;
    const paragraphRevisionId = nextRevisionId++;
    const revisionIds = [...(textRevisionId === null ? [] : [textRevisionId]), paragraphRevisionId];
    prepared.push({
      instruction,
      paragraph: trackedParagraph(
        tr,
        instruction.boundary,
        instruction.target,
        textRevisionId,
        paragraphRevisionId,
        author,
        date,
        instruction.type === "insertParagraph" ? "ins" : "moveTo",
      ),
      revisionIds: Object.freeze(revisionIds),
    });
  }
  tr.insert(
    run.position,
    prepared.map(({ paragraph }) => paragraph),
  );
  let position = run.position;
  const applied: AppliedInsertionMember[] = [];
  for (const member of prepared) {
    const inserted =
      tr.doc.nodeAt(position) ??
      panic("A preflighted insertion run did not produce its target paragraph");
    tr = applyBlockParagraphProperties({
      tr,
      position,
      node: inserted,
      properties: member.instruction.target.properties,
      styleResolver,
      numbering,
    }).tr;
    for (const authoredRun of member.instruction.target.runs) {
      applyExactDirectFormatting({
        tr,
        from: position + 1 + authoredRun.startOffset,
        to: position + 1 + authoredRun.endOffset,
        formatting: authoredRun.formatting,
        styleResolver,
      });
    }
    applied.push(
      Object.freeze({
        instruction: member.instruction,
        revisionIds: member.revisionIds,
      }),
    );
    position += member.paragraph.nodeSize;
  }
  return Object.freeze({
    transaction: tr,
    nextRevisionId,
    members: Object.freeze(applied),
  });
};

export type DocxComparisonInstructionReceipt = {
  readonly instructionIndex: number;
  readonly instructionType: DocxComparisonInstruction["type"];
  readonly revisionIds: readonly number[];
  readonly tableGeometry?: TableGeometryExecutionReceipt;
};

export type DocxComparisonExecutionReceipt = {
  readonly instructions: readonly DocxComparisonInstructionReceipt[];
  readonly nextRevisionId: number;
  /** Positional tasks executed after one global right-to-left schedule. */
  readonly executionTaskCount: number;
  /** Same-boundary paragraph runs inserted atomically. */
  readonly insertionRunCount: number;
  /** Step maps traversed only inside one compound instruction. */
  readonly localPositionMappingSteps: number;
  readonly transaction: Transaction;
};

export type DocxComparisonExecutionIssue =
  | { readonly reason: "stale-preflight" }
  | { readonly reason: "invalid-revision-stamp" }
  | {
      readonly reason: "table-geometry-execution";
      readonly instructionIndex: number;
      readonly issue: TableGeometryExecutionIssue;
    };

export type DocxComparisonExecutionResult =
  | { readonly status: "executed"; readonly receipt: DocxComparisonExecutionReceipt }
  | { readonly status: "unsupported"; readonly issue: DocxComparisonExecutionIssue };

const unsupportedExecution = (issue: DocxComparisonExecutionIssue): DocxComparisonExecutionResult =>
  Object.freeze({ status: "unsupported", issue: Object.freeze(issue) });

const consumePreparedDocxComparison = (
  prepared: PreparedDocxComparison,
  state: EditorState,
): PreparedDocxComparisonState | null => {
  const owned =
    preparedDocxComparisons.get(prepared) ??
    panic("A prepared DOCX comparison must come from preflightDocxComparisonProgram");
  if (owned.lifecycle !== "ready") {
    return panic("A prepared DOCX comparison was consumed more than once");
  }
  owned.lifecycle = "consumed";
  return owned.state === state ? owned : null;
};

/** Execute only preflighted instructions into one transaction; the caller dispatches once. */
export const executePreflightedDocxComparison = ({
  state,
  prepared,
  revisionStamp,
  author,
}: {
  readonly state: EditorState;
  readonly prepared: PreparedDocxComparison;
  readonly revisionStamp: FolioRevisionStamp;
  readonly author: string;
}): DocxComparisonExecutionResult => {
  const owned = consumePreparedDocxComparison(prepared, state);
  if (!owned) return unsupportedExecution({ reason: "stale-preflight" });
  if (
    typeof author !== "string" ||
    typeof revisionStamp.date !== "string" ||
    author.length + revisionStamp.date.length > 65_536 ||
    !Number.isSafeInteger(revisionStamp.idSeed) ||
    revisionStamp.idSeed < 0
  ) {
    return unsupportedExecution({ reason: "invalid-revision-stamp" });
  }
  let tr = state.tr;
  let revisionId = revisionStamp.idSeed;
  const styleResolver = getDocumentStyleResolver(state);
  const numbering = getDocumentNumbering(state);
  type ReceiptParts = {
    readonly instructionType: DocxComparisonInstruction["type"];
    readonly parts: Map<number, readonly number[]>;
    tableGeometry?: TableGeometryExecutionReceipt;
  };
  const receiptPartsByIndex = new Map<number, ReceiptParts>();
  const recordPart = ({
    instructionIndex,
    instructionType,
    part,
    revisionIds,
    tableGeometry,
  }: {
    readonly instructionIndex: number;
    readonly instructionType: DocxComparisonInstruction["type"];
    readonly part: number;
    readonly revisionIds: readonly number[];
    readonly tableGeometry?: TableGeometryExecutionReceipt;
  }): void => {
    const receipt = receiptPartsByIndex.get(instructionIndex) ?? {
      instructionType,
      parts: new Map<number, readonly number[]>(),
    };
    if (receipt.instructionType !== instructionType || receipt.parts.has(part)) {
      return panic("A DOCX comparison instruction produced a duplicate execution part", {
        instructionIndex,
        instructionType,
        part,
      });
    }
    receipt.parts.set(part, Object.freeze([...revisionIds]));
    if (tableGeometry !== undefined) receipt.tableGeometry = tableGeometry;
    receiptPartsByIndex.set(instructionIndex, receipt);
  };
  const geometry = owned.instructions.filter(
    (instruction): instruction is PreparedGeometryInstruction =>
      instruction.type === "matchTableGeometry",
  );
  for (const instruction of geometry) {
    const result = executeTableGeometryProgram({
      tr,
      program: instruction.program,
      revision: { author, date: revisionStamp.date, idSeed: revisionId },
    });
    if (result.status === "unsupported") {
      return unsupportedExecution({
        reason: "table-geometry-execution",
        instructionIndex: instruction.originalIndex,
        issue: result.issue,
      });
    }
    revisionId = result.receipt.nextRevisionId;
    recordPart({
      instructionIndex: instruction.originalIndex,
      instructionType: instruction.type,
      part: 0,
      revisionIds: result.receipt.revisions.map(
        ({ revisionId: appliedRevisionId }) => appliedRevisionId,
      ),
      tableGeometry: result.receipt,
    });
  }
  const ordered = executionTasks(owned.instructions);
  let insertionRunCount = 0;
  let localPositionMappingSteps = 0;
  for (const task of ordered) {
    const stepsBefore = tr.steps.length;
    if (task.type === "insertionRun") {
      insertionRunCount++;
      const inserted = applyInsertionRun({
        tr,
        run: task,
        revisionId,
        author,
        date: revisionStamp.date,
        styleResolver,
        numbering,
      });
      tr = inserted.transaction;
      revisionId = inserted.nextRevisionId;
      for (const member of inserted.members) {
        recordPart({
          instructionIndex: member.instruction.originalIndex,
          instructionType: member.instruction.type,
          part: 0,
          revisionIds: member.revisionIds,
        });
      }
      if (tr.steps.length === stepsBefore) {
        return panic("A preflighted DOCX insertion run produced no transaction step");
      }
      continue;
    }
    const instruction = task.instruction;
    const instructionRevisionIds: number[] = [];
    if (task.type === "moveSource") {
      if (instruction.type === "moveParagraph") {
        const deleted = markParagraphDeletion({
          tr,
          source: instruction.source,
          revisionId,
          author,
          date: revisionStamp.date,
          kind: "moveFrom",
        });
        tr = deleted.transaction;
        revisionId = deleted.nextRevisionId;
        instructionRevisionIds.push(...deleted.revisionIds);
      } else {
        const deletedText = markParagraphTextDeletion({
          tr,
          source: instruction.source,
          revisionId,
          author,
          date: revisionStamp.date,
          kind: "moveFrom",
        });
        tr = deletedText.transaction;
        revisionId = deletedText.nextRevisionId;
        instructionRevisionIds.push(...deletedText.revisionIds);
        const paragraphRevisionId = revisionId++;
        tr.setNodeAttribute(instruction.predecessor.from, "pPrMark", {
          kind: "moveFrom",
          info: { id: paragraphRevisionId, author, date: revisionStamp.date },
        });
        instructionRevisionIds.push(paragraphRevisionId);
        const carrier = tr.doc.nodeAt(instruction.source.from) ??
          panic("A preflighted terminal move lost its paragraph-mark carrier");
        const propertyRevisionIds: number[] = [];
        tr = applyBlockParagraphProperties({
          tr,
          position: instruction.source.from,
          node: carrier,
          properties: instruction.semantic.carrierTargetProperties,
          styleResolver,
          numbering,
          revisionInfo: () => {
            const id = revisionId++;
            propertyRevisionIds.push(id);
            return { id, author, date: revisionStamp.date };
          },
        }).tr;
        instructionRevisionIds.push(...propertyRevisionIds);
      }
      if (tr.steps.length === stepsBefore) {
        return panic("A preflighted DOCX move source produced no transaction step");
      }
      recordPart({
        instructionIndex: instruction.originalIndex,
        instructionType: instruction.type,
        part: 1,
        revisionIds: instructionRevisionIds,
      });
      continue;
    }
    const localMappingStart = tr.mapping.maps.length;
    switch (instruction.type) {
      case "replaceText":
      case "formatText": {
        const result = applyPreflightedDocxTextRange({
          tr,
          preflight: instruction.range,
          revisionIdSeed: revisionId,
          author,
          date: revisionStamp.date,
          styleResolver,
          application: instruction.type === "replaceText" ? "replace" : "format",
        });
        tr = result.transaction;
        revisionId = result.nextRevisionId;
        instructionRevisionIds.push(...result.revisionIds);
        break;
      }
      case "deleteParagraph": {
        const deleted = markParagraphDeletion({
          tr,
          source: instruction.source,
          revisionId,
          author,
          date: revisionStamp.date,
        });
        tr = deleted.transaction;
        revisionId = deleted.nextRevisionId;
        instructionRevisionIds.push(...deleted.revisionIds);
        break;
      }
      case "setParagraphProperties": {
        const at = instruction.source.from;
        const node = tr.doc.nodeAt(at) ?? panic("A preflighted paragraph change lost its source");
        const appliedRevisionIds: number[] = [];
        const appliedProperties = applyBlockParagraphProperties({
          tr,
          position: at,
          node,
          properties: instruction.semantic.targetProperties,
          styleResolver,
          numbering,
          revisionInfo: () => {
            const id = revisionId++;
            appliedRevisionIds.push(id);
            return { id, author, date: revisionStamp.date };
          },
        });
        tr = appliedProperties.tr;
        instructionRevisionIds.push(...appliedRevisionIds);
        break;
      }
      case "splitParagraph": {
        const sourcePosition = instruction.source.from;
        const live =
          tr.doc.nodeAt(sourcePosition) ?? panic("A preflighted paragraph split lost its source");
        const clean = buildCleanBlockText(live, sourcePosition);
        const splitAt =
          clean.offsets[instruction.semantic.offset] ??
          panic("A preflighted paragraph split lost its UTF-16 boundary");
        const separatorEnd =
          clean.offsets[instruction.semantic.offset + instruction.semantic.separatorText.length] ??
          panic("A preflighted paragraph split lost its separator boundary");
        if (separatorEnd > splitAt) {
          const deletionType =
            tr.doc.type.schema.marks["deletion"] ??
            panic("A preflighted DOCX comparison lost tracked-text schema support");
          const separatorRevisionId = revisionId++;
          tr.addMark(
            splitAt,
            separatorEnd,
            deletionType.create({
              revisionId: separatorRevisionId,
              author,
              date: revisionStamp.date,
            }),
          );
          instructionRevisionIds.push(separatorRevisionId);
        }
        tr.split(splitAt);
        const paragraphRevisionId = revisionId++;
        tr.setNodeAttribute(sourcePosition, "pPrMark", {
          kind: "ins",
          info: { id: paragraphRevisionId, author, date: revisionStamp.date },
        });
        instructionRevisionIds.push(paragraphRevisionId);
        const secondPosition = splitAt + 1;
        for (const target of [
          {
            position: secondPosition,
            range: instruction.semantic.second,
            paragraph: instruction.semantic.secondTarget,
          },
          {
            position: sourcePosition,
            range: instruction.semantic.first,
            paragraph: instruction.semantic.firstTarget,
          },
        ]) {
          const appliedRange = applyLiveRange({
            tr,
            blockPosition: target.position,
            sourceStartOffset: 0,
            semantic: target.range,
            application: "replace",
            revisionId,
            author,
            date: revisionStamp.date,
            styleResolver,
          });
          tr = appliedRange.transaction;
          revisionId = appliedRange.nextRevisionId;
          instructionRevisionIds.push(...appliedRange.revisionIds);
          const properties = applyParagraphTargetProperties({
            tr,
            position: target.position,
            target: target.paragraph,
            revisionId,
            author,
            date: revisionStamp.date,
            styleResolver,
            numbering,
            tracked: true,
          });
          tr = properties.transaction;
          revisionId = properties.nextRevisionId;
          instructionRevisionIds.push(...properties.revisionIds);
        }
        break;
      }
      case "mergeParagraphs": {
        const firstPosition = instruction.firstSource.from;
        const firstNode =
          tr.doc.nodeAt(firstPosition) ??
          panic("A preflighted paragraph merge lost its first source");
        const insertAt = firstPosition + firstNode.nodeSize - 1;
        if (instruction.semantic.separatorText.length > 0) {
          const insertionType =
            tr.doc.type.schema.marks["insertion"] ??
            panic("A preflighted DOCX comparison lost tracked-text schema support");
          const separatorRevisionId = revisionId++;
          tr.insertText(instruction.semantic.separatorText, insertAt);
          tr.addMark(
            insertAt,
            insertAt + instruction.semantic.separatorText.length,
            insertionType.create({
              revisionId: separatorRevisionId,
              author,
              date: revisionStamp.date,
            }),
          );
          for (const run of instruction.semantic.separatorRuns) {
            applyExactDirectFormatting({
              tr,
              from: insertAt + run.startOffset,
              to: insertAt + run.endOffset,
              formatting: run.formatting,
              styleResolver,
            });
          }
          instructionRevisionIds.push(separatorRevisionId);
        }
        const paragraphRevisionId = revisionId++;
        tr.setNodeAttribute(firstPosition, "pPrMark", {
          kind: "del",
          info: { id: paragraphRevisionId, author, date: revisionStamp.date },
        });
        instructionRevisionIds.push(paragraphRevisionId);
        const localMapping = tr.mapping.slice(localMappingStart);
        localPositionMappingSteps += localMapping.maps.length;
        const secondPosition = localMapping.map(instruction.secondSource.from);
        for (const target of [
          { position: secondPosition, range: instruction.semantic.second },
          { position: firstPosition, range: instruction.semantic.first },
        ]) {
          const appliedRange = applyLiveRange({
            tr,
            blockPosition: target.position,
            sourceStartOffset: 0,
            semantic: target.range,
            application: "replace",
            revisionId,
            author,
            date: revisionStamp.date,
            styleResolver,
          });
          tr = appliedRange.transaction;
          revisionId = appliedRange.nextRevisionId;
          instructionRevisionIds.push(...appliedRange.revisionIds);
        }
        const properties = applyParagraphTargetProperties({
          tr,
          position: firstPosition,
          target: instruction.semantic.target,
          revisionId,
          author,
          date: revisionStamp.date,
          styleResolver,
          numbering,
          tracked: true,
        });
        tr = properties.transaction;
        revisionId = properties.nextRevisionId;
        instructionRevisionIds.push(...properties.revisionIds);
        break;
      }
      case "deleteTrailingParagraphs": {
        for (const source of instruction.deleted.toReversed()) {
          const deleted = markParagraphDeletion({
            tr,
            source,
            revisionId,
            author,
            date: revisionStamp.date,
          });
          tr = deleted.transaction;
          revisionId = deleted.nextRevisionId;
          instructionRevisionIds.push(...deleted.revisionIds);
        }
        const paragraphRevisionId = revisionId++;
        tr.setNodeAttribute(instruction.chainStart.from, "pPrMark", {
          kind: "del",
          info: { id: paragraphRevisionId, author, date: revisionStamp.date },
        });
        instructionRevisionIds.push(paragraphRevisionId);
        break;
      }
      default: {
        const unreachable: never = instruction;
        return panic("Unhandled preflighted DOCX comparison instruction", {
          instruction: unreachable,
        });
      }
    }
    if (tr.steps.length === stepsBefore) {
      return panic("A preflighted DOCX comparison instruction produced no transaction step", {
        type: instruction.type,
      });
    }
    recordPart({
      instructionIndex: instruction.originalIndex,
      instructionType: instruction.type,
      part: 0,
      revisionIds: instructionRevisionIds,
    });
  }
  const synthesizedRevisionIdsByIndex = new Map<number, number[]>();
  if (tr.docChanged) {
    const revisionOwnerById = new Map<number, number>();
    for (const [instructionIndex, receipt] of receiptPartsByIndex) {
      for (const revisionIds of receipt.parts.values()) {
        for (const appliedRevisionId of revisionIds) {
          revisionOwnerById.set(appliedRevisionId, instructionIndex);
        }
      }
    }
    const rotated = withRotatedAddedFinalBreaks({
      tr,
      batchRevisionIds: new Set(revisionOwnerById.keys()),
      revisionSeed: revisionId,
      author,
      date: revisionStamp.date,
      initials: undefined,
    });
    tr = rotated.transaction;
    revisionId = rotated.nextRevisionId;
    for (const {
      ownerRevisionId,
      revisionId: synthesizedRevisionId,
    } of rotated.synthesizedRevisions) {
      const instructionIndex =
        revisionOwnerById.get(ownerRevisionId) ??
        panic("A final-mark normalization lost its comparison instruction receipt", {
          ownerRevisionId,
        });
      const synthesized = synthesizedRevisionIdsByIndex.get(instructionIndex) ?? [];
      synthesized.push(synthesizedRevisionId);
      synthesizedRevisionIdsByIndex.set(instructionIndex, synthesized);
    }
    markStructuralChange(tr);
    requestDeterministicParaIds(tr, `${revisionStamp.date}:${String(revisionStamp.idSeed)}`);
  }
  const receipts = owned.instructions
    .toSorted((left, right) => left.originalIndex - right.originalIndex)
    .map((instruction): DocxComparisonInstructionReceipt => {
      const recorded =
        receiptPartsByIndex.get(instruction.originalIndex) ??
        panic("A DOCX comparison instruction completed without an execution receipt", {
          instructionIndex: instruction.originalIndex,
        });
      const expectedPartCount =
        instruction.type === "moveParagraph" || instruction.type === "moveTerminalParagraph"
          ? 2
          : 1;
      if (
        recorded.parts.size !== expectedPartCount ||
        Array.from({ length: expectedPartCount }, (_, part) => part).some(
          (part) => !recorded.parts.has(part),
        )
      ) {
        return panic("A DOCX comparison instruction did not execute every required part", {
          instructionIndex: instruction.originalIndex,
          expectedPartCount,
          actualPartCount: recorded.parts.size,
        });
      }
      const revisionIds = Array.from(
        { length: expectedPartCount },
        (_, part) => recorded.parts.get(part) ?? [],
      ).flat();
      revisionIds.push(...(synthesizedRevisionIdsByIndex.get(instruction.originalIndex) ?? []));
      return Object.freeze({
        instructionIndex: instruction.originalIndex,
        instructionType: recorded.instructionType,
        revisionIds: Object.freeze(revisionIds),
        ...(recorded.tableGeometry !== undefined && {
          tableGeometry: recorded.tableGeometry,
        }),
      });
    });
  if (receipts.length !== owned.instructions.length) {
    return panic("A DOCX comparison execution receipt lost an instruction", {
      expected: owned.instructions.length,
      actual: receipts.length,
    });
  }
  return Object.freeze({
    status: "executed" as const,
    receipt: Object.freeze({
      instructions: Object.freeze(receipts),
      nextRevisionId: revisionId,
      executionTaskCount: geometry.length + ordered.length,
      insertionRunCount,
      localPositionMappingSteps,
      transaction: tr,
    }),
  });
};
