import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { canJoin, canSplit } from "prosemirror-transform";

import { applyBlockParagraphProperties, withRotatedAddedFinalBreaks } from "../../ai-edits/apply";
import { buildCleanBlockText } from "../../ai-edits/clean-text";
import { FolioStableBlockResolver } from "./stable-block-resolution";
import { expectParagraphAttrs } from "../../prosemirror/attrs";
import { hasSerializableParagraphPropertyChange } from "../../prosemirror/commands/propertyChangeScope";
import { paragraphEndsItsContainer } from "../../prosemirror/containerFinalParagraph";
import { markStructuralChange } from "../../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { requestDeterministicParaIds } from "../../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { getDocumentNumbering } from "../../prosemirror/plugins/documentNumbering";
import { getDocumentStyleResolver } from "../../prosemirror/plugins/documentStyles";
import {
  COMPARE_DOCX_PREFLIGHT_REASONS,
  type CompareChange,
  type CompareDocxPreflightReason,
} from "../../compare/types";
import { stripBlockIdentityAttrs } from "../../ai-edits/block-identity";
import type { FolioRevisionStamp } from "../../ai-edits/apply";
import {
  DocxComparisonProgram,
  type DocxComparisonInstruction,
  type DocxComparisonParagraphTarget,
  type DocxComparisonSemanticGroup,
  type DocxComparisonSourceOperand,
} from "./docx-program";
import {
  applyExactDirectFormatting,
  applyExactInlineOwnership,
  applyPreflightedDocxTextRange,
  preflightDocxTextRange,
  type DocxTextRangePreflight,
} from "./docx-text-executor";
import {
  type TableGeometryExecutionIssue,
  type TableGeometryExecutionReceipt,
  type TableGeometryUnsupportedIssue,
} from "./table-geometry-program";
import {
  executeTableStructureGeometry,
  executeTableStructureTask,
  preflightTableStructureComponents,
  type PreparedTableStructureProgram,
  type PreparedTableStructureTask,
  type TableStructureUnsupportedIssue,
} from "./table-structure-program";
import {
  resolvedDocxOperationSnapshot,
  resolvedDocxSourceDocument,
  resolvedDocxSourceOperandBlock,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";
import {
  resolvedDocxTableComponentOperands,
  resolvedDocxTableComponents,
  resolvedDocxTableStructureOperandPayload,
  type ResolvedDocxTableComponent,
  type ResolvedDocxTableFormatOperand,
  type ResolvedDocxTableStructureOperand,
} from "./resolved-docx-story-comparison";

type ResolvedDocxTableOperand = ResolvedDocxTableStructureOperand | ResolvedDocxTableFormatOperand;

type DocxComparisonInstructionType =
  | Exclude<DocxComparisonInstruction["type"], "tableStructure" | "tableFormat">
  | ResolvedDocxTableOperand["type"];

export const DOCX_COMPARISON_PREFLIGHT_REASONS = COMPARE_DOCX_PREFLIGHT_REASONS;

export type DocxComparisonPreflightReason = CompareDocxPreflightReason;

export type DocxComparisonPreflightIssue = {
  readonly instructionIndex: number;
  readonly instructionType: DocxComparisonInstructionType;
  readonly reason: DocxComparisonPreflightReason;
  readonly blockId?: string;
  readonly tableGeometry?: TableGeometryUnsupportedIssue;
  readonly tableStructure?: TableStructureUnsupportedIssue;
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
type PreparedSourceSchedule = {
  readonly phase: "source";
  readonly from: number;
  readonly to: number;
};

type PreparedInsertionSchedule = {
  readonly phase: "insertion";
  readonly position: number;
};

type PreparedInstructionPayload =
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
      readonly type: "insertParagraph" | "insertMovedParagraph";
      readonly boundary: ResolvedParagraphBoundary;
      readonly target: DocxComparisonParagraphTarget;
      readonly originalIndex: number;
      readonly schedule: PreparedInsertionSchedule;
    }
  | {
      readonly type: "insertTerminalCarrier";
      readonly boundary: ResolvedParagraphBoundary;
      readonly breakOwner: ResolvedBlock;
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
      readonly type: "removeMovedParagraph";
      readonly source: ResolvedBlock;
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "removeMovedParagraph" }
      >;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "transitionTerminalParagraphs";
      readonly chainStart: ResolvedBlock | null;
      readonly sourceMembers: readonly [
        { readonly source: ResolvedBlock; readonly kind: "del" | "moveFrom" },
        ...{ readonly source: ResolvedBlock; readonly kind: "del" | "moveFrom" }[],
      ];
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "transitionTerminalParagraphs" }
      >;
      readonly originalIndex: number;
      readonly schedule: PreparedSourceSchedule;
    }
  | {
      readonly type: "moveParagraph";
      readonly source: ResolvedBlock;
      readonly boundary: ResolvedParagraphBoundary;
      readonly target: DocxComparisonParagraphTarget;
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
      readonly type: "tableStructureOperand";
      readonly operand: ResolvedDocxTableOperand;
      readonly instructionType: ResolvedDocxTableOperand["type"];
      readonly originalIndex: number;
    };

type WithSemanticGroup<Instruction> = Instruction extends unknown
  ? Instruction & { readonly semanticGroupIndex: number }
  : never;

type PreparedInstruction = WithSemanticGroup<PreparedInstructionPayload>;

type PreparedSourceInstruction = Extract<
  PreparedInstruction,
  { readonly schedule: { readonly phase: "source" } }
>;
type PreparedInsertionInstruction = Extract<
  PreparedInstruction,
  { readonly type: "insertParagraph" | "insertMovedParagraph" | "insertTerminalCarrier" }
>;
type PreparedMoveInstruction = Extract<PreparedInstruction, { readonly type: "moveParagraph" }>;
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

type PreparedTableTask = {
  readonly type: "tableStructure";
  readonly position: number;
  readonly from: number;
  readonly originalIndex: number;
  readonly program: PreparedTableStructureProgram;
  readonly task: PreparedTableStructureTask;
};

type PreparedExecutionTask = PreparedInsertionRun | PreparedSourceTask | PreparedTableTask;

type ExecutionTaskSortKey = {
  readonly from: number;
  readonly originalIndex: number;
};

const executionTaskSortKey = (task: PreparedExecutionTask): ExecutionTaskSortKey => {
  switch (task.type) {
    case "insertionRun":
      return {
        from: task.position,
        originalIndex: task.members[0].instruction.originalIndex,
      };
    case "sourceInstruction":
    case "moveSource":
      return { from: task.from, originalIndex: task.instruction.originalIndex };
    case "tableStructure":
      return { from: task.from, originalIndex: task.originalIndex };
    default: {
      const unreachable: never = task;
      return panic("Unhandled execution task while deriving its sort key", {
        task: unreachable,
      });
    }
  }
};

const executionTasks = (
  instructions: readonly PreparedInstruction[],
  tablePrograms: readonly PreparedTableStructureProgram[],
): readonly PreparedExecutionTask[] => {
  const sourceTasks: PreparedSourceTask[] = [];
  const insertionMembersByPosition = new Map<number, PreparedInsertionMember[]>();
  for (const instruction of instructions) {
    switch (instruction.type) {
      case "insertParagraph":
      case "insertMovedParagraph":
      case "insertTerminalCarrier": {
        const members = insertionMembersByPosition.get(instruction.schedule.position) ?? [];
        members.push({ instruction });
        insertionMembersByPosition.set(instruction.schedule.position, members);
        break;
      }
      case "moveParagraph": {
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
      case "tableStructureOperand":
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
  const tableInstructionIndex = new Map(
    instructions.flatMap((instruction) =>
      instruction.type === "tableStructureOperand"
        ? [[instruction.operand, instruction.originalIndex] as const]
        : [],
    ),
  );
  const tableTasks: PreparedTableTask[] = tablePrograms.flatMap((program) =>
    program.tasks.map((task) => {
      let originalIndex = Number.MAX_SAFE_INTEGER;
      for (const operand of task.operands) {
        const operandIndex =
          tableInstructionIndex.get(operand) ??
          panic("A table-structure task lost its comparison instruction");
        originalIndex = Math.min(originalIndex, operandIndex);
      }
      if (originalIndex === Number.MAX_SAFE_INTEGER) {
        return panic("A table-structure task has no comparison instruction");
      }
      return {
        type: "tableStructure" as const,
        position: task.schedule.position,
        from: task.schedule.phase === "source" ? task.schedule.from : task.schedule.position,
        originalIndex,
        program,
        task,
      };
    }),
  );
  const tableInsertionIndexesByPosition = new Map<number, number[]>();
  for (const task of tableTasks) {
    if (task.task.schedule.phase !== "insertion") continue;
    const indexes = tableInsertionIndexesByPosition.get(task.position) ?? [];
    indexes.push(task.originalIndex);
    tableInsertionIndexesByPosition.set(task.position, indexes);
  }
  const insertionRuns: PreparedInsertionRun[] = [];
  for (const [position, members] of insertionMembersByPosition) {
    const tableIndexes = (tableInsertionIndexesByPosition.get(position) ?? []).toSorted(
      (left, right) => left - right,
    );
    const membersBySegment = new Map<number, PreparedInsertionMember[]>();
    let tableIndexCursor = 0;
    for (const member of members.toSorted(
      (left, right) => left.instruction.originalIndex - right.instruction.originalIndex,
    )) {
      while (
        (tableIndexes[tableIndexCursor] ?? Number.POSITIVE_INFINITY) <
        member.instruction.originalIndex
      ) {
        tableIndexCursor++;
      }
      const segmentMembers = membersBySegment.get(tableIndexCursor) ?? [];
      segmentMembers.push(member);
      membersBySegment.set(tableIndexCursor, segmentMembers);
    }
    for (const segmentMembers of membersBySegment.values()) {
      const first = segmentMembers.at(0) ?? panic("A prepared insertion run has no member");
      insertionRuns.push({
        type: "insertionRun",
        position,
        members: Object.freeze([first, ...segmentMembers.slice(1)]),
      });
    }
  }
  return Object.freeze(
    [...sourceTasks, ...insertionRuns, ...tableTasks].toSorted((left, right) => {
      const byPosition = right.position - left.position;
      if (byPosition !== 0) return byPosition;
      // A boundary insertion at the exact end of a source range runs first:
      // right-to-left execution then leaves every source coordinate original.
      const leftInsertion =
        left.type === "insertionRun" ||
        (left.type === "tableStructure" && left.task.schedule.phase === "insertion");
      const rightInsertion =
        right.type === "insertionRun" ||
        (right.type === "tableStructure" && right.task.schedule.phase === "insertion");
      if (leftInsertion !== rightInsertion) {
        // A paragraph inserted at a table's exact start would move the table's
        // preflighted source coordinate. Execute that source first; insertion
        // at a source range's end still executes first and preserves it.
        if (
          leftInsertion &&
          right.type === "tableStructure" &&
          right.task.schedule.phase === "source" &&
          right.from === left.position
        ) {
          return 1;
        }
        if (
          rightInsertion &&
          left.type === "tableStructure" &&
          left.task.schedule.phase === "source" &&
          left.from === right.position
        ) {
          return -1;
        }
        return leftInsertion ? -1 : 1;
      }
      const leftKey = executionTaskSortKey(left);
      const rightKey = executionTaskSortKey(right);
      const byStart = rightKey.from - leftKey.from;
      if (byStart !== 0) return byStart;
      return rightKey.originalIndex - leftKey.originalIndex;
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
  tableStructure?: TableStructureUnsupportedIssue,
): DocxComparisonPreflightIssue =>
  Object.freeze({
    instructionIndex,
    instructionType:
      instruction.type === "tableStructure" || instruction.type === "tableFormat"
        ? instruction.operation.type
        : instruction.type,
    reason,
    ...(blockId !== undefined && { blockId }),
    ...(tableGeometry !== undefined && { tableGeometry }),
    ...(tableStructure !== undefined && { tableStructure }),
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
  readonly supportedChangeCount: number;
};

type PreparedDocxComparisonState = {
  readonly state: EditorState;
  readonly instructions: readonly PreparedInstruction[];
  readonly semanticGroups: readonly DocxComparisonSemanticGroup[];
  readonly supportedSemanticGroupIndexes: readonly number[];
  readonly tablePrograms: readonly PreparedTableStructureProgram[];
  readonly totalInstructionCount: number;
  lifecycle: "ready" | "consumed";
};

const preparedDocxComparisons = new WeakMap<PreparedDocxComparison, PreparedDocxComparisonState>();

const ownPreparedDocxComparison = ({
  state,
  instructions,
  issues,
  semanticGroups,
  tablePrograms,
  totalInstructionCount,
}: {
  readonly state: EditorState;
  readonly instructions: readonly PreparedInstruction[];
  readonly issues: readonly DocxComparisonPreflightIssue[];
  readonly semanticGroups: readonly DocxComparisonSemanticGroup[];
  readonly tablePrograms: readonly PreparedTableStructureProgram[];
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
  const supportedSemanticGroupIndexes = Object.freeze(
    [...new Set(ownedInstructions.map(({ semanticGroupIndex }) => semanticGroupIndex))].toSorted(
      (left, right) => left - right,
    ),
  );
  const supportedChangeCount = supportedSemanticGroupIndexes.reduce(
    (count, groupIndex) =>
      count +
      (
        semanticGroups[groupIndex] ??
        panic("A prepared instruction lost its semantic group", { groupIndex })
      ).reports.length,
    0,
  );
  const prepared = Object.freeze({
    [PREPARED_DOCX_COMPARISON_BRAND]: true as const,
    issues: ownedIssues,
    supportedInstructionCount: ownedInstructions.length,
    totalInstructionCount,
    supportedChangeCount,
  });
  preparedDocxComparisons.set(prepared, {
    state,
    instructions: ownedInstructions,
    semanticGroups,
    supportedSemanticGroupIndexes,
    tablePrograms: Object.freeze([...tablePrograms]),
    totalInstructionCount,
    lifecycle: "ready",
  });
  return prepared;
};

/** Resolve every instruction against one immutable story before any transaction exists. */
export const preflightDocxComparisonProgram = ({
  state,
  program,
}: {
  readonly state: EditorState;
  readonly program: DocxComparisonProgram;
}): PreparedDocxComparison => {
  const { comparison, sourceSnapshot: snapshot, semanticGroups, instructions } = program.consume();
  const operationSnapshot = resolvedDocxOperationSnapshot(snapshot);
  if (state.doc !== resolvedDocxSourceDocument(snapshot)) {
    return ownPreparedDocxComparison({
      state,
      instructions: [],
      issues: instructions.map((instruction, instructionIndex) =>
        issue(instruction, instructionIndex, "source-expectation-mismatch"),
      ),
      semanticGroups,
      tablePrograms: Object.freeze([]),
      totalInstructionCount: instructions.length,
    });
  }
  const resolver = FolioStableBlockResolver.create(state.doc, operationSnapshot);
  const styleResolver = getDocumentStyleResolver(state);
  const prepared: PreparedInstructionPayload[] = [];
  const issues: DocxComparisonPreflightIssue[] = [];
  const resolveSource = (source: DocxComparisonSourceOperand) =>
    resolveExpectedBlock({ snapshot, resolver, source });
  const sourceBlockId = (source: DocxComparisonSourceOperand): string =>
    resolvedDocxSourceOperandBlock(source, snapshot).identity.id;
  for (const [instructionIndex, instruction] of instructions.entries()) {
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
          let reason: DocxComparisonPreflightReason = "unrepresentable-text-range";
          if (range.reason === "pending-run-change") reason = "pending-run-change";
          if (range.reason === "source-formatting-mismatch") {
            reason = "source-formatting-mismatch";
          }
          issues.push(
            issue(instruction, instructionIndex, reason, sourceBlockId(instruction.source)),
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
      case "insertParagraph":
      case "insertMovedParagraph":
      case "insertTerminalCarrier": {
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
        if (instruction.type === "insertTerminalCarrier") {
          const breakOwner = resolveSource(instruction.breakOwner);
          if (breakOwner.type === "unsupported") {
            issues.push(
              issue(
                instruction,
                instructionIndex,
                breakOwner.reason,
                sourceBlockId(instruction.breakOwner),
              ),
            );
            break;
          }
          const breakOwnerAttrs = expectParagraphAttrs(breakOwner.block.node);
          if (
            breakOwnerAttrs.pPrMark != null ||
            hasSerializableParagraphPropertyChange(breakOwnerAttrs._propertyChanges)
          ) {
            issues.push(
              issue(
                instruction,
                instructionIndex,
                "pending-paragraph-change",
                sourceBlockId(instruction.breakOwner),
              ),
            );
            break;
          }
          if (breakOwner.block.from !== boundary.from) {
            issues.push(
              issue(
                instruction,
                instructionIndex,
                "unrepresentable-paragraph-boundary",
                sourceBlockId(instruction.breakOwner),
              ),
            );
            break;
          }
          prepared.push(
            Object.freeze({
              type: instruction.type,
              boundary,
              breakOwner: breakOwner.block,
              target: instruction.target,
              originalIndex: instructionIndex,
              schedule: Object.freeze({
                phase: "insertion",
                position: boundary.insertionPosition,
              }),
            }),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: instruction.type,
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
      case "removeMovedParagraph": {
        const source = resolveSource(instruction.source);
        if (source.type === "unsupported") {
          issues.push(
            issue(instruction, instructionIndex, source.reason, sourceBlockId(instruction.source)),
          );
          break;
        }
        const sourcePosition = state.doc.resolve(source.block.from);
        let ownsRemovalBoundary = false;
        switch (instruction.removalBoundary.type) {
          case "successorParagraph": {
            const successor = resolveSource(instruction.removalBoundary.successor);
            if (successor.type === "unsupported") {
              issues.push(
                issue(
                  instruction,
                  instructionIndex,
                  successor.reason,
                  sourceBlockId(instruction.removalBoundary.successor),
                ),
              );
              break;
            }
            ownsRemovalBoundary = source.block.to === successor.block.from;
            break;
          }
          case "successorTable": {
            const firstTableBlock = resolveSource(instruction.removalBoundary.firstBlock);
            if (firstTableBlock.type === "unsupported") {
              issues.push(
                issue(
                  instruction,
                  instructionIndex,
                  firstTableBlock.reason,
                  sourceBlockId(instruction.removalBoundary.firstBlock),
                ),
              );
              break;
            }
            const tablePosition = state.doc.resolve(firstTableBlock.block.from);
            for (let depth = 1; depth <= tablePosition.depth; depth++) {
              if (
                tablePosition.node(depth - 1) === sourcePosition.parent &&
                tablePosition.node(depth).type.spec["tableRole"] === "table"
              ) {
                ownsRemovalBoundary = tablePosition.before(depth) === source.block.to;
                break;
              }
            }
            break;
          }
          default: {
            const unreachable: never = instruction.removalBoundary;
            return panic("Unhandled moved-source removal boundary", { boundary: unreachable });
          }
        }
        if (
          !ownsRemovalBoundary ||
          paragraphEndsItsContainer(sourcePosition, source.block.node.type.name)
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
            type: "removeMovedParagraph",
            source: source.block,
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: sourceSchedule(source.block),
          }),
        );
        break;
      }
      case "transitionTerminalParagraphs": {
        const chainStart =
          instruction.chainStart === null ? null : resolveSource(instruction.chainStart);
        const resolveMember = ({ source, kind }: (typeof instruction.sourceMembers)[number]) => ({
          source: resolveSource(source),
          operand: source,
          kind,
        });
        const members = Object.freeze([
          resolveMember(instruction.sourceMembers[0]),
          ...instruction.sourceMembers.slice(1).map(resolveMember),
        ] as const);
        if (chainStart?.type === "unsupported") {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              chainStart.reason,
              instruction.chainStart === null ? undefined : sourceBlockId(instruction.chainStart),
            ),
          );
          break;
        }
        const rejectedMember = members.find(({ source }) => source.type === "unsupported");
        if (rejectedMember?.source.type === "unsupported") {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              rejectedMember.source.reason,
              sourceBlockId(rejectedMember.operand),
            ),
          );
          break;
        }
        const requireResolvedMember = ({ source, kind }: (typeof members)[number]) => {
          if (source.type !== "ready") {
            return panic("A terminal transition retained an unresolved source member");
          }
          return Object.freeze({ source: source.block, kind });
        };
        const resolvedMembers = Object.freeze([
          requireResolvedMember(members[0]),
          ...members.slice(1).map(requireResolvedMember),
        ] as const);
        const first = resolvedMembers[0];
        const last = resolvedMembers.at(-1);
        if (!last) return panic("A terminal transition has no source member");
        const resolvedChainStart = chainStart?.type === "ready" ? chainStart.block : null;
        const sourceTables = instruction.sourceTables.map(({ operation }) => {
          const payload = resolvedDocxTableStructureOperandPayload(operation, comparison);
          if (payload.type !== "deleteTable") {
            return panic("A terminal transition source table lost its deletion role");
          }
          const source = resolveSource(payload.source);
          if (source.type === "unsupported") {
            return Object.freeze({
              type: "unsupported" as const,
              operand: payload.source,
              reason: source.reason,
            });
          }
          const position = state.doc.resolve(source.block.from);
          for (let depth = 1; depth <= position.depth; depth++) {
            const node = position.node(depth);
            if (position.node(depth - 1) === state.doc && node.type.spec["tableRole"] === "table") {
              const from = position.before(depth);
              return Object.freeze({
                type: "ready" as const,
                operation,
                range: Object.freeze({ from, to: from + node.nodeSize }),
              });
            }
          }
          return Object.freeze({
            type: "invalid" as const,
            operand: payload.source,
          });
        });
        let rejectedSourceTable = false;
        for (const sourceTable of sourceTables) {
          if (sourceTable.type === "ready") continue;
          issues.push(
            issue(
              instruction,
              instructionIndex,
              sourceTable.type === "unsupported"
                ? sourceTable.reason
                : "unrepresentable-paragraph-boundary",
              sourceBlockId(sourceTable.operand),
            ),
          );
          rejectedSourceTable = true;
          break;
        }
        if (rejectedSourceTable) break;
        const expectedParagraphs = [
          ...(resolvedChainStart === null ? [] : [resolvedChainStart]),
          ...resolvedMembers.map(({ source }) => source),
        ];
        const readySourceTables = sourceTables.map((table) => {
          if (table.type !== "ready") {
            return panic("A terminal transition retained an unresolved source table");
          }
          return table;
        });
        type PhysicalTerminalMember =
          | { readonly type: "paragraph"; readonly range: ResolvedBlock }
          | {
              readonly type: "table";
              readonly range: { readonly from: number; readonly to: number };
              readonly operation: ResolvedDocxTableStructureOperand;
            };
        const physicalSequence: PhysicalTerminalMember[] = [];
        let paragraphIndex = 0;
        let tableIndex = 0;
        while (
          paragraphIndex < expectedParagraphs.length ||
          tableIndex < readySourceTables.length
        ) {
          const paragraph = expectedParagraphs[paragraphIndex];
          const table = readySourceTables[tableIndex];
          if (paragraph && (!table || paragraph.from < table.range.from)) {
            physicalSequence.push({ type: "paragraph", range: paragraph });
            paragraphIndex++;
            continue;
          }
          if (!table) return panic("A terminal transition lost its physical source member");
          physicalSequence.push({ type: "table", range: table.range, operation: table.operation });
          tableIndex++;
        }
        const isContiguous = physicalSequence.every(
          ({ range }, index) => index === 0 || physicalSequence[index - 1]?.range.to === range.from,
        );
        const retainsCanonicalOrder =
          expectedParagraphs.every(
            ({ from }, index) =>
              index === 0 || (expectedParagraphs[index - 1]?.from ?? from) < from,
          ) &&
          readySourceTables.every(
            ({ range }, index) =>
              index === 0 || (readySourceTables[index - 1]?.range.from ?? range.from) < range.from,
          );
        const successorByParagraphFrom = new Map<number, PhysicalTerminalMember | undefined>();
        for (const [index, member] of physicalSequence.entries()) {
          if (member.type === "paragraph") {
            successorByParagraphFrom.set(member.range.from, physicalSequence[index + 1]);
          }
        }
        const chainStartOwnsFirstEdge =
          instruction.targetCarrierKind === "surviving" ||
          instruction.targetCarrierKind === "pairedRewrite";
        const edgeOwners = chainStartOwnsFirstEdge
          ? [
              ...(resolvedChainStart === null ? [] : [resolvedChainStart]),
              ...resolvedMembers.slice(0, -1).map(({ source }) => source),
            ]
          : resolvedMembers.slice(0, -1).map(({ source }) => source);
        const pendingParagraphChange = [...edgeOwners, last.source].find((block) => {
          const attrs = expectParagraphAttrs(block.node);
          return (
            attrs.pPrMark != null || hasSerializableParagraphPropertyChange(attrs._propertyChanges)
          );
        });
        const allJoinable = edgeOwners.every((block) => {
          const successor = successorByParagraphFrom.get(block.from);
          return successor?.type === "table" || canJoin(state.doc, block.to);
        });
        if (
          !isContiguous ||
          !retainsCanonicalOrder ||
          (chainStartOwnsFirstEdge && resolvedChainStart === null) ||
          !paragraphEndsItsContainer(
            state.doc.resolve(last.source.from),
            last.source.node.type.name,
          ) ||
          resolvedMembers
            .slice(0, -1)
            .some(({ source }) =>
              paragraphEndsItsContainer(state.doc.resolve(source.from), source.node.type.name),
            ) ||
          !allJoinable
        ) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "unrepresentable-paragraph-boundary",
              sourceBlockId(instruction.sourceMembers[0].source),
            ),
          );
          break;
        }
        if (pendingParagraphChange) {
          const member = [...edgeOwners, last.source].find(
            (block) => block.from === pendingParagraphChange.from,
          );
          const operand = instruction.sourceMembers.find(({ source }) => {
            const resolved = resolveSource(source);
            return resolved.type === "ready" && resolved.block.from === member?.from;
          });
          issues.push(
            issue(
              instruction,
              instructionIndex,
              "pending-paragraph-change",
              operand === undefined ? undefined : sourceBlockId(operand.source),
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "transitionTerminalParagraphs",
            chainStart: resolvedChainStart,
            sourceMembers: resolvedMembers,
            semantic: instruction,
            originalIndex: instructionIndex,
            schedule: Object.freeze({
              phase: "source",
              from: resolvedChainStart?.from ?? first.source.from,
              to: last.source.to,
            }),
          }),
        );
        break;
      }
      case "moveParagraph": {
        const source = resolveSource(instruction.source);
        const boundary = resolveParagraphBoundary(snapshot, resolver, instruction.boundary);
        if (source.type === "unsupported") {
          issues.push(
            issue(instruction, instructionIndex, source.reason, sourceBlockId(instruction.source)),
          );
          break;
        }
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
        const sourcePosition = state.doc.resolve(source.block.from);
        let ownsRemovalBoundary = false;
        let removalBoundaryFailed = false;
        switch (instruction.removalBoundary.type) {
          case "successorParagraph": {
            const successor = resolveSource(instruction.removalBoundary.successor);
            if (successor.type === "unsupported") {
              issues.push(
                issue(
                  instruction,
                  instructionIndex,
                  successor.reason,
                  sourceBlockId(instruction.removalBoundary.successor),
                ),
              );
              removalBoundaryFailed = true;
              break;
            }
            ownsRemovalBoundary = source.block.to === successor.block.from;
            break;
          }
          case "successorTable": {
            const firstTableBlock = resolveSource(instruction.removalBoundary.firstBlock);
            if (firstTableBlock.type === "unsupported") {
              issues.push(
                issue(
                  instruction,
                  instructionIndex,
                  firstTableBlock.reason,
                  sourceBlockId(instruction.removalBoundary.firstBlock),
                ),
              );
              removalBoundaryFailed = true;
              break;
            }
            const tablePosition = state.doc.resolve(firstTableBlock.block.from);
            for (let depth = 1; depth <= tablePosition.depth; depth++) {
              if (
                tablePosition.node(depth - 1) === sourcePosition.parent &&
                tablePosition.node(depth).type.spec["tableRole"] === "table"
              ) {
                ownsRemovalBoundary = tablePosition.before(depth) === source.block.to;
                break;
              }
            }
            break;
          }
          default: {
            const unreachable: never = instruction.removalBoundary;
            return panic("Unhandled DOCX paragraph-removal boundary", {
              boundary: unreachable,
            });
          }
        }
        if (removalBoundaryFailed) break;
        if (
          !ownsRemovalBoundary ||
          paragraphEndsItsContainer(sourcePosition, source.block.node.type.name)
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
        if (firstSource.type === "unsupported") {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              firstSource.reason,
              sourceBlockId(instruction.firstSource),
            ),
          );
          break;
        }
        if (secondSource.type === "unsupported") {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              secondSource.reason,
              sourceBlockId(instruction.secondSource),
            ),
          );
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
      case "tableStructure":
      case "tableFormat":
        break;
      default: {
        const unreachable: never = instruction;
        return panic("Unhandled DOCX comparison instruction during preflight", {
          instruction: unreachable,
        });
      }
    }
  }
  const tableEntries = instructions.flatMap((instruction, originalIndex) =>
    instruction.type === "tableStructure" || instruction.type === "tableFormat"
      ? [{ instruction, operand: instruction.operation, originalIndex }]
      : [],
  );
  const tableEntryByOperand = new Map(tableEntries.map((entry) => [entry.operand, entry]));
  const tableComponents = resolvedDocxTableComponents({
    comparison,
    operands: tableEntries.map(({ operand }) => operand),
  });
  const entriesByComponent = new Map(
    tableComponents.map((component) => [
      component,
      resolvedDocxTableComponentOperands(component, comparison).map(
        (operand) =>
          tableEntryByOperand.get(operand) ??
          panic("A canonical table component lost its comparison instruction"),
      ),
    ]),
  );
  const tablePrograms: PreparedTableStructureProgram[] = [];
  const tablePreflight =
    tableComponents.length === 0
      ? null
      : preflightTableStructureComponents({
          doc: state.doc,
          comparison,
          components: tableComponents,
        });
  if (tablePreflight?.status === "unsupported") {
    for (const { instruction, originalIndex } of tableEntries) {
      issues.push(
        issue(
          instruction,
          originalIndex,
          tablePreflight.issue.reason === "unrepresentable-table-geometry"
            ? "unrepresentable-table-geometry"
            : "unrepresentable-table-structure",
          undefined,
          tablePreflight.issue.reason === "unrepresentable-table-geometry"
            ? tablePreflight.issue.issue
            : undefined,
          tablePreflight.issue,
        ),
      );
    }
  } else {
    const readyComponents = new Map<ResolvedDocxTableComponent, PreparedTableStructureProgram>();
    for (const { component, result } of tablePreflight?.components ?? []) {
      const entries =
        entriesByComponent.get(component) ??
        panic("A preflighted table component lost its comparison instructions");
      if (result.status === "unsupported") {
        for (const { instruction, originalIndex } of entries) {
          issues.push(
            issue(
              instruction,
              originalIndex,
              result.issue.reason === "unrepresentable-table-geometry"
                ? "unrepresentable-table-geometry"
                : "unrepresentable-table-structure",
              undefined,
              result.issue.reason === "unrepresentable-table-geometry"
                ? result.issue.issue
                : undefined,
              result.issue,
            ),
          );
        }
        continue;
      }
      readyComponents.set(component, result.program);
    }
    const failedGroups = new Set(
      issues.map(({ instructionIndex }) => {
        const instruction = instructions[instructionIndex];
        return (
          instruction ??
          panic("A DOCX comparison preflight issue lost its instruction", { instructionIndex })
        ).semanticGroupIndex;
      }),
    );
    let removedComponent = true;
    while (removedComponent) {
      removedComponent = false;
      for (const [component] of readyComponents) {
        const entries =
          entriesByComponent.get(component) ??
          panic("A ready table component lost its comparison instructions");
        if (!entries.some(({ instruction }) => failedGroups.has(instruction.semanticGroupIndex))) {
          continue;
        }
        readyComponents.delete(component);
        removedComponent = true;
        for (const { instruction, originalIndex } of entries) {
          issues.push(issue(instruction, originalIndex, "semantic-group-incomplete"));
          failedGroups.add(instruction.semanticGroupIndex);
        }
      }
    }
    for (const [component, tableComponentProgram] of readyComponents) {
      tablePrograms.push(tableComponentProgram);
      const entries =
        entriesByComponent.get(component) ??
        panic("A ready table component lost its comparison instructions");
      for (const { operand, originalIndex } of entries) {
        prepared.push(
          Object.freeze({
            type: "tableStructureOperand",
            operand,
            instructionType: operand.type,
            originalIndex,
          }),
        );
      }
    }
  }
  const failedSemanticGroupIndexes = new Set(
    issues.map(({ instructionIndex }) => {
      const instruction = instructions[instructionIndex];
      return (
        instruction ??
        panic("A DOCX comparison preflight issue lost its instruction", { instructionIndex })
      ).semanticGroupIndex;
    }),
  );
  const groupedPrepared: PreparedInstruction[] = [];
  const groupedIssues = [...issues];
  for (const candidate of prepared) {
    const semantic =
      instructions[candidate.originalIndex] ??
      panic("A prepared DOCX instruction lost its semantic input", {
        instructionIndex: candidate.originalIndex,
      });
    if (failedSemanticGroupIndexes.has(semantic.semanticGroupIndex)) {
      groupedIssues.push(issue(semantic, candidate.originalIndex, "semantic-group-incomplete"));
      continue;
    }
    groupedPrepared.push(
      Object.freeze({
        ...candidate,
        semanticGroupIndex: semantic.semanticGroupIndex,
      }),
    );
  }
  return ownPreparedDocxComparison({
    state,
    instructions: groupedPrepared,
    issues: groupedIssues.toSorted((left, right) => left.instructionIndex - right.instructionIndex),
    semanticGroups,
    tablePrograms,
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
  const node =
    text.transaction.doc.nodeAt(source.from) ??
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
    if (instruction.type === "insertTerminalCarrier") {
      const textRevisionId = instruction.target.text.length > 0 ? nextRevisionId++ : null;
      const breakRevisionId = nextRevisionId++;
      const tracked = trackedParagraph(
        tr,
        instruction.boundary,
        instruction.target,
        textRevisionId,
        breakRevisionId,
        author,
        date,
      );
      prepared.push({
        instruction,
        paragraph: tracked.type.create({ ...tracked.attrs, pPrMark: null }, tracked.content),
        revisionIds: Object.freeze([
          ...(textRevisionId === null ? [] : [textRevisionId]),
          breakRevisionId,
        ]),
      });
      continue;
    }
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
  for (const member of prepared) {
    if (member.instruction.type !== "insertTerminalCarrier") continue;
    const breakRevisionId =
      member.revisionIds.at(-1) ?? panic("A terminal carrier lost its break id");
    tr.setNodeAttribute(member.instruction.breakOwner.from, "pPrMark", {
      kind: "ins",
      info: { id: breakRevisionId, author, date },
    });
  }
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
    for (const ownership of member.instruction.target.inlineOwnership) {
      applyExactInlineOwnership({
        tr,
        from: position + 1 + ownership.startOffset,
        to: position + 1 + ownership.endOffset,
        containers: ownership.containers,
      });
    }
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
  readonly instructionType: DocxComparisonInstructionType;
  readonly revisionIds: readonly number[];
  readonly tableGeometry?: TableGeometryExecutionReceipt;
};

export type DocxComparisonExecutionReceipt = {
  readonly instructions: readonly DocxComparisonInstructionReceipt[];
  /** Reports whose complete semantic instruction group executed. */
  readonly changes: readonly CompareChange[];
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

const unsupportedExecution = (
  executionIssue: DocxComparisonExecutionIssue,
): DocxComparisonExecutionResult =>
  Object.freeze({ status: "unsupported", issue: Object.freeze(executionIssue) });

type PropertyRevisionAllocation = {
  nextRevisionId: number;
  readonly revisionIds: number[];
};

type AllocatePropertyRevisionOptions = {
  readonly allocation: PropertyRevisionAllocation;
  readonly author: string;
  readonly date: string;
};

const allocatePropertyRevision =
  ({
    allocation,
    author,
    date,
  }: AllocatePropertyRevisionOptions): (() => {
    readonly id: number;
    readonly author: string;
    readonly date: string;
  }) =>
  () => {
    const id = allocation.nextRevisionId++;
    allocation.revisionIds.push(id);
    return { id, author, date };
  };

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
    readonly instructionType: DocxComparisonInstructionType;
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
    readonly instructionType: DocxComparisonInstructionType;
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
  const tableInstructionByOperand = new Map(
    owned.instructions.flatMap((instruction) =>
      instruction.type === "tableStructureOperand"
        ? [[instruction.operand, instruction] as const]
        : [],
    ),
  );
  for (const tableProgram of owned.tablePrograms) {
    const result = executeTableStructureGeometry({
      tr,
      program: tableProgram,
      revision: { author, date: revisionStamp.date },
      revisionId,
    });
    if ("issue" in result) {
      const geometryOperand =
        tableProgram.geometryOperand ??
        panic("A table geometry execution issue lost its canonical operand");
      const instruction =
        tableInstructionByOperand.get(geometryOperand) ??
        panic("A table geometry operand lost its comparison instruction");
      return unsupportedExecution({
        reason: "table-geometry-execution",
        instructionIndex: instruction.originalIndex,
        issue: result.issue,
      });
    }
    tr = result.transaction;
    revisionId = result.nextRevisionId;
    for (const applied of result.applied) {
      const instruction =
        tableInstructionByOperand.get(applied.operand) ??
        panic("A table geometry result lost its comparison instruction");
      recordPart({
        instructionIndex: instruction.originalIndex,
        instructionType: instruction.instructionType,
        part: 0,
        revisionIds: applied.revisionIds,
        ...(applied.tableGeometry === undefined ? {} : { tableGeometry: applied.tableGeometry }),
      });
    }
  }
  const ordered = executionTasks(owned.instructions, owned.tablePrograms);
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
    if (task.type === "tableStructure") {
      const result = executeTableStructureTask({
        tr,
        program: task.program,
        task: task.task,
        revision: { author, date: revisionStamp.date },
        revisionId,
      });
      tr = result.transaction;
      revisionId = result.nextRevisionId;
      localPositionMappingSteps += result.localPositionMappingSteps;
      for (const applied of result.applied) {
        const instruction =
          tableInstructionByOperand.get(applied.operand) ??
          panic("A table structure result lost its comparison instruction");
        recordPart({
          instructionIndex: instruction.originalIndex,
          instructionType: instruction.instructionType,
          part: 0,
          revisionIds: applied.revisionIds,
        });
      }
      if (tr.steps.length === stepsBefore) {
        return panic("A preflighted table-structure task produced no transaction step");
      }
      continue;
    }
    const instructionRevisionIds: number[] = [];
    if (task.type === "moveSource") {
      const { instruction } = task;
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
    const { instruction } = task;
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
      case "removeMovedParagraph": {
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
        break;
      }
      case "setParagraphProperties": {
        const at = instruction.source.from;
        const node = tr.doc.nodeAt(at) ?? panic("A preflighted paragraph change lost its source");
        const propertyRevisions: PropertyRevisionAllocation = {
          nextRevisionId: revisionId,
          revisionIds: [],
        };
        const appliedProperties = applyBlockParagraphProperties({
          tr,
          position: at,
          node,
          properties: instruction.semantic.targetProperties,
          styleResolver,
          numbering,
          revisionInfo: allocatePropertyRevision({
            allocation: propertyRevisions,
            author,
            date: revisionStamp.date,
          }),
        });
        tr = appliedProperties.tr;
        revisionId = propertyRevisions.nextRevisionId;
        instructionRevisionIds.push(...propertyRevisions.revisionIds);
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
          for (const ownership of instruction.semantic.separatorInlineOwnership) {
            applyExactInlineOwnership({
              tr,
              from: insertAt + ownership.startOffset,
              to: insertAt + ownership.endOffset,
              containers: ownership.containers,
            });
          }
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
      case "transitionTerminalParagraphs": {
        if (instruction.semantic.targetCarrierKind === "pairedRewrite") {
          const source =
            instruction.chainStart ??
            panic("A terminal chain-start rewrite lost its source operand");
          const deleted = markParagraphTextDeletion({
            tr,
            source,
            revisionId,
            author,
            date: revisionStamp.date,
            kind: "del",
          });
          tr = deleted.transaction;
          revisionId = deleted.nextRevisionId;
          instructionRevisionIds.push(...deleted.revisionIds);
        }
        for (const member of instruction.sourceMembers) {
          const deleted = markParagraphTextDeletion({
            tr,
            source: member.source,
            revisionId,
            author,
            date: revisionStamp.date,
            kind: member.kind,
          });
          tr = deleted.transaction;
          revisionId = deleted.nextRevisionId;
          instructionRevisionIds.push(...deleted.revisionIds);
        }
        const chainStartOwnsFirstEdge =
          instruction.semantic.targetCarrierKind === "surviving" ||
          instruction.semantic.targetCarrierKind === "pairedRewrite";
        const edgeOwners = chainStartOwnsFirstEdge
          ? [
              ...(instruction.chainStart === null ? [] : [instruction.chainStart]),
              ...instruction.sourceMembers.slice(0, -1).map(({ source }) => source),
            ]
          : instruction.sourceMembers.slice(0, -1).map(({ source }) => source);
        const edgeKinds = chainStartOwnsFirstEdge
          ? instruction.sourceMembers.map(({ kind }) => kind)
          : instruction.sourceMembers.slice(0, -1).map(({ kind }) => kind);
        for (const [index, owner] of edgeOwners.entries()) {
          const kind =
            edgeKinds[index] ?? panic("A terminal transition edge lost its revision kind");
          const paragraphRevisionId = revisionId++;
          tr.setNodeAttribute(owner.from, "pPrMark", {
            kind,
            info: { id: paragraphRevisionId, author, date: revisionStamp.date },
          });
          instructionRevisionIds.push(paragraphRevisionId);
        }
        const carrier = instruction.sourceMembers.at(-1)?.source;
        if (!carrier) return panic("A terminal transition lost its physical carrier");
        if (
          instruction.semantic.targetCarrierKind !== "surviving" &&
          instruction.semantic.targetCarrier.text.length > 0
        ) {
          const insertionType =
            tr.doc.type.schema.marks["insertion"] ??
            panic("A terminal transition lost tracked-text schema support");
          const textRevisionId = revisionId++;
          const from = carrier.from + 1;
          tr.insertText(instruction.semantic.targetCarrier.text, from);
          tr.addMark(
            from,
            from + instruction.semantic.targetCarrier.text.length,
            insertionType.create({
              revisionId: textRevisionId,
              author,
              date: revisionStamp.date,
              ...(instruction.semantic.targetCarrierKind === "moveTo" && {
                moveKind: "moveTo",
              }),
            }),
          );
          for (const ownership of instruction.semantic.targetCarrier.inlineOwnership) {
            applyExactInlineOwnership({
              tr,
              from: from + ownership.startOffset,
              to: from + ownership.endOffset,
              containers: ownership.containers,
            });
          }
          for (const run of instruction.semantic.targetCarrier.runs) {
            applyExactDirectFormatting({
              tr,
              from: from + run.startOffset,
              to: from + run.endOffset,
              formatting: run.formatting,
              styleResolver,
            });
          }
          instructionRevisionIds.push(textRevisionId);
        }
        if (instruction.semantic.targetCarrierProperties !== null) {
          const liveCarrier =
            tr.doc.nodeAt(carrier.from) ?? panic("A terminal transition lost its live carrier");
          const propertyRevisions: PropertyRevisionAllocation = {
            nextRevisionId: revisionId,
            revisionIds: [],
          };
          tr = applyBlockParagraphProperties({
            tr,
            position: carrier.from,
            node: liveCarrier,
            properties: instruction.semantic.targetCarrierProperties,
            styleResolver,
            numbering,
            revisionInfo: allocatePropertyRevision({
              allocation: propertyRevisions,
              author,
              date: revisionStamp.date,
            }),
          }).tr;
          revisionId = propertyRevisions.nextRevisionId;
          instructionRevisionIds.push(...propertyRevisions.revisionIds);
        }
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
      const expectedPartCount = instruction.type === "moveParagraph" ? 2 : 1;
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
  const changes = owned.supportedSemanticGroupIndexes
    .flatMap((groupIndex) => {
      const group =
        owned.semanticGroups[groupIndex] ??
        panic("A completed DOCX comparison lost its semantic group", { groupIndex });
      return group.reports;
    })
    .toSorted((left, right) => left.sequence - right.sequence)
    .map(({ change }) => change);
  return Object.freeze({
    status: "executed" as const,
    receipt: Object.freeze({
      instructions: Object.freeze(receipts),
      changes: Object.freeze(changes),
      nextRevisionId: revisionId,
      executionTaskCount:
        owned.tablePrograms.filter(({ geometryOperand }) => geometryOperand !== undefined).length +
        ordered.length,
      insertionRunCount,
      localPositionMappingSteps,
      transaction: tr,
    }),
  });
};
