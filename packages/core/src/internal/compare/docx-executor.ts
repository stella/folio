import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { canJoin, canSplit } from "prosemirror-transform";

import { applyBlockParagraphProperties, withRotatedAddedFinalBreaks } from "../../ai-edits/apply";
import { buildCleanBlockText } from "../../ai-edits/clean-text";
import { FolioStableBlockResolver } from "./stable-block-resolution";
import { storyTablesOf } from "../../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../../ai-edits/types";
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
import { findOutermostTableBoundary } from "../../ai-edits/table-targets";
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

type ResolvedAnchor = ResolvedBlock & {
  readonly insertionPosition: number;
};

type ReadyTextRange = Extract<DocxTextRangePreflight, { readonly type: "ready" }>;
type ReadyTableGeometryProgram = Extract<
  TableGeometryPreflightResult,
  { readonly status: "ready" }
>["program"];

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
      readonly schedule: { readonly phase: "source"; readonly position: number };
    }
  | {
      readonly type: "insertParagraph";
      readonly anchor: ResolvedAnchor;
      readonly target: DocxComparisonParagraphTarget;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "insertion"; readonly position: number };
    }
  | {
      readonly type: "deleteParagraph" | "mergeTerminalCarrier";
      readonly source: ResolvedBlock;
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "deleteParagraph" | "mergeTerminalCarrier" }
      >;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "source"; readonly position: number };
    }
  | {
      readonly type: "moveParagraph";
      readonly source: ResolvedBlock;
      readonly anchor: ResolvedAnchor;
      readonly target: DocxComparisonParagraphTarget;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "insertion"; readonly position: number };
    }
  | {
      readonly type: "splitParagraph";
      readonly source: ResolvedBlock;
      readonly splitPosition: number;
      readonly first: ReadyTextRange;
      readonly second: ReadyTextRange;
      readonly semantic: Extract<DocxComparisonInstruction, { readonly type: "splitParagraph" }>;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "source"; readonly position: number };
    }
  | {
      readonly type: "mergeParagraphs";
      readonly firstSource: ResolvedBlock;
      readonly secondSource: ResolvedBlock;
      readonly first: ReadyTextRange;
      readonly second: ReadyTextRange;
      readonly semantic: Extract<DocxComparisonInstruction, { readonly type: "mergeParagraphs" }>;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "source"; readonly position: number };
    }
  | {
      readonly type: "setParagraphProperties";
      readonly source: ResolvedBlock;
      readonly semantic: Extract<
        DocxComparisonInstruction,
        { readonly type: "setParagraphProperties" }
      >;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "source"; readonly position: number };
    }
  | {
      readonly type: "matchTableGeometry";
      readonly program: ReadyTableGeometryProgram;
      readonly originalIndex: number;
      readonly schedule: { readonly phase: "geometry" };
    };

type PreparedInstructionForPhase<Phase extends PreparedInstruction["schedule"]["phase"]> = Extract<
  PreparedInstruction,
  { readonly schedule: { readonly phase: Phase } }
>;

const preparedForPhase = <Phase extends PreparedInstruction["schedule"]["phase"]>(
  instruction: PreparedInstruction,
  phase: Phase,
): instruction is PreparedInstructionForPhase<Phase> => instruction.schedule.phase === phase;
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
const resolveAnchor = (
  doc: PMNode,
  blockById: ReadonlyMap<string, FolioAIEditSnapshot["blocks"][number]>,
  resolver: FolioStableBlockResolver,
  anchor: Extract<DocxComparisonInstruction, { readonly type: "insertParagraph" }>["anchor"],
): ResolvedAnchor | null => {
  const source = blockById.get(anchor.blockId);
  if (!source) return null;
  const resolved = resolver.resolve(anchor.blockId);
  if (resolved.type === "unsupported") return null;
  const tableBoundary = findOutermostTableBoundary(doc, resolved.blockFrom);
  const insertionPosition =
    anchor.position === "after"
      ? (tableBoundary?.after ?? resolved.blockTo)
      : (tableBoundary?.before ?? resolved.blockFrom);
  return Object.freeze({
    node: resolved.blockNode,
    from: resolved.blockFrom,
    to: resolved.blockTo,
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
  const blockById = new Map(operationSnapshot.blocks.map((block) => [block.id, block]));
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
            schedule: Object.freeze({ phase: "source", position: source.block.from }),
          }),
        );
        break;
      }
      case "insertParagraph": {
        const anchor = resolveAnchor(state.doc, blockById, resolver, instruction.anchor);
        if (!anchor) {
          issues.push(
            issue(instruction, instructionIndex, "missing-anchor", instruction.anchor.blockId),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "insertParagraph",
            anchor,
            target: instruction.target,
            originalIndex: instructionIndex,
            schedule: Object.freeze({ phase: "insertion", position: anchor.insertionPosition }),
          }),
        );
        break;
      }
      case "deleteParagraph":
      case "mergeTerminalCarrier": {
        const source = resolveSource(instruction.source);
        if (source.type === "unsupported") {
          issues.push(
            issue(instruction, instructionIndex, source.reason, sourceBlockId(instruction.source)),
          );
          break;
        }
        if (
          instruction.type === "deleteParagraph" &&
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
            schedule: Object.freeze({ phase: "source", position: source.block.from }),
          }),
        );
        break;
      }
      case "moveParagraph": {
        const source = resolveSource(instruction.source);
        const anchor = resolveAnchor(state.doc, blockById, resolver, instruction.anchor);
        if (source.type === "unsupported" || !anchor) {
          issues.push(
            issue(
              instruction,
              instructionIndex,
              source.type === "unsupported" ? source.reason : "missing-anchor",
              source.type === "unsupported"
                ? sourceBlockId(instruction.source)
                : instruction.anchor.blockId,
            ),
          );
          break;
        }
        prepared.push(
          Object.freeze({
            type: "moveParagraph",
            source: source.block,
            anchor,
            target: instruction.target,
            originalIndex: instructionIndex,
            schedule: Object.freeze({ phase: "insertion", position: anchor.insertionPosition }),
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
            schedule: Object.freeze({ phase: "source", position: source.block.from }),
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
            schedule: Object.freeze({ phase: "source", position: firstSource.block.from }),
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
            schedule: Object.freeze({ phase: "source", position: source.block.from }),
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
  anchor: ResolvedAnchor,
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
  return anchor.node.type.create(
    {
      ...stripBlockIdentityAttrs(anchor.node.attrs),
      pPrMark: { kind, info: { id: paragraphRevisionId, author, date } },
      _propertyChanges: null,
    },
    content,
  );
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
  const at = tr.mapping.map(source.from);
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
  if (!paragraphEndsItsContainer(tr.doc.resolve(at), node.type.name)) {
    const paragraphRevisionId = nextRevisionId++;
    tr.setNodeAttribute(at, "pPrMark", {
      kind,
      info: { id: paragraphRevisionId, author, date },
    });
    revisionIds.push(paragraphRevisionId);
  }
  return { transaction: tr, nextRevisionId, revisionIds: Object.freeze(revisionIds) };
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

const mapPreparedTextRange = (tr: Transaction, range: ReadyTextRange): ReadyTextRange => {
  if (tr.mapping.maps.length === 0) return range;
  return Object.freeze({
    type: "ready" as const,
    steps: Object.freeze(
      range.steps.map((step) => {
        switch (step.type) {
          case "del":
          case "format":
            return Object.freeze({
              ...step,
              from: tr.mapping.map(step.from, 1),
              to: tr.mapping.map(step.to, -1),
            });
          case "ins":
            return Object.freeze({ ...step, at: tr.mapping.map(step.at, 1) });
          default: {
            const unreachable: never = step;
            return unreachable;
          }
        }
      }),
    ),
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
  const paragraphPropertyRevisionIds: number[] = [];
  const nextParagraphPropertyRevision = () => {
    const id = revisionId++;
    paragraphPropertyRevisionIds.push(id);
    return { id, author, date: revisionStamp.date };
  };
  const styleResolver = getDocumentStyleResolver(state);
  const numbering = getDocumentNumbering(state);
  const receiptsByIndex = new Map<number, DocxComparisonInstructionReceipt>();
  const geometry = owned.instructions.filter(
    (instruction): instruction is PreparedInstructionForPhase<"geometry"> =>
      preparedForPhase(instruction, "geometry"),
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
    receiptsByIndex.set(
      instruction.originalIndex,
      Object.freeze({
        instructionIndex: instruction.originalIndex,
        instructionType: instruction.type,
        revisionIds: Object.freeze(
          result.receipt.revisions.map(({ revisionId: appliedRevisionId }) => appliedRevisionId),
        ),
        tableGeometry: result.receipt,
      }),
    );
  }
  const sourcePhase = owned.instructions
    .filter((instruction): instruction is PreparedInstructionForPhase<"source"> =>
      preparedForPhase(instruction, "source"),
    )
    .toSorted(
      (left, right) =>
        right.schedule.position - left.schedule.position ||
        right.originalIndex - left.originalIndex,
    );
  const insertionPhase = owned.instructions
    .filter((instruction): instruction is PreparedInstructionForPhase<"insertion"> =>
      preparedForPhase(instruction, "insertion"),
    )
    .toSorted(
      (left, right) =>
        right.schedule.position - left.schedule.position ||
        left.originalIndex - right.originalIndex,
    );
  // Source edits must run before boundary insertions. Within an insertion boundary,
  // canonical order is forward because mapping with assoc +1 appends each peer.
  const ordered = [...sourcePhase, ...insertionPhase];
  for (const instruction of ordered) {
    const stepsBefore = tr.steps.length;
    const instructionRevisionIds: number[] = [];
    switch (instruction.type) {
      case "replaceText":
      case "formatText": {
        const result = applyPreflightedDocxTextRange({
          tr,
          preflight: mapPreparedTextRange(tr, instruction.range),
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
      case "insertParagraph": {
        const at = tr.mapping.map(instruction.anchor.insertionPosition, 1);
        const textRevisionId = instruction.target.text.length > 0 ? revisionId++ : null;
        const paragraphRevisionId = revisionId++;
        const paragraph = trackedParagraph(
          tr,
          instruction.anchor,
          instruction.target,
          textRevisionId,
          paragraphRevisionId,
          author,
          revisionStamp.date,
        );
        if (textRevisionId !== null) instructionRevisionIds.push(textRevisionId);
        instructionRevisionIds.push(paragraphRevisionId);
        tr.insert(at, paragraph);
        const inserted =
          tr.doc.nodeAt(at) ??
          panic("A preflighted paragraph insertion did not produce its target node");
        tr = applyBlockParagraphProperties({
          tr,
          position: at,
          node: inserted,
          properties: instruction.target.properties,
          styleResolver,
          numbering,
        }).tr;
        for (const run of instruction.target.runs) {
          applyExactDirectFormatting({
            tr,
            from: at + 1 + run.startOffset,
            to: at + 1 + run.endOffset,
            formatting: run.formatting,
            styleResolver,
          });
        }
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
        const at = tr.mapping.map(instruction.source.from);
        const node = tr.doc.nodeAt(at) ?? panic("A preflighted paragraph change lost its source");
        paragraphPropertyRevisionIds.length = 0;
        const appliedProperties = applyBlockParagraphProperties({
          tr,
          position: at,
          node,
          properties: instruction.semantic.targetProperties,
          styleResolver,
          numbering,
          revisionInfo: nextParagraphPropertyRevision,
        });
        tr = appliedProperties.tr;
        instructionRevisionIds.push(...paragraphPropertyRevisionIds);
        break;
      }
      case "moveParagraph": {
        const at = tr.mapping.map(instruction.anchor.insertionPosition, 1);
        const textRevisionId = instruction.target.text.length > 0 ? revisionId++ : null;
        const paragraphRevisionId = revisionId++;
        const paragraph = trackedParagraph(
          tr,
          instruction.anchor,
          instruction.target,
          textRevisionId,
          paragraphRevisionId,
          author,
          revisionStamp.date,
          "moveTo",
        );
        if (textRevisionId !== null) instructionRevisionIds.push(textRevisionId);
        instructionRevisionIds.push(paragraphRevisionId);
        tr.insert(at, paragraph);
        const targetProperties = applyParagraphTargetProperties({
          tr,
          position: at,
          target: instruction.target,
          revisionId,
          author,
          date: revisionStamp.date,
          styleResolver,
          numbering,
          tracked: false,
        });
        tr = targetProperties.transaction;
        for (const run of instruction.target.runs) {
          applyExactDirectFormatting({
            tr,
            from: at + 1 + run.startOffset,
            to: at + 1 + run.endOffset,
            formatting: run.formatting,
            styleResolver,
          });
        }
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
      case "splitParagraph": {
        const sourcePosition = tr.mapping.map(instruction.source.from);
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
        const firstPosition = tr.mapping.map(instruction.firstSource.from);
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
        const secondPosition = tr.mapping.map(instruction.secondSource.from);
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
      case "mergeTerminalCarrier": {
        const at = tr.mapping.map(instruction.source.from);
        const paragraphRevisionId = revisionId++;
        tr.setNodeAttribute(at, "pPrMark", {
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
    receiptsByIndex.set(
      instruction.originalIndex,
      Object.freeze({
        instructionIndex: instruction.originalIndex,
        instructionType: instruction.type,
        revisionIds: Object.freeze(instructionRevisionIds),
      }),
    );
  }
  if (tr.docChanged) {
    const revisionOwnerById = new Map<number, number>();
    for (const [instructionIndex, receipt] of receiptsByIndex) {
      for (const appliedRevisionId of receipt.revisionIds) {
        revisionOwnerById.set(appliedRevisionId, instructionIndex);
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
      const receipt =
        receiptsByIndex.get(instructionIndex) ??
        panic("A final-mark normalization named a missing comparison receipt", {
          instructionIndex,
        });
      receiptsByIndex.set(
        instructionIndex,
        Object.freeze({
          ...receipt,
          revisionIds: Object.freeze([...receipt.revisionIds, synthesizedRevisionId]),
        }),
      );
    }
    markStructuralChange(tr);
    requestDeterministicParaIds(tr, `${revisionStamp.date}:${String(revisionStamp.idSeed)}`);
  }
  const receipts = owned.instructions
    .toSorted((left, right) => left.originalIndex - right.originalIndex)
    .map(
      ({ originalIndex }) =>
        receiptsByIndex.get(originalIndex) ??
        panic("A DOCX comparison instruction completed without an execution receipt", {
          instructionIndex: originalIndex,
        }),
    );
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
      transaction: tr,
    }),
  });
};
