import { panic, Result } from "better-result";

import type { FolioDocumentStoryHandle } from "../../ai-edits/headless";
import {
  type FolioContentComparison,
  type FolioContentComparisonEvent,
  type FolioContentComparisonError,
  type FolioContentPairRelation,
  type FolioContentRangePairRelation,
  type FolioContentSeparatorRelation,
  type FolioContentComparisonSessionError,
  type FolioContentComparisonWorkSession,
  type FolioContentStructuralChange,
  type FolioContentTableReplacement,
  type FolioContentWholePairRelation,
} from "../../compare/content";
import type { FolioContentBlock } from "../../compare/content-types";
import type { TableGeometryPairing } from "./table-geometry-program";
import {
  resolvedDocxContentSnapshot,
  resolvedDocxSourceOperand,
  resolvedDocxSourceOperandBlock,
  resolvedDocxStoryHandle,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";

const RESOLVED_DOCX_STORY_PAIR_BRAND: unique symbol = Symbol("resolved-docx-story-pair");
const RESOLVED_DOCX_STORY_COMPARISON_BRAND: unique symbol = Symbol(
  "resolved-docx-story-comparison",
);
const RESOLVED_DOCX_TARGET_BLOCK_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-target-block-operand",
);
const RESOLVED_DOCX_PAIR_RANGE_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-pair-range-operand",
);
const RESOLVED_DOCX_FORMATTING_RANGE_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-formatting-range-operand",
);
const RESOLVED_DOCX_SEPARATOR_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-separator-operand",
);
const RESOLVED_DOCX_WHOLE_BLOCK_REPLACEMENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-whole-block-replacement-operand",
);
const RESOLVED_DOCX_TARGET_TABLE_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-target-table-operand",
);
const RESOLVED_DOCX_TARGET_ROW_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-target-row-operand",
);
const RESOLVED_DOCX_TARGET_COLUMN_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-target-column-operand",
);
const RESOLVED_DOCX_PAIRED_EVENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-paired-event-operand",
);
const RESOLVED_DOCX_INSERTED_EVENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-inserted-event-operand",
);
const RESOLVED_DOCX_DELETED_EVENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-deleted-event-operand",
);
const RESOLVED_DOCX_MOVE_EVENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-move-event-operand",
);
const RESOLVED_DOCX_SPLIT_EVENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-split-event-operand",
);
const RESOLVED_DOCX_MERGE_EVENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-merge-event-operand",
);
const RESOLVED_DOCX_TRAILING_DELETION_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-trailing-deletion-operand",
);
const RESOLVED_DOCX_TERMINAL_REPLACEMENT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-terminal-replacement-operand",
);

/** Two exact DOCX story projections bound before semantic comparison begins. */
export type ResolvedDocxStoryPair = {
  readonly [RESOLVED_DOCX_STORY_PAIR_BRAND]: true;
};

/** The canonical neutral comparison owned by one exact DOCX story pair. */
export type ResolvedDocxStoryComparison = {
  readonly [RESOLVED_DOCX_STORY_COMPARISON_BRAND]: true;
};

/** One exact target block owned by a specific completed story comparison. */
export type ResolvedDocxTargetBlockOperand = {
  readonly [RESOLVED_DOCX_TARGET_BLOCK_OPERAND_BRAND]: true;
};

/** One canonical paired range, including its exact source and target blocks. */
export type ResolvedDocxPairRangeOperand = {
  readonly type: "pair-range";
  readonly [RESOLVED_DOCX_PAIR_RANGE_OPERAND_BRAND]: true;
};

/** One authored-formatting range selected from a canonical whole-block relation. */
export type ResolvedDocxFormattingRangeOperand = {
  readonly [RESOLVED_DOCX_FORMATTING_RANGE_OPERAND_BRAND]: true;
};

/** The exact separator relation owned by one split or merge event. */
export type ResolvedDocxSeparatorOperand = {
  readonly [RESOLVED_DOCX_SEPARATOR_OPERAND_BRAND]: true;
};

/** A canonical base/target block pair synthesized inside one aligned container. */
export type ResolvedDocxWholeBlockReplacementOperand = {
  readonly type: "whole-block-replacement";
  readonly [RESOLVED_DOCX_WHOLE_BLOCK_REPLACEMENT_OPERAND_BRAND]: true;
};

/** One exact inserted or replacement table from the target story. */
export type ResolvedDocxTargetTableOperand = {
  readonly [RESOLVED_DOCX_TARGET_TABLE_OPERAND_BRAND]: true;
};

/** One exact inserted row from the target story. */
export type ResolvedDocxTargetRowOperand = {
  readonly [RESOLVED_DOCX_TARGET_ROW_OPERAND_BRAND]: true;
};

/** One exact inserted grid column from the target story. */
export type ResolvedDocxTargetColumnOperand = {
  readonly [RESOLVED_DOCX_TARGET_COLUMN_OPERAND_BRAND]: true;
};

/** One exact modified or formatting-only event from the canonical stream. */
export type ResolvedDocxPairedEventOperand = {
  readonly [RESOLVED_DOCX_PAIRED_EVENT_OPERAND_BRAND]: true;
};

/** One exact inserted-block event from the canonical stream. */
export type ResolvedDocxInsertedEventOperand = {
  readonly [RESOLVED_DOCX_INSERTED_EVENT_OPERAND_BRAND]: true;
};

/** One exact deleted-block event from the canonical stream. */
export type ResolvedDocxDeletedEventOperand = {
  readonly [RESOLVED_DOCX_DELETED_EVENT_OPERAND_BRAND]: true;
};

/** The exact destination occurrence of one canonical move. */
export type ResolvedDocxMoveEventOperand = {
  readonly [RESOLVED_DOCX_MOVE_EVENT_OPERAND_BRAND]: true;
};

/** One exact canonical split event. */
export type ResolvedDocxSplitEventOperand = {
  readonly [RESOLVED_DOCX_SPLIT_EVENT_OPERAND_BRAND]: true;
};

/** One exact canonical merge event. */
export type ResolvedDocxMergeEventOperand = {
  readonly [RESOLVED_DOCX_MERGE_EVENT_OPERAND_BRAND]: true;
};

/** One validated same-container trailing deletion rewrite. */
export type ResolvedDocxTrailingDeletionOperand = {
  readonly [RESOLVED_DOCX_TRAILING_DELETION_OPERAND_BRAND]: true;
};

/** One validated same-container terminal delete/insert replacement. */
export type ResolvedDocxTerminalReplacementOperand = {
  readonly [RESOLVED_DOCX_TERMINAL_REPLACEMENT_OPERAND_BRAND]: true;
};

type PairedEvent =
  | Extract<FolioContentComparisonEvent, { readonly type: "modified" }>
  | Extract<FolioContentComparisonEvent, { readonly type: "formatting" }>;
type InsertedEvent = Extract<FolioContentComparisonEvent, { readonly type: "inserted" }>;
type DeletedEvent = Extract<FolioContentComparisonEvent, { readonly type: "deleted" }>;
type MoveEvent = Extract<FolioContentComparisonEvent, { readonly type: "movedTo" }>;
type SplitEvent = Extract<FolioContentComparisonEvent, { readonly type: "split" }>;
type MergeEvent = Extract<FolioContentComparisonEvent, { readonly type: "merge" }>;

export type ResolvedDocxEventOperandPayload<Event extends FolioContentComparisonEvent> = {
  readonly event: Event;
  /** Position in the canonical comparison stream. */
  readonly sequence: number;
};

type NonEmptyReadonlyArray<Value> = readonly [Value, ...Value[]];

export type ResolvedDocxTrailingDeletionOperandPayload = {
  readonly events: NonEmptyReadonlyArray<ResolvedDocxEventOperandPayload<DeletedEvent>>;
  readonly chainStart: ResolvedDocxSourceOperand;
  readonly targetCarrier?: FolioContentBlock;
};

export type ResolvedDocxTerminalReplacementOperandPayload = {
  readonly deleted: ResolvedDocxEventOperandPayload<DeletedEvent>;
  readonly inserted: ResolvedDocxEventOperandPayload<InsertedEvent>;
};

export type ResolvedDocxStoryPairPayload = {
  readonly baseStory: FolioDocumentStoryHandle;
  readonly targetStory: FolioDocumentStoryHandle;
  readonly baseSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
};

export type ResolvedDocxStoryComparisonPayload = ResolvedDocxStoryPairPayload & {
  readonly pair: ResolvedDocxStoryPair;
  readonly comparison: FolioContentComparison;
};

const payloadByPair = new WeakMap<ResolvedDocxStoryPair, ResolvedDocxStoryPairPayload>();
const payloadByComparison = new WeakMap<
  ResolvedDocxStoryComparison,
  ResolvedDocxStoryComparisonPayload
>();

type FolioContentFormattingRange = NonNullable<
  FolioContentWholePairRelation["formatting"]
>["ranges"][number];

type ResolvedDocxStoryComparisonIndex = {
  readonly eventSequence: ReadonlyMap<FolioContentComparisonEvent, number>;
  readonly relations: ReadonlySet<FolioContentPairRelation>;
  readonly separators: ReadonlySet<FolioContentSeparatorRelation>;
  readonly structuralChanges: ReadonlySet<FolioContentStructuralChange>;
  readonly tableReplacements: ReadonlySet<FolioContentTableReplacement>;
  readonly baseContainerByBlock: ReadonlyMap<FolioContentBlock, object>;
  readonly targetContainerByBlock: ReadonlyMap<FolioContentBlock, object>;
  readonly baseTableIndexesByTargetTableIndex: ReadonlyMap<number, ReadonlySet<number>>;
  readonly tableGeometryPairings: readonly TableGeometryPairing[];
};

type ComparisonOwnedPayload<Payload> = {
  readonly comparison: ResolvedDocxStoryComparison;
  readonly value: Payload;
};

type WholeBlockReplacementPayload = {
  readonly baseBlock: FolioContentBlock;
  readonly targetBlock: FolioContentBlock;
};

type TargetTableOwner =
  | Extract<FolioContentStructuralChange, { readonly type: "table-insert" }>
  | FolioContentTableReplacement;

const structuralChangeSnapshot = (
  change: FolioContentStructuralChange,
  payload: ResolvedDocxStoryComparisonPayload,
): ResolvedDocxStorySnapshot => {
  switch (change.type) {
    case "table-insert":
    case "table-row-insert":
    case "table-column-insert":
      return payload.targetSnapshot;
    case "table-delete":
    case "table-row-delete":
    case "table-column-delete":
      return payload.baseSnapshot;
    default: {
      const unreachable: never = change;
      return panic("Unhandled structural change while binding DOCX operands", {
        change: unreachable,
      });
    }
  }
};

const indexByComparison = new WeakMap<
  ResolvedDocxStoryComparison,
  ResolvedDocxStoryComparisonIndex
>();
const payloadByTargetBlockOperand = new WeakMap<
  ResolvedDocxTargetBlockOperand,
  ComparisonOwnedPayload<FolioContentBlock>
>();
const payloadByPairRangeOperand = new WeakMap<
  ResolvedDocxPairRangeOperand,
  ComparisonOwnedPayload<FolioContentWholePairRelation | FolioContentRangePairRelation>
>();
const payloadByFormattingRangeOperand = new WeakMap<
  ResolvedDocxFormattingRangeOperand,
  ComparisonOwnedPayload<{
    readonly relation: FolioContentWholePairRelation;
    readonly range: FolioContentFormattingRange;
  }>
>();
const payloadBySeparatorOperand = new WeakMap<
  ResolvedDocxSeparatorOperand,
  ComparisonOwnedPayload<FolioContentSeparatorRelation>
>();
const payloadByWholeBlockReplacementOperand = new WeakMap<
  ResolvedDocxWholeBlockReplacementOperand,
  ComparisonOwnedPayload<WholeBlockReplacementPayload>
>();
const payloadByTargetTableOperand = new WeakMap<
  ResolvedDocxTargetTableOperand,
  ComparisonOwnedPayload<TargetTableOwner>
>();
const payloadByTargetRowOperand = new WeakMap<
  ResolvedDocxTargetRowOperand,
  ComparisonOwnedPayload<
    Extract<FolioContentStructuralChange, { readonly type: "table-row-insert" }>
  >
>();
const payloadByTargetColumnOperand = new WeakMap<
  ResolvedDocxTargetColumnOperand,
  ComparisonOwnedPayload<
    Extract<FolioContentStructuralChange, { readonly type: "table-column-insert" }>
  >
>();
const payloadByPairedEventOperand = new WeakMap<
  ResolvedDocxPairedEventOperand,
  ComparisonOwnedPayload<ResolvedDocxEventOperandPayload<PairedEvent>>
>();
const payloadByInsertedEventOperand = new WeakMap<
  ResolvedDocxInsertedEventOperand,
  ComparisonOwnedPayload<ResolvedDocxEventOperandPayload<InsertedEvent>>
>();
const payloadByDeletedEventOperand = new WeakMap<
  ResolvedDocxDeletedEventOperand,
  ComparisonOwnedPayload<ResolvedDocxEventOperandPayload<DeletedEvent>>
>();
const payloadByMoveEventOperand = new WeakMap<
  ResolvedDocxMoveEventOperand,
  ComparisonOwnedPayload<ResolvedDocxEventOperandPayload<MoveEvent>>
>();
const payloadBySplitEventOperand = new WeakMap<
  ResolvedDocxSplitEventOperand,
  ComparisonOwnedPayload<ResolvedDocxEventOperandPayload<SplitEvent>>
>();
const payloadByMergeEventOperand = new WeakMap<
  ResolvedDocxMergeEventOperand,
  ComparisonOwnedPayload<ResolvedDocxEventOperandPayload<MergeEvent>>
>();
const payloadByTrailingDeletionOperand = new WeakMap<
  ResolvedDocxTrailingDeletionOperand,
  ComparisonOwnedPayload<ResolvedDocxTrailingDeletionOperandPayload>
>();
const payloadByTerminalReplacementOperand = new WeakMap<
  ResolvedDocxTerminalReplacementOperand,
  ComparisonOwnedPayload<ResolvedDocxTerminalReplacementOperandPayload>
>();

const comparisonIndexOf = (
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxStoryComparisonIndex =>
  indexByComparison.get(comparison) ??
  panic("A resolved DOCX story comparison was not created by Folio");

const requireComparisonOwned = <Payload>(
  payload: ComparisonOwnedPayload<Payload> | undefined,
  comparison: ResolvedDocxStoryComparison,
  name: string,
): Payload => {
  if (!payload) return panic(`A ${name} was not created by Folio`);
  if (payload.comparison !== comparison) {
    return panic(`A ${name} belongs to another story comparison`);
  }
  return payload.value;
};

const createComparisonIndex = (
  payload: ResolvedDocxStoryComparisonPayload,
): ResolvedDocxStoryComparisonIndex => {
  const relations = new Set<FolioContentPairRelation>();
  const eventSequence = new Map<FolioContentComparisonEvent, number>();
  const separators = new Set<FolioContentSeparatorRelation>();
  const structuralChanges = new Set<FolioContentStructuralChange>();
  const tableReplacements = new Set<FolioContentTableReplacement>();
  const baseContainerByBlock = new Map<FolioContentBlock, object>();
  const targetContainerByBlock = new Map<FolioContentBlock, object>();
  const baseTableIndexesByTargetTableIndex = new Map<number, Set<number>>();
  const tableGeometryPairings: TableGeometryPairing[] = [];
  const pairedBaseCells = new Set<string>();
  const registerRelation = (relation: FolioContentPairRelation): void => {
    if (relations.has(relation)) return;
    resolvedDocxSourceOperand(payload.baseSnapshot, relation.base.block);
    resolvedDocxSourceOperand(payload.targetSnapshot, relation.revised.block);
    relations.add(relation);
    baseContainerByBlock.set(relation.base.block, relation.baseContainerAlignment);
    targetContainerByBlock.set(relation.revised.block, relation.revisedContainerAlignment);
    if (relation.relationType === "separator") separators.add(relation);
    const baseTable = relation.base.block.table;
    const targetTable = relation.revised.block.table;
    if (!baseTable || !targetTable) return;
    const baseTableIndexes =
      baseTableIndexesByTargetTableIndex.get(targetTable.tableIndex) ?? new Set<number>();
    baseTableIndexes.add(baseTable.tableIndex);
    baseTableIndexesByTargetTableIndex.set(targetTable.tableIndex, baseTableIndexes);
    if (pairedBaseCells.has(baseTable.cellIdentity.id)) return;
    pairedBaseCells.add(baseTable.cellIdentity.id);
    tableGeometryPairings.push(
      Object.freeze({
        base: Object.freeze({
          tableIndex: baseTable.tableIndex,
          rowIndex: baseTable.rowIndex,
          cellIndex: baseTable.cellIndex,
        }),
        target: Object.freeze({
          tableIndex: targetTable.tableIndex,
          rowIndex: targetTable.rowIndex,
          cellIndex: targetTable.cellIndex,
        }),
      }),
    );
  };
  for (const [sequence, event] of payload.comparison.events.entries()) {
    eventSequence.set(event, sequence);
    switch (event.type) {
      case "unchanged":
      case "modified":
      case "formatting":
        registerRelation(event.relation);
        break;
      case "movedFrom":
      case "movedTo":
        registerRelation(event.move.relation);
        break;
      case "split":
      case "merge":
        registerRelation(event.relations[0]);
        registerRelation(event.relations[1]);
        registerRelation(event.separator);
        break;
      case "inserted":
        resolvedDocxSourceOperand(payload.targetSnapshot, event.block);
        targetContainerByBlock.set(event.block, event.containerAlignment);
        break;
      case "deleted":
        resolvedDocxSourceOperand(payload.baseSnapshot, event.block);
        baseContainerByBlock.set(event.block, event.containerAlignment);
        break;
      case "tableReplacement":
        tableReplacements.add(event.replacement);
        for (const block of event.replacement.baseBlocks) {
          resolvedDocxSourceOperand(payload.baseSnapshot, block);
        }
        for (const block of event.replacement.revisedBlocks) {
          resolvedDocxSourceOperand(payload.targetSnapshot, block);
        }
        break;
      case "structural": {
        if (structuralChanges.has(event.change)) break;
        structuralChanges.add(event.change);
        const snapshot = structuralChangeSnapshot(event.change, payload);
        for (const block of event.change.blocks) {
          resolvedDocxSourceOperand(snapshot, block);
        }
        break;
      }
      default: {
        const unreachable: never = event;
        return panic("Unhandled comparison event while binding DOCX operands", {
          event: unreachable,
        });
      }
    }
  }
  return Object.freeze({
    eventSequence,
    relations,
    separators,
    structuralChanges,
    tableReplacements,
    baseContainerByBlock,
    targetContainerByBlock,
    baseTableIndexesByTargetTableIndex,
    tableGeometryPairings: Object.freeze(tableGeometryPairings),
  });
};

/** Bind two genuine snapshots before they can enter comparison or transport lowering. */
export const createResolvedDocxStoryPair = ({
  baseSnapshot,
  targetSnapshot,
}: {
  readonly baseSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
}): ResolvedDocxStoryPair => {
  const baseStory = resolvedDocxStoryHandle(baseSnapshot);
  const targetStory = resolvedDocxStoryHandle(targetSnapshot);
  if (baseStory.type !== targetStory.type) {
    return panic("A resolved DOCX story pair must contain compatible story kinds", {
      baseKind: baseStory.type,
      targetKind: targetStory.type,
    });
  }
  const payload = Object.freeze({
    baseStory,
    targetStory,
    baseSnapshot,
    targetSnapshot,
  });
  const pair = Object.freeze({ [RESOLVED_DOCX_STORY_PAIR_BRAND]: true as const });
  payloadByPair.set(pair, payload);
  return pair;
};

/** Resolve only a pair issued by the factory in this module. */
export const resolvedDocxStoryPairPayload = (
  pair: ResolvedDocxStoryPair,
): ResolvedDocxStoryPairPayload =>
  payloadByPair.get(pair) ?? panic("A resolved DOCX story pair was not created by Folio");

/**
 * Run the neutral engine for one exact pair. The returned capsule is the only
 * value the DOCX planner accepts, so semantic output cannot be joined later to
 * structurally similar or independently supplied snapshots.
 */
export const compareResolvedDocxStoryPair = ({
  pair,
  workSession,
}: {
  readonly pair: ResolvedDocxStoryPair;
  readonly workSession: FolioContentComparisonWorkSession;
}): Result<
  ResolvedDocxStoryComparison,
  FolioContentComparisonError | FolioContentComparisonSessionError
> => {
  const pairPayload = resolvedDocxStoryPairPayload(pair);
  const captured = workSession.captureComparison({
    base: resolvedDocxContentSnapshot(pairPayload.baseSnapshot),
    revised: resolvedDocxContentSnapshot(pairPayload.targetSnapshot),
  });
  if (captured.isErr()) return Result.err(captured.error);
  const compared = captured.value.compare();
  if (compared.isErr()) return Result.err(compared.error);

  const comparison = Object.freeze({
    [RESOLVED_DOCX_STORY_COMPARISON_BRAND]: true as const,
  });
  payloadByComparison.set(
    comparison,
    Object.freeze({ ...pairPayload, pair, comparison: compared.value }),
  );
  indexByComparison.set(
    comparison,
    createComparisonIndex(
      payloadByComparison.get(comparison) ??
        panic("A resolved DOCX story comparison lost its canonical payload"),
    ),
  );
  return Result.ok(comparison);
};

/** Resolve only semantic output issued for a genuine exact story pair. */
export const resolvedDocxStoryComparisonPayload = (
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxStoryComparisonPayload =>
  payloadByComparison.get(comparison) ??
  panic("A resolved DOCX story comparison was not created by Folio");

const canonicalEventPayload = <Event extends FolioContentComparisonEvent>(
  comparison: ResolvedDocxStoryComparison,
  event: Event,
): ResolvedDocxEventOperandPayload<Event> => {
  const sequence = comparisonIndexOf(comparison).eventSequence.get(event);
  if (sequence === undefined) {
    return panic("A DOCX event operand must name an exact canonical event");
  }
  return Object.freeze({ event, sequence });
};

/** Issue an opaque operand for one exact modified or formatting-only event. */
export const resolvedDocxPairedEventOperand = (
  comparison: ResolvedDocxStoryComparison,
  event: PairedEvent,
): ResolvedDocxPairedEventOperand => {
  if (event.type !== "modified" && event.type !== "formatting") {
    return panic("A DOCX paired-event operand must name a paired change event");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_PAIRED_EVENT_OPERAND_BRAND]: true as const });
  payloadByPairedEventOperand.set(operand, {
    comparison,
    value: canonicalEventPayload(comparison, event),
  });
  return operand;
};

/** Resolve a paired event only inside the comparison that issued it. */
export const resolvedDocxPairedEventOperandPayload = (
  operand: ResolvedDocxPairedEventOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxEventOperandPayload<PairedEvent> =>
  requireComparisonOwned(
    payloadByPairedEventOperand.get(operand),
    comparison,
    "DOCX paired-event operand",
  );

/** Issue an opaque operand for one exact inserted-block event. */
export const resolvedDocxInsertedEventOperand = (
  comparison: ResolvedDocxStoryComparison,
  event: InsertedEvent,
): ResolvedDocxInsertedEventOperand => {
  if (event.type !== "inserted") {
    return panic("A DOCX inserted-event operand must name an inserted event");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_INSERTED_EVENT_OPERAND_BRAND]: true as const });
  payloadByInsertedEventOperand.set(operand, {
    comparison,
    value: canonicalEventPayload(comparison, event),
  });
  return operand;
};

/** Resolve an inserted event only inside the comparison that issued it. */
export const resolvedDocxInsertedEventOperandPayload = (
  operand: ResolvedDocxInsertedEventOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxEventOperandPayload<InsertedEvent> =>
  requireComparisonOwned(
    payloadByInsertedEventOperand.get(operand),
    comparison,
    "DOCX inserted-event operand",
  );

/** Issue an opaque operand for one exact deleted-block event. */
export const resolvedDocxDeletedEventOperand = (
  comparison: ResolvedDocxStoryComparison,
  event: DeletedEvent,
): ResolvedDocxDeletedEventOperand => {
  if (event.type !== "deleted") {
    return panic("A DOCX deleted-event operand must name a deleted event");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_DELETED_EVENT_OPERAND_BRAND]: true as const });
  payloadByDeletedEventOperand.set(operand, {
    comparison,
    value: canonicalEventPayload(comparison, event),
  });
  return operand;
};

/** Resolve a deleted event only inside the comparison that issued it. */
export const resolvedDocxDeletedEventOperandPayload = (
  operand: ResolvedDocxDeletedEventOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxEventOperandPayload<DeletedEvent> =>
  requireComparisonOwned(
    payloadByDeletedEventOperand.get(operand),
    comparison,
    "DOCX deleted-event operand",
  );

/** Issue an opaque operand for the destination occurrence of one exact move. */
export const resolvedDocxMoveEventOperand = (
  comparison: ResolvedDocxStoryComparison,
  event: MoveEvent,
): ResolvedDocxMoveEventOperand => {
  if (event.type !== "movedTo") {
    return panic("A DOCX move-event operand must name a move destination event");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_MOVE_EVENT_OPERAND_BRAND]: true as const });
  payloadByMoveEventOperand.set(operand, {
    comparison,
    value: canonicalEventPayload(comparison, event),
  });
  return operand;
};

/** Resolve a move event only inside the comparison that issued it. */
export const resolvedDocxMoveEventOperandPayload = (
  operand: ResolvedDocxMoveEventOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxEventOperandPayload<MoveEvent> =>
  requireComparisonOwned(
    payloadByMoveEventOperand.get(operand),
    comparison,
    "DOCX move-event operand",
  );

/** Issue an opaque operand for one exact split event. */
export const resolvedDocxSplitEventOperand = (
  comparison: ResolvedDocxStoryComparison,
  event: SplitEvent,
): ResolvedDocxSplitEventOperand => {
  if (event.type !== "split") {
    return panic("A DOCX split-event operand must name a split event");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_SPLIT_EVENT_OPERAND_BRAND]: true as const });
  payloadBySplitEventOperand.set(operand, {
    comparison,
    value: canonicalEventPayload(comparison, event),
  });
  return operand;
};

/** Resolve a split event only inside the comparison that issued it. */
export const resolvedDocxSplitEventOperandPayload = (
  operand: ResolvedDocxSplitEventOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxEventOperandPayload<SplitEvent> =>
  requireComparisonOwned(
    payloadBySplitEventOperand.get(operand),
    comparison,
    "DOCX split-event operand",
  );

/** Issue an opaque operand for one exact merge event. */
export const resolvedDocxMergeEventOperand = (
  comparison: ResolvedDocxStoryComparison,
  event: MergeEvent,
): ResolvedDocxMergeEventOperand => {
  if (event.type !== "merge") {
    return panic("A DOCX merge-event operand must name a merge event");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_MERGE_EVENT_OPERAND_BRAND]: true as const });
  payloadByMergeEventOperand.set(operand, {
    comparison,
    value: canonicalEventPayload(comparison, event),
  });
  return operand;
};

/** Resolve a merge event only inside the comparison that issued it. */
export const resolvedDocxMergeEventOperandPayload = (
  operand: ResolvedDocxMergeEventOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxEventOperandPayload<MergeEvent> =>
  requireComparisonOwned(
    payloadByMergeEventOperand.get(operand),
    comparison,
    "DOCX merge-event operand",
  );

/** Issue one opaque operand for an exact canonical target block. */
export const resolvedDocxTargetBlockOperand = (
  comparison: ResolvedDocxStoryComparison,
  block: FolioContentBlock,
): ResolvedDocxTargetBlockOperand => {
  const { targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
  resolvedDocxSourceOperand(targetSnapshot, block);
  const operand = Object.freeze({ [RESOLVED_DOCX_TARGET_BLOCK_OPERAND_BRAND]: true as const });
  payloadByTargetBlockOperand.set(operand, { comparison, value: block });
  return operand;
};

/** Resolve a target block only inside the comparison that issued its operand. */
export const resolvedDocxTargetBlockOperandBlock = (
  operand: ResolvedDocxTargetBlockOperand,
  comparison: ResolvedDocxStoryComparison,
): FolioContentBlock =>
  requireComparisonOwned(
    payloadByTargetBlockOperand.get(operand),
    comparison,
    "DOCX target-block operand",
  );

/** Bind a non-empty trailing deletion run and its retained mark carrier. */
export const resolvedDocxTrailingDeletionOperand = (
  comparison: ResolvedDocxStoryComparison,
  {
    events,
    chainStart,
    targetCarrier,
  }: {
    readonly events: NonEmptyReadonlyArray<ResolvedDocxDeletedEventOperand>;
    readonly chainStart: ResolvedDocxSourceOperand;
    readonly targetCarrier?: ResolvedDocxTargetBlockOperand;
  },
): ResolvedDocxTrailingDeletionOperand => {
  const { baseSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
  const index = comparisonIndexOf(comparison);
  const firstEvent = resolvedDocxDeletedEventOperandPayload(events[0], comparison);
  const remainingEvents = events
    .slice(1)
    .map((event) => resolvedDocxDeletedEventOperandPayload(event, comparison));
  const ownedEvents: NonEmptyReadonlyArray<ResolvedDocxEventOperandPayload<DeletedEvent>> = [
    firstEvent,
    ...remainingEvents,
  ];
  const container = index.baseContainerByBlock.get(firstEvent.event.block);
  const chainStartBlock = resolvedDocxSourceOperandBlock(chainStart, baseSnapshot);
  if (
    !container ||
    index.baseContainerByBlock.get(chainStartBlock) !== container ||
    ownedEvents.some(({ event }) => index.baseContainerByBlock.get(event.block) !== container) ||
    ownedEvents.some(({ event }) => event.block === chainStartBlock)
  ) {
    return panic("A trailing DOCX deletion operand must remain in one canonical container");
  }
  const targetBlock = targetCarrier
    ? resolvedDocxTargetBlockOperandBlock(targetCarrier, comparison)
    : undefined;
  if (targetBlock && index.targetContainerByBlock.get(targetBlock) !== container) {
    return panic("A trailing DOCX deletion target carrier belongs to another container");
  }
  const operand = Object.freeze({
    [RESOLVED_DOCX_TRAILING_DELETION_OPERAND_BRAND]: true as const,
  });
  payloadByTrailingDeletionOperand.set(operand, {
    comparison,
    value: Object.freeze({
      events: Object.freeze(ownedEvents),
      chainStart,
      ...(targetBlock && { targetCarrier: targetBlock }),
    }),
  });
  return operand;
};

/** Resolve a trailing deletion only inside the comparison that issued it. */
export const resolvedDocxTrailingDeletionOperandPayload = (
  operand: ResolvedDocxTrailingDeletionOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxTrailingDeletionOperandPayload =>
  requireComparisonOwned(
    payloadByTrailingDeletionOperand.get(operand),
    comparison,
    "DOCX trailing-deletion operand",
  );

/** Bind the exact delete/insert pair rewritten onto one terminal paragraph. */
export const resolvedDocxTerminalReplacementOperand = (
  comparison: ResolvedDocxStoryComparison,
  {
    deleted,
    inserted,
  }: {
    readonly deleted: ResolvedDocxDeletedEventOperand;
    readonly inserted: ResolvedDocxInsertedEventOperand;
  },
): ResolvedDocxTerminalReplacementOperand => {
  const deletedPayload = resolvedDocxDeletedEventOperandPayload(deleted, comparison);
  const insertedPayload = resolvedDocxInsertedEventOperandPayload(inserted, comparison);
  const index = comparisonIndexOf(comparison);
  const baseContainer = index.baseContainerByBlock.get(deletedPayload.event.block);
  const targetContainer = index.targetContainerByBlock.get(insertedPayload.event.block);
  if (!baseContainer || baseContainer !== targetContainer) {
    return panic("A terminal DOCX replacement must remain in one canonical container");
  }
  const operand = Object.freeze({
    [RESOLVED_DOCX_TERMINAL_REPLACEMENT_OPERAND_BRAND]: true as const,
  });
  payloadByTerminalReplacementOperand.set(operand, {
    comparison,
    value: Object.freeze({ deleted: deletedPayload, inserted: insertedPayload }),
  });
  return operand;
};

/** Resolve a terminal replacement only inside the comparison that issued it. */
export const resolvedDocxTerminalReplacementOperandPayload = (
  operand: ResolvedDocxTerminalReplacementOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxTerminalReplacementOperandPayload =>
  requireComparisonOwned(
    payloadByTerminalReplacementOperand.get(operand),
    comparison,
    "DOCX terminal-replacement operand",
  );

/** Issue one opaque operand for an exact whole or split/merge range relation. */
export const resolvedDocxPairRangeOperand = (
  comparison: ResolvedDocxStoryComparison,
  relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
): ResolvedDocxPairRangeOperand => {
  if (!comparisonIndexOf(comparison).relations.has(relation)) {
    return panic("A DOCX pair-range operand must name an exact canonical relation");
  }
  const operand = Object.freeze({
    type: "pair-range" as const,
    [RESOLVED_DOCX_PAIR_RANGE_OPERAND_BRAND]: true as const,
  });
  payloadByPairRangeOperand.set(operand, { comparison, value: relation });
  return operand;
};

/** Resolve a paired range only inside the comparison that issued it. */
export const resolvedDocxPairRangeOperandRelation = (
  operand: ResolvedDocxPairRangeOperand,
  comparison: ResolvedDocxStoryComparison,
): FolioContentWholePairRelation | FolioContentRangePairRelation =>
  requireComparisonOwned(
    payloadByPairRangeOperand.get(operand),
    comparison,
    "DOCX pair-range operand",
  );

/** Issue one exact authored-formatting range from a canonical whole relation. */
export const resolvedDocxFormattingRangeOperand = (
  comparison: ResolvedDocxStoryComparison,
  relation: FolioContentWholePairRelation,
  range: FolioContentFormattingRange,
): ResolvedDocxFormattingRangeOperand => {
  if (
    !comparisonIndexOf(comparison).relations.has(relation) ||
    !relation.formatting?.ranges.includes(range)
  ) {
    return panic("A DOCX formatting-range operand must name an exact canonical range");
  }
  const operand = Object.freeze({
    [RESOLVED_DOCX_FORMATTING_RANGE_OPERAND_BRAND]: true as const,
  });
  payloadByFormattingRangeOperand.set(operand, {
    comparison,
    value: Object.freeze({ relation, range }),
  });
  return operand;
};

/** Resolve an authored-formatting range only inside its issuing comparison. */
export const resolvedDocxFormattingRangeOperandPayload = (
  operand: ResolvedDocxFormattingRangeOperand,
  comparison: ResolvedDocxStoryComparison,
): {
  readonly relation: FolioContentWholePairRelation;
  readonly range: FolioContentFormattingRange;
} =>
  requireComparisonOwned(
    payloadByFormattingRangeOperand.get(operand),
    comparison,
    "DOCX formatting-range operand",
  );

/** Issue one exact split/merge separator relation. */
export const resolvedDocxSeparatorOperand = (
  comparison: ResolvedDocxStoryComparison,
  relation: FolioContentSeparatorRelation,
): ResolvedDocxSeparatorOperand => {
  if (!comparisonIndexOf(comparison).separators.has(relation)) {
    return panic("A DOCX separator operand must name an exact canonical relation");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_SEPARATOR_OPERAND_BRAND]: true as const });
  payloadBySeparatorOperand.set(operand, { comparison, value: relation });
  return operand;
};

/** Resolve a separator relation only inside its issuing comparison. */
export const resolvedDocxSeparatorOperandRelation = (
  operand: ResolvedDocxSeparatorOperand,
  comparison: ResolvedDocxStoryComparison,
): FolioContentSeparatorRelation =>
  requireComparisonOwned(
    payloadBySeparatorOperand.get(operand),
    comparison,
    "DOCX separator operand",
  );

/**
 * Issue the synthetic full-range pairing used when a terminal deleted carrier
 * becomes the last inserted paragraph in the same canonical container.
 */
export const resolvedDocxWholeBlockReplacementOperand = (
  comparison: ResolvedDocxStoryComparison,
  baseBlock: FolioContentBlock,
  targetBlock: FolioContentBlock,
): ResolvedDocxWholeBlockReplacementOperand => {
  const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
  resolvedDocxSourceOperand(baseSnapshot, baseBlock);
  resolvedDocxSourceOperand(targetSnapshot, targetBlock);
  const index = comparisonIndexOf(comparison);
  const baseContainer = index.baseContainerByBlock.get(baseBlock);
  const targetContainer = index.targetContainerByBlock.get(targetBlock);
  if (!baseContainer || baseContainer !== targetContainer) {
    return panic("A DOCX whole-block replacement must remain in one canonical container");
  }
  const operand = Object.freeze({
    type: "whole-block-replacement" as const,
    [RESOLVED_DOCX_WHOLE_BLOCK_REPLACEMENT_OPERAND_BRAND]: true as const,
  });
  payloadByWholeBlockReplacementOperand.set(operand, {
    comparison,
    value: Object.freeze({ baseBlock, targetBlock }),
  });
  return operand;
};

/** Discriminate the two replacement-range operand kinds without trusting shape. */
export const resolvedDocxReplacementRangeOperandPayload = (
  operand: ResolvedDocxPairRangeOperand | ResolvedDocxWholeBlockReplacementOperand,
  comparison: ResolvedDocxStoryComparison,
):
  | {
      readonly type: "pair";
      readonly relation: FolioContentWholePairRelation | FolioContentRangePairRelation;
    }
  | ({ readonly type: "whole-block" } & WholeBlockReplacementPayload) => {
  switch (operand.type) {
    case "pair-range":
      return Object.freeze({
        type: "pair",
        relation: requireComparisonOwned(
          payloadByPairRangeOperand.get(operand),
          comparison,
          "DOCX pair-range operand",
        ),
      });
    case "whole-block-replacement": {
      const whole = requireComparisonOwned(
        payloadByWholeBlockReplacementOperand.get(operand),
        comparison,
        "DOCX whole-block replacement operand",
      );
      return Object.freeze({ type: "whole-block", ...whole });
    }
    default: {
      const unreachable: never = operand;
      return panic("Unhandled DOCX replacement-range operand", { operand: unreachable });
    }
  }
};

/** Issue one exact inserted or replacement table from the target story. */
export const resolvedDocxTargetTableOperand = (
  comparison: ResolvedDocxStoryComparison,
  owner: TargetTableOwner,
): ResolvedDocxTargetTableOperand => {
  const index = comparisonIndexOf(comparison);
  const exact =
    "type" in owner
      ? index.structuralChanges.has(owner) && owner.type === "table-insert"
      : index.tableReplacements.has(owner);
  if (!exact) {
    return panic("A DOCX target-table operand must name an exact canonical table change");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_TARGET_TABLE_OPERAND_BRAND]: true as const });
  payloadByTargetTableOperand.set(operand, { comparison, value: owner });
  return operand;
};

/** Resolve a target table only inside its issuing comparison. */
export const resolvedDocxTargetTableOperandOwner = (
  operand: ResolvedDocxTargetTableOperand,
  comparison: ResolvedDocxStoryComparison,
): TargetTableOwner =>
  requireComparisonOwned(
    payloadByTargetTableOperand.get(operand),
    comparison,
    "DOCX target-table operand",
  );

/** Issue one exact inserted target row. */
export const resolvedDocxTargetRowOperand = (
  comparison: ResolvedDocxStoryComparison,
  change: Extract<FolioContentStructuralChange, { readonly type: "table-row-insert" }>,
): ResolvedDocxTargetRowOperand => {
  if (!comparisonIndexOf(comparison).structuralChanges.has(change)) {
    return panic("A DOCX target-row operand must name an exact canonical row change");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_TARGET_ROW_OPERAND_BRAND]: true as const });
  payloadByTargetRowOperand.set(operand, { comparison, value: change });
  return operand;
};

/** Resolve a target row only inside its issuing comparison. */
export const resolvedDocxTargetRowOperandChange = (
  operand: ResolvedDocxTargetRowOperand,
  comparison: ResolvedDocxStoryComparison,
): Extract<FolioContentStructuralChange, { readonly type: "table-row-insert" }> =>
  requireComparisonOwned(
    payloadByTargetRowOperand.get(operand),
    comparison,
    "DOCX target-row operand",
  );

/** Issue one exact inserted target column. */
export const resolvedDocxTargetColumnOperand = (
  comparison: ResolvedDocxStoryComparison,
  change: Extract<FolioContentStructuralChange, { readonly type: "table-column-insert" }>,
): ResolvedDocxTargetColumnOperand => {
  if (!comparisonIndexOf(comparison).structuralChanges.has(change)) {
    return panic("A DOCX target-column operand must name an exact canonical column change");
  }
  const operand = Object.freeze({ [RESOLVED_DOCX_TARGET_COLUMN_OPERAND_BRAND]: true as const });
  payloadByTargetColumnOperand.set(operand, { comparison, value: change });
  return operand;
};

/** Resolve a target column only inside its issuing comparison. */
export const resolvedDocxTargetColumnOperandChange = (
  operand: ResolvedDocxTargetColumnOperand,
  comparison: ResolvedDocxStoryComparison,
): Extract<FolioContentStructuralChange, { readonly type: "table-column-insert" }> =>
  requireComparisonOwned(
    payloadByTargetColumnOperand.get(operand),
    comparison,
    "DOCX target-column operand",
  );

const EMPTY_TABLE_INDEX_SET: ReadonlySet<number> = new Set<number>();

/** Base table indices canonically paired with one target table in this comparison. */
export const resolvedDocxPairedBaseTableIndexes = (
  comparison: ResolvedDocxStoryComparison,
  targetTableIndex: number,
): ReadonlySet<number> =>
  comparisonIndexOf(comparison).baseTableIndexesByTargetTableIndex.get(targetTableIndex) ??
  EMPTY_TABLE_INDEX_SET;

/** Exact cell pairings derived once from this comparison's canonical relation graph. */
export const resolvedDocxTableGeometryPairings = (
  comparison: ResolvedDocxStoryComparison,
): readonly TableGeometryPairing[] => comparisonIndexOf(comparison).tableGeometryPairings;
