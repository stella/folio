import { panic, Result } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { FolioDocumentStoryHandle } from "../../ai-edits/headless";
import { storyTablesOf } from "../../ai-edits/snapshot";
import {
  type FolioContentComparison,
  type FolioContentBlockGroup,
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
import type {
  FolioContentBaseContainerAlignment,
  FolioContentBlock,
  FolioContentRevisedContainerAlignment,
} from "../../compare/content-types";
import { canonicalJson } from "../../utils/canonicalJson";
import {
  preflightTableGeometry,
  tableGeometryProgramSemanticChanges,
  type TableGeometryPairing,
  type TableGeometryProgram,
  type TableGeometrySemanticChange,
  type TableGeometryUnsupportedIssue,
} from "./table-geometry-program";
import {
  resolvedDocxContentSnapshot,
  resolvedDocxContentBlocks,
  resolvedDocxOperationSnapshot,
  resolvedDocxSourceOperand,
  resolvedDocxSourceOperandBlock,
  resolvedDocxStoryHandle,
  resolvedDocxTableNodes,
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
const RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-table-structure-operand",
);
const RESOLVED_DOCX_TABLE_FORMAT_OPERAND_BRAND: unique symbol = Symbol(
  "resolved-docx-table-format-operand",
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

/**
 * One exact structural obligation issued by a completed comparison. Its
 * semantic owner stays private, so compilation cannot split coordinates from
 * the relation that proved them.
 */
export type ResolvedDocxTableStructureOperand =
  | {
      readonly type: "insertTable";
      readonly [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true;
    }
  | {
      readonly type: "deleteTable";
      readonly [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true;
    }
  | {
      readonly type: "replaceTable";
      readonly [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true;
    }
  | {
      readonly type: "insertTableRow";
      readonly [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true;
    }
  | {
      readonly type: "deleteTableRow";
      readonly [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true;
    }
  | {
      readonly type: "insertTableColumn";
      readonly [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true;
    }
  | {
      readonly type: "deleteTableColumn";
      readonly [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true;
    };

/** One exact set of changed table properties issued by a completed comparison. */
export type ResolvedDocxTableFormatOperand = {
  readonly type: "matchTableFormatting";
  readonly [RESOLVED_DOCX_TABLE_FORMAT_OPERAND_BRAND]: true;
};

type StructuralInsertionBoundary = {
  readonly source: ResolvedDocxSourceOperand;
  readonly position: "after" | "before";
};

type ResolvedDocxTerminalTableCarrier = {
  readonly source: ResolvedDocxSourceOperand;
  readonly event: Extract<FolioContentComparisonEvent, { readonly type: "deleted" }>;
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

export type ResolvedDocxTableReplacementOwner =
  | {
      readonly type: "canonical-replacement";
      readonly replacement: FolioContentTableReplacement;
    }
  | {
      readonly type: "structural-pair";
      readonly deleted: Extract<FolioContentStructuralChange, { readonly type: "table-delete" }>;
      readonly inserted: Extract<FolioContentStructuralChange, { readonly type: "table-insert" }>;
    };

export type ResolvedDocxTableStructureOperandPayload =
  | {
      readonly type: "insertTable";
      readonly anchor: StructuralInsertionBoundary;
      readonly change: Extract<FolioContentStructuralChange, { readonly type: "table-insert" }>;
      readonly terminalCarrier?: ResolvedDocxTerminalTableCarrier;
    }
  | {
      readonly type: "deleteTable";
      readonly source: ResolvedDocxSourceOperand;
      readonly change: Extract<FolioContentStructuralChange, { readonly type: "table-delete" }>;
    }
  | {
      readonly type: "replaceTable";
      readonly source: ResolvedDocxSourceOperand;
      readonly owner: ResolvedDocxTableReplacementOwner;
      readonly terminalCarrier?: ResolvedDocxTerminalTableCarrier;
    }
  | {
      readonly type: "insertTableRow";
      readonly anchor: StructuralInsertionBoundary;
      readonly change: Extract<FolioContentStructuralChange, { readonly type: "table-row-insert" }>;
    }
  | {
      readonly type: "deleteTableRow";
      readonly source: ResolvedDocxSourceOperand;
      readonly change: Extract<FolioContentStructuralChange, { readonly type: "table-row-delete" }>;
    }
  | {
      readonly type: "insertTableColumn";
      readonly anchor: StructuralInsertionBoundary;
      readonly change: Extract<
        FolioContentStructuralChange,
        { readonly type: "table-column-insert" }
      >;
    }
  | {
      readonly type: "deleteTableColumn";
      readonly source: ResolvedDocxSourceOperand;
      readonly change: Extract<
        FolioContentStructuralChange,
        { readonly type: "table-column-delete" }
      >;
    };

export type ResolvedDocxTableFormatOperandPayload =
  | {
      readonly status: "ready";
      readonly program: TableGeometryProgram;
      readonly changes: readonly ResolvedDocxTableFormatChange[];
    }
  | {
      readonly status: "unsupported";
      readonly issue:
        | { readonly reason: "unprojected-table-structure"; readonly side: "source" | "target" }
        | {
            readonly reason: "unrepresentable-table-geometry";
            readonly issue: TableGeometryUnsupportedIssue;
          };
    };

export type ResolvedDocxTableFormatChange = TableGeometrySemanticChange & {
  /** Canonical event whose paired cell owns this table property scope. */
  readonly sequence: number;
};

export type ResolvedDocxTableStructureOperandInput = ResolvedDocxTableStructureOperandPayload;

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
  readonly structuralEventIndexes: ReadonlyMap<
    FolioContentStructuralChange,
    readonly [number, ...number[]]
  >;
  readonly deletedEventByBlock: ReadonlyMap<
    FolioContentBlock,
    Extract<FolioContentComparisonEvent, { readonly type: "deleted" }>
  >;
  readonly eventSequenceByDeletedEvent: ReadonlyMap<
    Extract<FolioContentComparisonEvent, { readonly type: "deleted" }>,
    number
  >;
  readonly eventSequenceByTableReplacement: ReadonlyMap<FolioContentTableReplacement, number>;
  readonly baseContainerByBlock: ReadonlyMap<FolioContentBlock, FolioContentBaseContainerAlignment>;
  readonly targetContainerByBlock: ReadonlyMap<
    FolioContentBlock,
    FolioContentRevisedContainerAlignment
  >;
  readonly baseTableIndexesByTargetTableIndex: ReadonlyMap<number, ReadonlySet<number>>;
  readonly tableGeometryPairings: readonly TableGeometryPairing[];
  readonly tableGeometrySequenceByPairing: ReadonlyMap<string, number>;
};

type ComparisonOwnedPayload<Payload> = {
  readonly comparison: ResolvedDocxStoryComparison;
  readonly value: Payload;
};

type WholeBlockReplacementPayload = {
  readonly baseBlock: FolioContentBlock;
  readonly targetBlock: FolioContentBlock;
};

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
const payloadByTableStructureOperand = new WeakMap<
  ResolvedDocxTableStructureOperand,
  ComparisonOwnedPayload<ResolvedDocxTableStructureOperandPayload>
>();
const payloadByTableFormatOperand = new WeakMap<
  ResolvedDocxTableFormatOperand,
  ComparisonOwnedPayload<ResolvedDocxTableFormatOperandPayload>
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

const tableGeometryPairingKey = ({ base, target }: TableGeometryPairing): string =>
  `${String(base.tableIndex)}:${String(base.rowIndex)}:${String(base.cellIndex)}>` +
  `${String(target.tableIndex)}:${String(target.rowIndex)}:${String(target.cellIndex)}`;

const createComparisonIndex = (
  payload: ResolvedDocxStoryComparisonPayload,
): ResolvedDocxStoryComparisonIndex => {
  const relations = new Set<FolioContentPairRelation>();
  const eventSequence = new Map<FolioContentComparisonEvent, number>();
  const separators = new Set<FolioContentSeparatorRelation>();
  const structuralChanges = new Set<FolioContentStructuralChange>();
  const tableReplacements = new Set<FolioContentTableReplacement>();
  const mutableStructuralEventIndexes = new Map<FolioContentStructuralChange, number[]>();
  const deletedEventByBlock = new Map<
    FolioContentBlock,
    Extract<FolioContentComparisonEvent, { readonly type: "deleted" }>
  >();
  const eventSequenceByDeletedEvent = new Map<
    Extract<FolioContentComparisonEvent, { readonly type: "deleted" }>,
    number
  >();
  const eventSequenceByTableReplacement = new Map<FolioContentTableReplacement, number>();
  const baseContainerByBlock = new Map<FolioContentBlock, FolioContentBaseContainerAlignment>();
  const targetContainerByBlock = new Map<
    FolioContentBlock,
    FolioContentRevisedContainerAlignment
  >();
  const baseTableIndexesByTargetTableIndex = new Map<number, Set<number>>();
  const tableGeometryPairings: TableGeometryPairing[] = [];
  const tableGeometrySequenceByPairing = new Map<string, number>();
  const pairedBaseCells = new Set<string>();
  const registerRelation = (relation: FolioContentPairRelation, sequence: number): void => {
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
    const pairing = Object.freeze({
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
    });
    tableGeometryPairings.push(pairing);
    tableGeometrySequenceByPairing.set(tableGeometryPairingKey(pairing), sequence);
  };
  for (const [sequence, event] of payload.comparison.events.entries()) {
    eventSequence.set(event, sequence);
    const eventIndex = sequence;
    switch (event.type) {
      case "unchanged":
      case "modified":
      case "formatting":
        registerRelation(event.relation, sequence);
        break;
      case "movedFrom":
      case "movedTo":
        registerRelation(event.move.relation, sequence);
        break;
      case "split":
      case "merge":
        registerRelation(event.relations[0], sequence);
        registerRelation(event.relations[1], sequence);
        registerRelation(event.separator, sequence);
        break;
      case "inserted":
        resolvedDocxSourceOperand(payload.targetSnapshot, event.block);
        targetContainerByBlock.set(event.block, event.containerAlignment);
        break;
      case "deleted":
        resolvedDocxSourceOperand(payload.baseSnapshot, event.block);
        baseContainerByBlock.set(event.block, event.containerAlignment);
        deletedEventByBlock.set(event.block, event);
        eventSequenceByDeletedEvent.set(event, eventIndex);
        break;
      case "tableReplacement":
        tableReplacements.add(event.replacement);
        eventSequenceByTableReplacement.set(event.replacement, eventIndex);
        for (const block of event.replacement.baseBlocks) {
          resolvedDocxSourceOperand(payload.baseSnapshot, block);
        }
        for (const block of event.replacement.revisedBlocks) {
          resolvedDocxSourceOperand(payload.targetSnapshot, block);
        }
        break;
      case "structural": {
        const eventIndexes = mutableStructuralEventIndexes.get(event.change) ?? [];
        eventIndexes.push(eventIndex);
        mutableStructuralEventIndexes.set(event.change, eventIndexes);
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
  const structuralEventIndexes = new Map<
    FolioContentStructuralChange,
    readonly [number, ...number[]]
  >();
  for (const [change, indexes] of mutableStructuralEventIndexes) {
    const first = indexes.at(0) ?? panic("A structural change has no canonical event occurrence");
    structuralEventIndexes.set(change, Object.freeze([first, ...indexes.slice(1)]));
  }
  return Object.freeze({
    eventSequence,
    relations,
    separators,
    structuralChanges,
    tableReplacements,
    structuralEventIndexes,
    deletedEventByBlock,
    eventSequenceByDeletedEvent,
    eventSequenceByTableReplacement,
    baseContainerByBlock,
    targetContainerByBlock,
    baseTableIndexesByTargetTableIndex,
    tableGeometryPairings: Object.freeze(tableGeometryPairings),
    tableGeometrySequenceByPairing,
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
  }: {
    readonly events: NonEmptyReadonlyArray<ResolvedDocxDeletedEventOperand>;
  },
): ResolvedDocxTrailingDeletionOperand => {
  const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
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
  if (!container || container.base.end !== "paragraph") {
    return panic("A trailing DOCX deletion operand must remain in one canonical container");
  }
  const baseBlocks = resolvedDocxContentBlocks(baseSnapshot);
  const basePositionByBlock = new Map(baseBlocks.map((block, position) => [block, position]));
  const eventPositions = ownedEvents.map(({ event }) => basePositionByBlock.get(event.block));
  const firstPosition = eventPositions.at(0);
  if (
    firstPosition === undefined ||
    ownedEvents.some(({ event }) => index.baseContainerByBlock.get(event.block) !== container) ||
    eventPositions.some((position, offset) => position !== firstPosition + offset) ||
    ownedEvents.some(({ sequence }, offset) => {
      if (offset === 0) return false;
      const previous = ownedEvents.at(offset - 1);
      return previous === undefined || sequence <= previous.sequence;
    })
  ) {
    return panic("A trailing DOCX deletion operand must follow canonical base order");
  }
  const chainStartBlock = baseBlocks[firstPosition - 1];
  const terminalBlock = baseBlocks.findLast(
    (block) => index.baseContainerByBlock.get(block) === container,
  );
  if (
    !chainStartBlock ||
    index.baseContainerByBlock.get(chainStartBlock) !== container ||
    terminalBlock !== ownedEvents.at(-1)?.event.block
  ) {
    return panic("A trailing DOCX deletion operand must be contiguous and terminal");
  }
  const chainStart = resolvedDocxSourceOperand(baseSnapshot, chainStartBlock);
  const targetBlock = resolvedDocxContentBlocks(targetSnapshot).findLast(
    (block) => index.targetContainerByBlock.get(block) === container,
  );
  if (targetBlock && container.type !== "paired") {
    return panic("A trailing DOCX deletion target carrier belongs to an unpaired container");
  }
  const operand = Object.freeze({
    [RESOLVED_DOCX_TRAILING_DELETION_OPERAND_BRAND]: true as const,
  });
  payloadByTrailingDeletionOperand.set(operand, {
    comparison,
    value: Object.freeze({
      events: Object.freeze(ownedEvents),
      chainStart,
      ...(targetBlock !== undefined && { targetCarrier: targetBlock }),
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
  const deletedBoundary = deletedPayload.event.removalBoundary;
  const insertedBoundary = insertedPayload.event.boundary;
  if (!baseContainer || baseContainer !== targetContainer || baseContainer.type !== "paired") {
    return panic("A terminal DOCX replacement must remain in one canonical container");
  }
  if (
    baseContainer.base.end !== "paragraph" ||
    baseContainer.revised.end !== "paragraph" ||
    deletedBoundary.type !== "terminalPredecessor" ||
    insertedBoundary.type === "unanchoredContainer" ||
    deletedBoundary.containerAlignment !== baseContainer ||
    insertedBoundary.containerAlignment !== baseContainer ||
    deletedBoundary.targetCarrier !== insertedPayload.event.block ||
    insertedBoundary.paragraph !== deletedBoundary.predecessor
  ) {
    return panic("A terminal DOCX replacement must own the exact terminal event pair");
  }
  const baseTerminal = resolvedDocxContentBlocks(
    resolvedDocxStoryComparisonPayload(comparison).baseSnapshot,
  ).findLast((block) => index.baseContainerByBlock.get(block) === baseContainer);
  const targetTerminal = resolvedDocxContentBlocks(
    resolvedDocxStoryComparisonPayload(comparison).targetSnapshot,
  ).findLast((block) => index.targetContainerByBlock.get(block) === baseContainer);
  if (
    baseTerminal !== deletedPayload.event.block ||
    targetTerminal !== insertedPayload.event.block ||
    deletedPayload.sequence >= insertedPayload.sequence
  ) {
    return panic("A terminal DOCX replacement must own the exact terminal event pair");
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

const structuralChangeOf = (
  input: Exclude<ResolvedDocxTableStructureOperandInput, { readonly type: "replaceTable" }>,
): FolioContentStructuralChange => input.change;

const structuralSourceOf = (
  input: ResolvedDocxTableStructureOperandInput,
): ResolvedDocxSourceOperand | null => {
  switch (input.type) {
    case "insertTable":
    case "insertTableRow":
    case "insertTableColumn":
      return input.anchor.source;
    case "deleteTable":
    case "deleteTableRow":
    case "deleteTableColumn":
    case "replaceTable":
      return input.source;
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX table-structure operand input", { input: unreachable });
    }
  }
};

const requireExactStructuralSource = (
  input: Extract<
    ResolvedDocxTableStructureOperandInput,
    { readonly type: "deleteTable" | "deleteTableRow" | "deleteTableColumn" }
  >,
  payload: ResolvedDocxStoryComparisonPayload,
): void => {
  const sourceBlock = resolvedDocxSourceOperandBlock(input.source, payload.baseSnapshot);
  if (input.change.blocks.at(0) !== sourceBlock) {
    return panic("A DOCX table-structure source must own the canonical change's first block");
  }
};

const tableReplacementBaseBlocks = (
  owner: ResolvedDocxTableReplacementOwner,
): FolioContentBlockGroup =>
  owner.type === "canonical-replacement" ? owner.replacement.baseBlocks : owner.deleted.blocks;

const tableReplacementIndexes = (
  owner: ResolvedDocxTableReplacementOwner,
): { readonly base: number; readonly target: number } =>
  owner.type === "canonical-replacement"
    ? { base: owner.replacement.baseTableIndex, target: owner.replacement.revisedTableIndex }
    : { base: owner.deleted.tableIndex, target: owner.inserted.tableIndex };

const requireExactTableReplacementOwner = (
  owner: ResolvedDocxTableReplacementOwner,
  comparison: ResolvedDocxStoryComparison,
): void => {
  const index = comparisonIndexOf(comparison);
  if (owner.type === "canonical-replacement") {
    if (!index.tableReplacements.has(owner.replacement)) {
      return panic("A DOCX table replacement must name an exact canonical replacement");
    }
    return;
  }
  if (!index.structuralChanges.has(owner.deleted) || !index.structuralChanges.has(owner.inserted)) {
    return panic("A DOCX table replacement pair must name exact canonical structural changes");
  }
  const deletedIndexes = index.structuralEventIndexes.get(owner.deleted);
  const insertedIndexes = index.structuralEventIndexes.get(owner.inserted);
  if (!deletedIndexes || !insertedIndexes) {
    return panic("A DOCX table replacement pair lost its canonical event sequence");
  }
  const spanStart = Math.min(deletedIndexes.at(0) ?? 0, insertedIndexes.at(0) ?? 0);
  const spanEnd = Math.max(deletedIndexes.at(-1) ?? 0, insertedIndexes.at(-1) ?? 0);
  const events = resolvedDocxStoryComparisonPayload(comparison).comparison.events;
  for (let eventIndex = spanStart; eventIndex <= spanEnd; eventIndex++) {
    const event = events[eventIndex];
    if (
      event?.type !== "structural" ||
      (event.change !== owner.deleted && event.change !== owner.inserted)
    ) {
      return panic("A DOCX table replacement pair is not one contiguous canonical event sequence");
    }
  }
  const basePath = canonicalJson(owner.deleted.blocks.at(0)?.containerPath ?? []);
  const targetPath = canonicalJson(owner.inserted.blocks.at(0)?.containerPath ?? []);
  if (basePath !== targetPath) {
    return panic("A DOCX table replacement pair crosses canonical container ownership");
  }
};

const requirePairedStructuralAnchor = (
  input: Extract<
    ResolvedDocxTableStructureOperandInput,
    { readonly type: "insertTableRow" | "insertTableColumn" }
  >,
  comparison: ResolvedDocxStoryComparison,
  payload: ResolvedDocxStoryComparisonPayload,
): void => {
  const anchor = resolvedDocxSourceOperandBlock(input.anchor.source, payload.baseSnapshot);
  const anchorTableIndex = anchor.table?.tableIndex;
  if (
    anchorTableIndex === undefined ||
    !resolvedDocxPairedBaseTableIndexes(comparison, input.change.tableIndex).has(anchorTableIndex)
  ) {
    return panic("A DOCX table-structure anchor must belong to the canonically paired table");
  }
};

const requireExactTerminalCarrier = (
  carrier: ResolvedDocxTerminalTableCarrier,
  comparison: ResolvedDocxStoryComparison,
  payload: ResolvedDocxStoryComparisonPayload,
): void => {
  const block = resolvedDocxSourceOperandBlock(carrier.source, payload.baseSnapshot);
  if (
    carrier.event.block !== block ||
    comparisonIndexOf(comparison).deletedEventByBlock.get(block) !== carrier.event
  ) {
    return panic("A terminal table carrier must own its exact canonical deletion event");
  }
};

/** Issue one opaque structural obligation from exact comparison-owned semantics. */
export const resolvedDocxTableStructureOperand = (
  comparison: ResolvedDocxStoryComparison,
  input: ResolvedDocxTableStructureOperandInput,
): ResolvedDocxTableStructureOperand => {
  const payload = resolvedDocxStoryComparisonPayload(comparison);
  const index = comparisonIndexOf(comparison);
  const source = structuralSourceOf(input);
  if (source) resolvedDocxSourceOperandBlock(source, payload.baseSnapshot);
  switch (input.type) {
    case "replaceTable": {
      requireExactTableReplacementOwner(input.owner, comparison);
      if (
        tableReplacementBaseBlocks(input.owner).at(0) !==
        resolvedDocxSourceOperandBlock(input.source, payload.baseSnapshot)
      ) {
        return panic("A DOCX table replacement must own its exact base table");
      }
      if (input.terminalCarrier) {
        requireExactTerminalCarrier(input.terminalCarrier, comparison, payload);
      }
      break;
    }
    case "insertTableRow":
    case "insertTableColumn": {
      if (!index.structuralChanges.has(structuralChangeOf(input))) {
        return panic("A DOCX table-structure operand must name an exact canonical change");
      }
      requirePairedStructuralAnchor(input, comparison, payload);
      break;
    }
    case "deleteTable":
    case "deleteTableRow":
    case "deleteTableColumn": {
      if (!index.structuralChanges.has(structuralChangeOf(input))) {
        return panic("A DOCX table-structure operand must name an exact canonical change");
      }
      requireExactStructuralSource(input, payload);
      break;
    }
    case "insertTable": {
      if (!index.structuralChanges.has(structuralChangeOf(input))) {
        return panic("A DOCX table-structure operand must name an exact canonical change");
      }
      if (input.terminalCarrier) {
        requireExactTerminalCarrier(input.terminalCarrier, comparison, payload);
      }
      break;
    }
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX table-structure operand input", { input: unreachable });
    }
  }
  let operand: ResolvedDocxTableStructureOperand;
  switch (input.type) {
    case "insertTable":
      operand = Object.freeze({
        type: "insertTable",
        [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true as const,
      });
      break;
    case "deleteTable":
      operand = Object.freeze({
        type: "deleteTable",
        [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true as const,
      });
      break;
    case "replaceTable":
      operand = Object.freeze({
        type: "replaceTable",
        [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true as const,
      });
      break;
    case "insertTableRow":
      operand = Object.freeze({
        type: "insertTableRow",
        [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true as const,
      });
      break;
    case "deleteTableRow":
      operand = Object.freeze({
        type: "deleteTableRow",
        [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true as const,
      });
      break;
    case "insertTableColumn":
      operand = Object.freeze({
        type: "insertTableColumn",
        [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true as const,
      });
      break;
    case "deleteTableColumn":
      operand = Object.freeze({
        type: "deleteTableColumn",
        [RESOLVED_DOCX_TABLE_STRUCTURE_OPERAND_BRAND]: true as const,
      });
      break;
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX table-structure operand input", { input: unreachable });
    }
  }
  payloadByTableStructureOperand.set(operand, {
    comparison,
    value: Object.freeze({ ...input }),
  });
  return operand;
};

/** Resolve a structural obligation only inside its issuing comparison. */
export const resolvedDocxTableStructureOperandPayload = (
  operand: ResolvedDocxTableStructureOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxTableStructureOperandPayload => {
  const payload = requireComparisonOwned(
    payloadByTableStructureOperand.get(operand),
    comparison,
    "DOCX table-structure operand",
  );
  if (payload.type !== operand.type) {
    return panic("A DOCX table-structure operand lost its semantic discriminator");
  }
  return payload;
};

export type ResolvedDocxTableStructureReportOwner =
  | {
      readonly type: "structural";
      readonly sequence: number;
      readonly change: FolioContentStructuralChange;
    }
  | {
      readonly type: "replacement";
      readonly sequence: number;
      readonly replacement: FolioContentTableReplacement;
    }
  | {
      readonly type: "deletedCarrier";
      readonly sequence: number;
      readonly event: Extract<FolioContentComparisonEvent, { readonly type: "deleted" }>;
    };

/** Exact canonical report owners retained by one atomic table operation. */
export const resolvedDocxTableStructureReportOwners = (
  operand: ResolvedDocxTableStructureOperand,
  comparison: ResolvedDocxStoryComparison,
): readonly ResolvedDocxTableStructureReportOwner[] => {
  const payload = resolvedDocxTableStructureOperandPayload(operand, comparison);
  const index = comparisonIndexOf(comparison);
  const owners: ResolvedDocxTableStructureReportOwner[] = [];
  const pushStructural = (change: FolioContentStructuralChange): void => {
    const sequence = index.structuralEventIndexes.get(change)?.at(0);
    if (sequence === undefined) {
      return panic("A table operation lost its canonical structural event sequence");
    }
    owners.push(Object.freeze({ type: "structural", sequence, change }));
  };
  switch (payload.type) {
    case "insertTable":
    case "deleteTable":
    case "insertTableRow":
    case "deleteTableRow":
    case "insertTableColumn":
    case "deleteTableColumn":
      pushStructural(payload.change);
      break;
    case "replaceTable":
      if (payload.owner.type === "canonical-replacement") {
        const sequence = index.eventSequenceByTableReplacement.get(payload.owner.replacement);
        if (sequence === undefined) {
          return panic("A table operation lost its canonical replacement event sequence");
        }
        owners.push(
          Object.freeze({
            type: "replacement",
            sequence,
            replacement: payload.owner.replacement,
          }),
        );
      } else {
        pushStructural(payload.owner.deleted);
        pushStructural(payload.owner.inserted);
      }
      break;
    default: {
      const unreachable: never = payload;
      return panic("Unhandled table report owner", { payload: unreachable });
    }
  }
  const carrier =
    payload.type === "insertTable" || payload.type === "replaceTable"
      ? payload.terminalCarrier
      : undefined;
  if (carrier) {
    const sequence = index.eventSequenceByDeletedEvent.get(carrier.event);
    if (sequence === undefined) {
      return panic("A table operation lost its canonical terminal-carrier event sequence");
    }
    owners.push(Object.freeze({ type: "deletedCarrier", sequence, event: carrier.event }));
  }
  return Object.freeze(owners.toSorted((left, right) => left.sequence - right.sequence));
};

const tableHasHiddenRows = (table: PMNode): boolean => {
  for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
    if (table.child(rowIndex).attrs["hidden"] === true) return true;
  }
  return false;
};

const issueTableFormatOperand = (
  comparison: ResolvedDocxStoryComparison,
  payload: ResolvedDocxTableFormatOperandPayload,
): ResolvedDocxTableFormatOperand => {
  const operand = Object.freeze({
    type: "matchTableFormatting" as const,
    [RESOLVED_DOCX_TABLE_FORMAT_OPERAND_BRAND]: true as const,
  });
  payloadByTableFormatOperand.set(operand, { comparison, value: Object.freeze(payload) });
  return operand;
};

/**
 * Resolve paired table properties once. Unchanged tables produce no semantic
 * operation; changed and unsupported projections retain one exact owner.
 */
export const resolvedDocxTableFormatOperand = (
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxTableFormatOperand | null => {
  const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
  const pairings = resolvedDocxTableGeometryPairings(comparison);
  if (pairings.length === 0) return null;
  const baseTables = storyTablesOf(resolvedDocxOperationSnapshot(baseSnapshot));
  const baseByIndex = new Map(baseTables.map((table) => [table.index, table.node]));
  const targetTables = resolvedDocxTableNodes(targetSnapshot);
  const inspectedPairs = new Set<string>();
  for (const { base, target } of pairings) {
    const pair = `${String(base.tableIndex)}:${String(target.tableIndex)}`;
    if (inspectedPairs.has(pair)) continue;
    inspectedPairs.add(pair);
    const baseTable = baseByIndex.get(base.tableIndex);
    const targetTable = targetTables.get(target.tableIndex);
    if (baseTable && tableHasHiddenRows(baseTable)) {
      return issueTableFormatOperand(comparison, {
        status: "unsupported",
        issue: { reason: "unprojected-table-structure", side: "source" },
      });
    }
    if (targetTable && tableHasHiddenRows(targetTable)) {
      return issueTableFormatOperand(comparison, {
        status: "unsupported",
        issue: { reason: "unprojected-table-structure", side: "target" },
      });
    }
  }
  const preflight = preflightTableGeometry({ baseTables, targetTables, pairings });
  if (preflight.status === "unsupported") {
    return issueTableFormatOperand(comparison, {
      status: "unsupported",
      issue: { reason: "unrepresentable-table-geometry", issue: preflight.issue },
    });
  }
  const semanticChanges = tableGeometryProgramSemanticChanges(preflight.program);
  if (semanticChanges.length === 0) return null;
  const index = comparisonIndexOf(comparison);
  const changes = semanticChanges.map((change) => {
    const sequence = index.tableGeometrySequenceByPairing.get(tableGeometryPairingKey(change));
    if (sequence === undefined) {
      return panic("A table-format change lost its canonical paired-cell event");
    }
    return Object.freeze({
      scope: change.scope,
      base: change.base,
      target: change.target,
      sequence,
    });
  });
  return issueTableFormatOperand(comparison, {
    status: "ready",
    program: preflight.program,
    changes,
  });
};

/** Resolve an exact table-format operation inside its issuing comparison. */
export const resolvedDocxTableFormatOperandPayload = (
  operand: ResolvedDocxTableFormatOperand,
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxTableFormatOperandPayload =>
  requireComparisonOwned(
    payloadByTableFormatOperand.get(operand),
    comparison,
    "DOCX table-format operand",
  );

/** Bind the exact final empty carrier into its replacement obligation. */
export const resolvedDocxTerminalTableReplacementOperand = (
  comparison: ResolvedDocxStoryComparison,
  operand: Extract<ResolvedDocxTableStructureOperand, { readonly type: "replaceTable" }>,
  terminalCarrier: ResolvedDocxSourceOperand,
): Extract<ResolvedDocxTableStructureOperand, { readonly type: "replaceTable" }> => {
  const payload = resolvedDocxTableStructureOperandPayload(operand, comparison);
  if (payload.type !== "replaceTable") {
    return panic("A terminal table carrier can only belong to a replacement");
  }
  const comparisonPayload = resolvedDocxStoryComparisonPayload(comparison);
  const carrier = resolvedDocxSourceOperandBlock(terminalCarrier, comparisonPayload.baseSnapshot);
  const event = comparisonIndexOf(comparison).deletedEventByBlock.get(carrier);
  if (!event) {
    return panic("A terminal table replacement carrier must be an exact canonical deletion");
  }
  const targetBlocks = resolvedDocxContentBlocks(comparisonPayload.targetSnapshot);
  if (
    carrier.kind !== "paragraph" ||
    carrier.text.length !== 0 ||
    targetBlocks.at(-1)?.table?.outerTableIndex !== tableReplacementIndexes(payload.owner).target
  ) {
    return panic("A terminal table replacement must own an empty final carrier and target table");
  }
  const resolved = resolvedDocxTableStructureOperand(comparison, {
    ...payload,
    terminalCarrier: Object.freeze({ source: terminalCarrier, event }),
  });
  if (resolved.type !== "replaceTable") {
    return panic("A terminal table replacement lost its structural discriminator");
  }
  return resolved;
};

/** Bind the exact final empty carrier into its terminal table insertion. */
export const resolvedDocxTerminalTableInsertionOperand = (
  comparison: ResolvedDocxStoryComparison,
  operand: Extract<ResolvedDocxTableStructureOperand, { readonly type: "insertTable" }>,
  terminalCarrier: ResolvedDocxSourceOperand,
): Extract<ResolvedDocxTableStructureOperand, { readonly type: "insertTable" }> => {
  const payload = resolvedDocxTableStructureOperandPayload(operand, comparison);
  if (payload.type !== "insertTable") {
    return panic("A terminal table carrier can only belong to an insertion");
  }
  const comparisonPayload = resolvedDocxStoryComparisonPayload(comparison);
  const carrier = resolvedDocxSourceOperandBlock(terminalCarrier, comparisonPayload.baseSnapshot);
  const event = comparisonIndexOf(comparison).deletedEventByBlock.get(carrier);
  if (!event) {
    return panic("A terminal table insertion carrier must be an exact canonical deletion");
  }
  const targetBlocks = resolvedDocxContentBlocks(comparisonPayload.targetSnapshot);
  if (
    carrier.kind !== "paragraph" ||
    carrier.text.length !== 0 ||
    targetBlocks.at(-1)?.table?.outerTableIndex !== payload.change.tableIndex
  ) {
    return panic("A terminal table insertion must own an empty final carrier and target table");
  }
  const resolved = resolvedDocxTableStructureOperand(comparison, {
    ...payload,
    terminalCarrier: Object.freeze({ source: terminalCarrier, event }),
  });
  if (resolved.type !== "insertTable") {
    return panic("A terminal table insertion lost its structural discriminator");
  }
  return resolved;
};

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
