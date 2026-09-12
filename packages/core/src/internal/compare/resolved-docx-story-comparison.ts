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
  FolioContentContainerAlignment,
  FolioContentPairedContainerAlignment,
  FolioContentRevisedContainerAlignment,
} from "../../compare/content-types";
import { canonicalJson } from "../../utils/canonicalJson";
import {
  preflightTableGeometryComponents,
  tableGeometryProgramSemanticChangeOccurrences,
  tableGeometryProgramTableGridTransitions,
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
const RESOLVED_DOCX_TABLE_COMPONENT_MEMBER_BRAND: unique symbol = Symbol(
  "resolved-docx-table-component-member",
);
const RESOLVED_DOCX_TABLE_COMPONENT_BRAND: unique symbol = Symbol("resolved-docx-table-component");
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

/** One comparison-owned table scope used only to derive atomic components. */
type ResolvedDocxTableComponentMember = {
  readonly [RESOLVED_DOCX_TABLE_COMPONENT_MEMBER_BRAND]: true;
};

/** One connected, comparison-owned table execution scope. */
export type ResolvedDocxTableComponent = {
  readonly [RESOLVED_DOCX_TABLE_COMPONENT_BRAND]: true;
};

export type ResolvedDocxStructuralInsertionBoundary = {
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
      readonly anchor: ResolvedDocxStructuralInsertionBoundary;
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
      readonly anchor: ResolvedDocxStructuralInsertionBoundary;
      readonly change: Extract<FolioContentStructuralChange, { readonly type: "table-row-insert" }>;
    }
  | {
      readonly type: "deleteTableRow";
      readonly source: ResolvedDocxSourceOperand;
      readonly change: Extract<FolioContentStructuralChange, { readonly type: "table-row-delete" }>;
    }
  | {
      readonly type: "insertTableColumn";
      readonly anchor: ResolvedDocxStructuralInsertionBoundary;
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

export type ResolvedDocxTableFormatChange = {
  /** Canonical event whose paired cell owns this table property scope. */
  readonly sequence: number;
  readonly change: TableGeometrySemanticChange;
};

export type ResolvedDocxTableStructureOperandInput =
  | Exclude<
      ResolvedDocxTableStructureOperandPayload,
      { readonly type: "insertTable" | "replaceTable" }
    >
  | Omit<
      Extract<ResolvedDocxTableStructureOperandPayload, { readonly type: "insertTable" }>,
      "terminalCarrier"
    >
  | Omit<
      Extract<ResolvedDocxTableStructureOperandPayload, { readonly type: "replaceTable" }>,
      "terminalCarrier"
    >;

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
  readonly baseContainerByBlockId: ReadonlyMap<string, FolioContentBaseContainerAlignment>;
  readonly baseBlockById: ReadonlyMap<string, FolioContentBlock>;
  readonly basePositionByBlock: ReadonlyMap<FolioContentBlock, number>;
  readonly baseTerminalBlockByContainer: ReadonlyMap<
    FolioContentBaseContainerAlignment,
    FolioContentBlock
  >;
  readonly targetTerminalBlockByContainer: ReadonlyMap<
    FolioContentRevisedContainerAlignment,
    FolioContentBlock
  >;
  readonly baseTerminalBlocks: readonly ResolvedDocxBaseTerminalBlock[];
  readonly targetTerminalOuterTableIndex: number | undefined;
  readonly baseTableIndexByTargetTableIndex: ReadonlyMap<number, number>;
  readonly structuralInsertionBoundaryByChange: ReadonlyMap<
    ResolvedDocxInsertionStructuralChange,
    ResolvedDocxStructuralInsertionBoundary
  >;
  readonly structurallyOwnedTablePlacementRelations: ReadonlySet<FolioContentPairRelation>;
  readonly tableGeometryPairings: readonly TableGeometryPairing[];
  readonly tableGeometrySequenceByPairing: ReadonlyMap<string, number>;
  readonly baseTableComponentMemberByTableIndex: ReadonlyMap<
    number,
    ResolvedDocxTableComponentMember
  >;
  readonly targetTableComponentMemberByTableIndex: ReadonlyMap<
    number,
    ResolvedDocxTableComponentMember
  >;
  readonly terminalComponentMemberByBlock: ReadonlyMap<
    FolioContentBlock,
    ResolvedDocxTableComponentMember
  >;
  readonly work: ResolvedDocxComparisonIndexWork;
};

export type ResolvedDocxInsertionStructuralChange = Extract<
  FolioContentStructuralChange,
  {
    readonly type: "table-insert" | "table-row-insert" | "table-column-insert";
  }
>;

export type ResolvedDocxBaseTerminalBlock = {
  readonly alignment: FolioContentBaseContainerAlignment;
  readonly block: FolioContentBlock;
};

/** Deterministic construction work retained for focused linearity invariants. @internal */
export type ResolvedDocxComparisonIndexWork = {
  readonly eventVisits: number;
  readonly alignmentVisits: number;
  readonly relationVisits: number;
  readonly baseBlockVisits: number;
  readonly revisedBlockVisits: number;
  readonly anchorEventVisits: number;
  readonly structuralChangeVisits: number;
  readonly total: number;
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
const componentMembersByTableOperand = new WeakMap<
  ResolvedDocxTableStructureOperand | ResolvedDocxTableFormatOperand,
  ComparisonOwnedPayload<readonly ResolvedDocxTableComponentMember[]>
>();
const operandsByTableComponent = new WeakMap<
  ResolvedDocxTableComponent,
  ComparisonOwnedPayload<
    readonly (ResolvedDocxTableStructureOperand | ResolvedDocxTableFormatOperand)[]
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

const tableGeometryPairingKey = ({ base, target }: TableGeometryPairing): string =>
  `${String(base.tableIndex)}:${String(base.rowIndex)}:${String(base.cellIndex)}>` +
  `${String(target.tableIndex)}:${String(target.rowIndex)}:${String(target.cellIndex)}`;

const tableCellCoordinateKey = ({
  tableIndex,
  rowIndex,
  cellIndex,
}: TableGeometryPairing["base"]): string =>
  `${String(tableIndex)}:${String(rowIndex)}:${String(cellIndex)}`;

const compareTableCellCoordinates = (
  left: TableGeometryPairing["base"],
  right: TableGeometryPairing["base"],
): number =>
  left.tableIndex - right.tableIndex ||
  left.rowIndex - right.rowIndex ||
  left.cellIndex - right.cellIndex;

type PairedTableCellOccurrence = Extract<
  FolioContentPairedContainerAlignment["base"],
  { readonly type: "tableCell" }
>;

const tableLocationMatchesOccurrence = (
  location: NonNullable<FolioContentBlock["table"]>,
  occurrence: PairedTableCellOccurrence,
): boolean =>
  location.outerTableIdentity.type === occurrence.table.outerTableIdentity.type &&
  location.outerTableIdentity.id === occurrence.table.outerTableIdentity.id &&
  location.tableIdentity.type === occurrence.table.tableIdentity.type &&
  location.tableIdentity.id === occurrence.table.tableIdentity.id &&
  location.rowIdentity.type === occurrence.table.rowIdentity.type &&
  location.rowIdentity.id === occurrence.table.rowIdentity.id &&
  location.cellIdentity.type === occurrence.table.cellIdentity.type &&
  location.cellIdentity.id === occurrence.table.cellIdentity.id &&
  location.outerTableIndex === occurrence.table.outerTableIndex &&
  location.tableIndex === occurrence.table.tableIndex &&
  location.rowIndex === occurrence.table.rowIndex &&
  location.cellIndex === occurrence.table.cellIndex &&
  location.gridColumnIndex === occurrence.table.gridColumnIndex &&
  location.columnSpan === occurrence.table.columnSpan &&
  location.rowSpan === occurrence.table.rowSpan;

const firstBaseBlockOfEvent = (event: FolioContentComparisonEvent): FolioContentBlock | null => {
  switch (event.type) {
    case "unchanged":
    case "modified":
    case "formatting":
      return event.relation.base.block;
    case "deleted":
      return event.block;
    case "inserted":
    case "movedTo":
      return null;
    case "movedFrom":
      return event.move.relation.base.block;
    case "split":
    case "merge":
      return event.relations[0].base.block;
    case "tableReplacement":
      return event.replacement.baseBlocks[0];
    case "structural":
      return event.change.type === "table-delete" ||
        event.change.type === "table-row-delete" ||
        event.change.type === "table-column-delete"
        ? (event.change.blocks[event.memberIndex] ??
            panic("A structural deletion lost its canonical base member"))
        : null;
    default: {
      const unreachable: never = event;
      return panic("Unhandled comparison event while indexing base occurrences", {
        event: unreachable,
      });
    }
  }
};

const baseTableBlockOfEvent = (event: FolioContentComparisonEvent): FolioContentBlock | null => {
  const block = firstBaseBlockOfEvent(event);
  return block?.table === undefined ? null : block;
};

const createComparisonIndex = (
  payload: ResolvedDocxStoryComparisonPayload,
): ResolvedDocxStoryComparisonIndex => {
  const work = {
    eventVisits: 0,
    alignmentVisits: 0,
    relationVisits: 0,
    baseBlockVisits: 0,
    revisedBlockVisits: 0,
    anchorEventVisits: 0,
    structuralChangeVisits: 0,
  };
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
  const baseContainerByBlockId = new Map<string, FolioContentBaseContainerAlignment>();
  const basePositionByBlock = new Map<FolioContentBlock, number>();
  const targetPositionByBlock = new Map<FolioContentBlock, number>();
  const baseBlockById = new Map<string, FolioContentBlock>();
  const baseTableComponentMemberByOuterIndex = new Map<number, ResolvedDocxTableComponentMember>();
  const targetTableComponentMemberByOuterIndex = new Map<
    number,
    ResolvedDocxTableComponentMember
  >();
  const baseTableComponentMemberByTableIndex = new Map<number, ResolvedDocxTableComponentMember>();
  const targetTableComponentMemberByTableIndex = new Map<
    number,
    ResolvedDocxTableComponentMember
  >();
  const componentMemberFor = (
    members: Map<number, ResolvedDocxTableComponentMember>,
    outerTableIndex: number,
  ): ResolvedDocxTableComponentMember => {
    const existing = members.get(outerTableIndex);
    if (existing) return existing;
    const member = Object.freeze({
      [RESOLVED_DOCX_TABLE_COMPONENT_MEMBER_BRAND]: true as const,
    });
    members.set(outerTableIndex, member);
    return member;
  };
  const registerTableComponentMember = (
    block: FolioContentBlock,
    membersByOuterIndex: Map<number, ResolvedDocxTableComponentMember>,
    membersByTableIndex: Map<number, ResolvedDocxTableComponentMember>,
  ): void => {
    if (!block.table) return;
    const member = componentMemberFor(membersByOuterIndex, block.table.outerTableIndex);
    const existing = membersByTableIndex.get(block.table.tableIndex);
    if (existing !== undefined && existing !== member) {
      return panic("A canonical table belongs to two outer-table components", {
        tableIndex: block.table.tableIndex,
      });
    }
    membersByTableIndex.set(block.table.tableIndex, member);
  };
  let trailingBodyBlock: FolioContentBlock | null = null;
  const baseBlocks = resolvedDocxContentBlocks(payload.baseSnapshot);
  const targetBlocks = resolvedDocxContentBlocks(payload.targetSnapshot);
  for (const [position, block] of baseBlocks.entries()) {
    work.baseBlockVisits += 1;
    basePositionByBlock.set(block, position);
    baseBlockById.set(block.identity.id, block);
    registerTableComponentMember(
      block,
      baseTableComponentMemberByOuterIndex,
      baseTableComponentMemberByTableIndex,
    );
    if (block.table === undefined) trailingBodyBlock = block;
  }
  for (const [position, block] of targetBlocks.entries()) {
    work.revisedBlockVisits += 1;
    targetPositionByBlock.set(block, position);
    registerTableComponentMember(
      block,
      targetTableComponentMemberByOuterIndex,
      targetTableComponentMemberByTableIndex,
    );
  }
  const baseTerminalBlockByContainer = new Map<
    FolioContentBaseContainerAlignment,
    FolioContentBlock
  >();
  const targetTerminalBlockByContainer = new Map<
    FolioContentRevisedContainerAlignment,
    FolioContentBlock
  >();
  const baseTerminalPositionByContainer = new Map<FolioContentBaseContainerAlignment, number>();
  const targetTerminalPositionByContainer = new Map<
    FolioContentRevisedContainerAlignment,
    number
  >();
  const pairedContainerAlignments = new Set<FolioContentPairedContainerAlignment>();
  const alignmentsWithParagraphStructureChange = new Set<FolioContentContainerAlignment>();
  const moveRelations = new Set<FolioContentPairRelation>();
  const baseTablesWithRowStructureChange = new Set<number>();
  const targetTablesWithRowStructureChange = new Set<number>();
  const baseTablesWithColumnStructureChange = new Set<number>();
  const targetTablesWithColumnStructureChange = new Set<number>();
  const baseSequenceByAlignment = new Map<FolioContentPairedContainerAlignment, number>();
  const revisedSequenceByAlignment = new Map<FolioContentPairedContainerAlignment, number>();
  const registerBaseBlock = (
    block: FolioContentBlock,
    alignment: FolioContentBaseContainerAlignment,
  ): void => {
    resolvedDocxSourceOperand(payload.baseSnapshot, block);
    const existing = baseContainerByBlock.get(block);
    const existingById = baseContainerByBlockId.get(block.identity.id);
    if (
      (existing !== undefined && existing !== alignment) ||
      (existingById !== undefined && existingById !== alignment)
    ) {
      return panic("A base block belongs to two canonical container alignments", {
        blockId: block.identity.id,
      });
    }
    baseContainerByBlock.set(block, alignment);
    baseContainerByBlockId.set(block.identity.id, alignment);
    const position =
      basePositionByBlock.get(block) ??
      panic("A comparison base block is absent from its canonical snapshot", {
        blockId: block.identity.id,
      });
    if ((baseTerminalPositionByContainer.get(alignment) ?? -1) < position) {
      baseTerminalPositionByContainer.set(alignment, position);
      baseTerminalBlockByContainer.set(alignment, block);
    }
  };
  const registerTargetBlock = (
    block: FolioContentBlock,
    alignment: FolioContentRevisedContainerAlignment,
  ): void => {
    resolvedDocxSourceOperand(payload.targetSnapshot, block);
    const existing = targetContainerByBlock.get(block);
    if (existing !== undefined && existing !== alignment) {
      return panic("A revised block belongs to two canonical container alignments", {
        blockId: block.identity.id,
      });
    }
    targetContainerByBlock.set(block, alignment);
    const position =
      targetPositionByBlock.get(block) ??
      panic("A comparison revised block is absent from its canonical snapshot", {
        blockId: block.identity.id,
      });
    if ((targetTerminalPositionByContainer.get(alignment) ?? -1) < position) {
      targetTerminalPositionByContainer.set(alignment, position);
      targetTerminalBlockByContainer.set(alignment, block);
    }
  };
  const registerBaseOccurrence = (
    block: FolioContentBlock,
    alignment: FolioContentBaseContainerAlignment,
    sequence: number,
  ): void => {
    registerBaseBlock(block, alignment);
    if (alignment.type !== "paired") return;
    pairedContainerAlignments.add(alignment);
    const current = baseSequenceByAlignment.get(alignment);
    if (current === undefined || sequence < current)
      baseSequenceByAlignment.set(alignment, sequence);
  };
  const registerTargetOccurrence = (
    block: FolioContentBlock,
    alignment: FolioContentRevisedContainerAlignment,
    sequence: number,
  ): void => {
    registerTargetBlock(block, alignment);
    if (alignment.type !== "paired") return;
    pairedContainerAlignments.add(alignment);
    const current = revisedSequenceByAlignment.get(alignment);
    if (current === undefined || sequence < current) {
      revisedSequenceByAlignment.set(alignment, sequence);
    }
  };
  const registerRelationOwnership = (relation: FolioContentPairRelation): void => {
    if (relations.has(relation)) return;
    relations.add(relation);
    registerBaseBlock(relation.base.block, relation.baseContainerAlignment);
    registerTargetBlock(relation.revised.block, relation.revisedContainerAlignment);
    if (relation.relationType === "separator") separators.add(relation);
  };
  const registerTwoSidedRelation = (relation: FolioContentPairRelation, sequence: number): void => {
    registerRelationOwnership(relation);
    registerBaseOccurrence(relation.base.block, relation.baseContainerAlignment, sequence);
    registerTargetOccurrence(relation.revised.block, relation.revisedContainerAlignment, sequence);
  };
  const baseTableIndexByTargetTableIndex = new Map<number, number>();
  const targetTableIndexByBaseTableIndex = new Map<number, number>();
  const structuralInsertionBoundaryByChange = new Map<
    ResolvedDocxInsertionStructuralChange,
    ResolvedDocxStructuralInsertionBoundary
  >();
  const tableGeometryPairingRecords: {
    readonly pairing: TableGeometryPairing;
    readonly sequence: number;
  }[] = [];
  const tableGeometrySequenceByPairing = new Map<string, number>();
  for (const [sequence, event] of payload.comparison.events.entries()) {
    work.eventVisits += 1;
    eventSequence.set(event, sequence);
    const eventIndex = sequence;
    switch (event.type) {
      case "unchanged":
      case "modified":
      case "formatting":
        registerTwoSidedRelation(event.relation, sequence);
        break;
      case "movedFrom": {
        const { relation } = event.move;
        moveRelations.add(relation);
        alignmentsWithParagraphStructureChange.add(relation.baseContainerAlignment);
        registerRelationOwnership(relation);
        registerBaseOccurrence(relation.base.block, relation.baseContainerAlignment, sequence);
        break;
      }
      case "movedTo": {
        const { relation } = event.move;
        moveRelations.add(relation);
        alignmentsWithParagraphStructureChange.add(relation.revisedContainerAlignment);
        registerRelationOwnership(relation);
        registerTargetOccurrence(
          relation.revised.block,
          relation.revisedContainerAlignment,
          sequence,
        );
        break;
      }
      case "split":
      case "merge":
        for (const relation of event.relations) {
          alignmentsWithParagraphStructureChange.add(relation.baseContainerAlignment);
          alignmentsWithParagraphStructureChange.add(relation.revisedContainerAlignment);
        }
        registerTwoSidedRelation(event.relations[0], sequence);
        registerTwoSidedRelation(event.relations[1], sequence);
        registerTwoSidedRelation(event.separator, sequence);
        break;
      case "inserted":
        alignmentsWithParagraphStructureChange.add(event.containerAlignment);
        registerTargetOccurrence(event.block, event.containerAlignment, sequence);
        break;
      case "deleted":
        alignmentsWithParagraphStructureChange.add(event.containerAlignment);
        registerBaseOccurrence(event.block, event.containerAlignment, sequence);
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
        switch (event.change.type) {
          case "table-row-delete":
            baseTablesWithRowStructureChange.add(event.change.tableIndex);
            break;
          case "table-row-insert":
            targetTablesWithRowStructureChange.add(event.change.tableIndex);
            break;
          case "table-column-delete":
            baseTablesWithColumnStructureChange.add(event.change.tableIndex);
            break;
          case "table-column-insert":
            targetTablesWithColumnStructureChange.add(event.change.tableIndex);
            break;
          case "table-delete":
          case "table-insert":
            break;
          default: {
            const unreachable: never = event.change;
            return panic("Unhandled structural change while indexing ownership", {
              change: unreachable,
            });
          }
        }
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

  const baseCoordinateTargets = new Map<string, string>();
  const targetCoordinateBases = new Map<string, string>();
  for (const alignment of pairedContainerAlignments) {
    work.alignmentVisits += 1;
    if (alignment.base.type === "body") {
      if (alignment.revised.type !== "body") {
        return panic("A paired container alignment crosses body and table-cell ownership");
      }
      continue;
    }
    if (alignment.revised.type !== "tableCell") {
      return panic("A paired container alignment crosses table-cell and body ownership");
    }
    const base = Object.freeze({
      tableIndex: alignment.base.table.tableIndex,
      rowIndex: alignment.base.table.rowIndex,
      cellIndex: alignment.base.table.cellIndex,
    });
    const target = Object.freeze({
      tableIndex: alignment.revised.table.tableIndex,
      rowIndex: alignment.revised.table.rowIndex,
      cellIndex: alignment.revised.table.cellIndex,
    });
    const baseKey = tableCellCoordinateKey(base);
    const targetKey = tableCellCoordinateKey(target);
    if (baseCoordinateTargets.has(baseKey) || targetCoordinateBases.has(targetKey)) {
      return panic("A table-cell occurrence belongs to two canonical container alignments", {
        base,
        target,
      });
    }
    baseCoordinateTargets.set(baseKey, targetKey);
    targetCoordinateBases.set(targetKey, baseKey);
    const priorTargetTable = targetTableIndexByBaseTableIndex.get(base.tableIndex);
    const priorBaseTable = baseTableIndexByTargetTableIndex.get(target.tableIndex);
    if (
      (priorTargetTable !== undefined && priorTargetTable !== target.tableIndex) ||
      (priorBaseTable !== undefined && priorBaseTable !== base.tableIndex)
    ) {
      return panic("A table occurrence belongs to two canonical table alignments", {
        baseTableIndex: base.tableIndex,
        targetTableIndex: target.tableIndex,
      });
    }
    targetTableIndexByBaseTableIndex.set(base.tableIndex, target.tableIndex);
    baseTableIndexByTargetTableIndex.set(target.tableIndex, base.tableIndex);
    const sequence =
      revisedSequenceByAlignment.get(alignment) ?? baseSequenceByAlignment.get(alignment);
    if (sequence === undefined) {
      return panic("A paired table-cell alignment has no canonical stream occurrence", {
        alignmentId: alignment.id,
      });
    }
    tableGeometryPairingRecords.push({
      pairing: Object.freeze({ base, target }),
      sequence,
    });
  }
  tableGeometryPairingRecords.sort(
    (left, right) =>
      left.sequence - right.sequence ||
      compareTableCellCoordinates(left.pairing.target, right.pairing.target) ||
      compareTableCellCoordinates(left.pairing.base, right.pairing.base),
  );
  const tableGeometryPairings = Object.freeze(
    tableGeometryPairingRecords.map(({ pairing, sequence }) => {
      tableGeometrySequenceByPairing.set(tableGeometryPairingKey(pairing), sequence);
      return pairing;
    }),
  );

  const structurallyOwnedTablePlacementRelations = new Set<FolioContentPairRelation>();
  for (const relation of relations) {
    work.relationVisits += 1;
    if (!relation.blockChanges.some(({ field }) => field === "table")) continue;
    if (moveRelations.has(relation)) {
      structurallyOwnedTablePlacementRelations.add(relation);
      continue;
    }
    const alignment = relation.baseContainerAlignment;
    if (
      alignment !== relation.revisedContainerAlignment ||
      alignment.type !== "paired" ||
      alignment.base.type !== "tableCell" ||
      alignment.revised.type !== "tableCell"
    ) {
      continue;
    }
    const base = relation.base.block.table;
    const target = relation.revised.block.table;
    if (
      !base ||
      !target ||
      !tableLocationMatchesOccurrence(base, alignment.base) ||
      !tableLocationMatchesOccurrence(target, alignment.revised) ||
      baseTableIndexByTargetTableIndex.get(target.tableIndex) !== base.tableIndex ||
      base.columnSpan !== target.columnSpan ||
      base.rowSpan !== target.rowSpan
    ) {
      continue;
    }
    const rowPlacementOwned =
      base.rowIndex === target.rowIndex ||
      baseTablesWithRowStructureChange.has(base.tableIndex) ||
      targetTablesWithRowStructureChange.has(target.tableIndex);
    const columnPlacementOwned =
      (base.cellIndex === target.cellIndex && base.gridColumnIndex === target.gridColumnIndex) ||
      baseTablesWithColumnStructureChange.has(base.tableIndex) ||
      targetTablesWithColumnStructureChange.has(target.tableIndex);
    const paragraphPlacementOwned =
      base.paragraphIndex === target.paragraphIndex ||
      alignmentsWithParagraphStructureChange.has(alignment);
    if (rowPlacementOwned && columnPlacementOwned && paragraphPlacementOwned) {
      structurallyOwnedTablePlacementRelations.add(relation);
    }
  }

  const structuralEventIndexes = new Map<
    FolioContentStructuralChange,
    readonly [number, ...number[]]
  >();
  const firstStructuralChangeByEventIndex = new Map<number, FolioContentStructuralChange>();
  for (const [change, indexes] of mutableStructuralEventIndexes) {
    work.structuralChangeVisits += 1;
    const first = indexes.at(0) ?? panic("A structural change has no canonical event occurrence");
    structuralEventIndexes.set(change, Object.freeze([first, ...indexes.slice(1)]));
    firstStructuralChangeByEventIndex.set(first, change);
    if (change.type !== "table-column-insert") continue;
    const anchorBlock =
      baseBlockById.get(change.anchor.blockId) ??
      panic("A table-column insertion lost its canonical base anchor", {
        blockId: change.anchor.blockId,
      });
    const pairedBaseTableIndex = baseTableIndexByTargetTableIndex.get(change.tableIndex);
    if (
      anchorBlock.table === undefined ||
      pairedBaseTableIndex === undefined ||
      anchorBlock.table.tableIndex !== pairedBaseTableIndex
    ) {
      return panic("A table-column insertion crosses canonical table ownership", {
        blockId: change.anchor.blockId,
        targetTableIndex: change.tableIndex,
      });
    }
    structuralInsertionBoundaryByChange.set(
      change,
      Object.freeze({
        source: resolvedDocxSourceOperand(payload.baseSnapshot, anchorBlock),
        position: change.anchor.position,
      }),
    );
  }

  let nextBaseBlock: FolioContentBlock | null = null;
  const nextBaseTableBlockByTableIndex = new Map<number, FolioContentBlock>();
  for (let eventIndex = payload.comparison.events.length - 1; eventIndex >= 0; eventIndex--) {
    work.anchorEventVisits += 1;
    const change = firstStructuralChangeByEventIndex.get(eventIndex);
    if (change?.type === "table-insert") {
      const anchor = nextBaseBlock ?? trailingBodyBlock;
      if (anchor) {
        structuralInsertionBoundaryByChange.set(
          change,
          Object.freeze({
            source: resolvedDocxSourceOperand(payload.baseSnapshot, anchor),
            position: nextBaseBlock === null ? "after" : "before",
          }),
        );
      }
    } else if (change?.type === "table-row-insert") {
      const pairedBaseTableIndex = baseTableIndexByTargetTableIndex.get(change.tableIndex);
      const anchor =
        pairedBaseTableIndex === undefined
          ? undefined
          : nextBaseTableBlockByTableIndex.get(pairedBaseTableIndex);
      if (anchor) {
        structuralInsertionBoundaryByChange.set(
          change,
          Object.freeze({
            source: resolvedDocxSourceOperand(payload.baseSnapshot, anchor),
            position: "before",
          }),
        );
      }
    }
    const event = payload.comparison.events[eventIndex];
    if (!event) continue;
    const baseBlock = firstBaseBlockOfEvent(event);
    if (baseBlock) nextBaseBlock = baseBlock;
    const tableBlock = baseTableBlockOfEvent(event);
    if (tableBlock?.table) {
      nextBaseTableBlockByTableIndex.set(tableBlock.table.tableIndex, tableBlock);
    }
  }
  const previousBaseTableBlockByTableIndex = new Map<number, FolioContentBlock>();
  for (const [eventIndex, event] of payload.comparison.events.entries()) {
    work.anchorEventVisits += 1;
    const change = firstStructuralChangeByEventIndex.get(eventIndex);
    if (change?.type === "table-row-insert") {
      const pairedBaseTableIndex = baseTableIndexByTargetTableIndex.get(change.tableIndex);
      const anchor =
        pairedBaseTableIndex === undefined
          ? undefined
          : previousBaseTableBlockByTableIndex.get(pairedBaseTableIndex);
      if (anchor) {
        structuralInsertionBoundaryByChange.set(
          change,
          Object.freeze({
            source: resolvedDocxSourceOperand(payload.baseSnapshot, anchor),
            position: "after",
          }),
        );
      }
    }
    const tableBlock = baseTableBlockOfEvent(event);
    if (tableBlock?.table) {
      previousBaseTableBlockByTableIndex.set(tableBlock.table.tableIndex, tableBlock);
    }
  }
  const baseTerminalBlocks = Object.freeze(
    [...baseTerminalBlockByContainer]
      .map(([alignment, block]) => Object.freeze({ alignment, block }))
      .toSorted(
        (left, right) =>
          (basePositionByBlock.get(left.block) ?? 0) - (basePositionByBlock.get(right.block) ?? 0),
      ),
  );
  const terminalComponentMemberByBlock = new Map(
    baseTerminalBlocks.map(({ block }) => [
      block,
      Object.freeze({ [RESOLVED_DOCX_TABLE_COMPONENT_MEMBER_BRAND]: true as const }),
    ]),
  );
  const frozenWork = Object.freeze({
    ...work,
    total:
      work.eventVisits +
      work.alignmentVisits +
      work.relationVisits +
      work.baseBlockVisits +
      work.revisedBlockVisits +
      work.anchorEventVisits +
      work.structuralChangeVisits,
  });
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
    baseContainerByBlockId,
    baseBlockById,
    basePositionByBlock,
    baseTerminalBlockByContainer,
    targetTerminalBlockByContainer,
    baseTerminalBlocks,
    targetTerminalOuterTableIndex: targetBlocks.at(-1)?.table?.outerTableIndex,
    baseTableIndexByTargetTableIndex,
    structuralInsertionBoundaryByChange,
    structurallyOwnedTablePlacementRelations,
    tableGeometryPairings,
    tableGeometrySequenceByPairing,
    baseTableComponentMemberByTableIndex,
    targetTableComponentMemberByTableIndex,
    terminalComponentMemberByBlock,
    work: frozenWork,
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
  if (!container || container.base.end !== "paragraph") {
    return panic("A trailing DOCX deletion operand must remain in one canonical container");
  }
  const baseBlocks = resolvedDocxContentBlocks(baseSnapshot);
  const eventPositions = ownedEvents.map(({ event }) => index.basePositionByBlock.get(event.block));
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
  const terminalBlock = index.baseTerminalBlockByContainer.get(container);
  if (
    !chainStartBlock ||
    index.baseContainerByBlock.get(chainStartBlock) !== container ||
    terminalBlock !== ownedEvents.at(-1)?.event.block
  ) {
    return panic("A trailing DOCX deletion operand must be contiguous and terminal");
  }
  const chainStart = resolvedDocxSourceOperand(baseSnapshot, chainStartBlock);
  const targetBlock =
    container.type === "paired" ? index.targetTerminalBlockByContainer.get(container) : undefined;
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
  const baseTerminal = index.baseTerminalBlockByContainer.get(baseContainer);
  const targetTerminal = index.targetTerminalBlockByContainer.get(baseContainer);
  if (
    baseTerminal !== deletedPayload.event.block ||
    targetTerminal !== insertedPayload.event.block
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

const requireExactStructuralAnchor = (
  input: Extract<
    ResolvedDocxTableStructureOperandInput,
    { readonly type: "insertTable" | "insertTableRow" | "insertTableColumn" }
  >,
  comparison: ResolvedDocxStoryComparison,
  payload: ResolvedDocxStoryComparisonPayload,
): void => {
  resolvedDocxSourceOperandBlock(input.anchor.source, payload.baseSnapshot);
  const expected = comparisonIndexOf(comparison).structuralInsertionBoundaryByChange.get(
    input.change,
  );
  if (
    expected === undefined ||
    expected.source !== input.anchor.source ||
    expected.position !== input.anchor.position
  ) {
    return panic("A DOCX table-structure anchor must equal its canonical insertion boundary");
  }
};

const ownedTerminalTableCarrier = (
  carrier: ResolvedDocxTerminalTableCarrier,
): ResolvedDocxTerminalTableCarrier =>
  Object.freeze({ source: carrier.source, event: carrier.event });

const ownedTableReplacementOwner = (
  owner: ResolvedDocxTableReplacementOwner,
): ResolvedDocxTableReplacementOwner => {
  switch (owner.type) {
    case "canonical-replacement":
      return Object.freeze({ type: "canonical-replacement", replacement: owner.replacement });
    case "structural-pair":
      return Object.freeze({
        type: "structural-pair",
        deleted: owner.deleted,
        inserted: owner.inserted,
      });
    default: {
      const unreachable: never = owner;
      return panic("Unhandled DOCX table replacement owner", { owner: unreachable });
    }
  }
};

const ownedStructuralInsertionBoundary = ({
  source,
  position,
}: ResolvedDocxStructuralInsertionBoundary): ResolvedDocxStructuralInsertionBoundary =>
  Object.freeze({ source, position });

const ownedTableStructurePayload = (
  input: ResolvedDocxTableStructureOperandPayload,
): ResolvedDocxTableStructureOperandPayload => {
  switch (input.type) {
    case "insertTable":
      return Object.freeze({
        type: "insertTable",
        anchor: ownedStructuralInsertionBoundary(input.anchor),
        change: input.change,
        ...(input.terminalCarrier !== undefined && {
          terminalCarrier: ownedTerminalTableCarrier(input.terminalCarrier),
        }),
      });
    case "deleteTable":
      return Object.freeze({ type: "deleteTable", source: input.source, change: input.change });
    case "replaceTable":
      return Object.freeze({
        type: "replaceTable",
        source: input.source,
        owner: ownedTableReplacementOwner(input.owner),
        ...(input.terminalCarrier !== undefined && {
          terminalCarrier: ownedTerminalTableCarrier(input.terminalCarrier),
        }),
      });
    case "insertTableRow":
      return Object.freeze({
        type: "insertTableRow",
        anchor: ownedStructuralInsertionBoundary(input.anchor),
        change: input.change,
      });
    case "deleteTableRow":
      return Object.freeze({ type: "deleteTableRow", source: input.source, change: input.change });
    case "insertTableColumn":
      return Object.freeze({
        type: "insertTableColumn",
        anchor: ownedStructuralInsertionBoundary(input.anchor),
        change: input.change,
      });
    case "deleteTableColumn":
      return Object.freeze({
        type: "deleteTableColumn",
        source: input.source,
        change: input.change,
      });
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX table-structure payload", { input: unreachable });
    }
  }
};

const tableComponentMember = ({
  comparison,
  side,
  tableIndex,
}: {
  readonly comparison: ResolvedDocxStoryComparison;
  readonly side: "base" | "target";
  readonly tableIndex: number;
}): ResolvedDocxTableComponentMember => {
  const index = comparisonIndexOf(comparison);
  const member =
    side === "base"
      ? index.baseTableComponentMemberByTableIndex.get(tableIndex)
      : index.targetTableComponentMemberByTableIndex.get(tableIndex);
  return (
    member ??
    panic("A DOCX table operand lost its canonical outer-table ownership", {
      side,
      tableIndex,
    })
  );
};

const tableComponentMemberForSource = (
  comparison: ResolvedDocxStoryComparison,
  source: ResolvedDocxSourceOperand,
): ResolvedDocxTableComponentMember | null => {
  const { baseSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
  const table = resolvedDocxSourceOperandBlock(source, baseSnapshot).table;
  return table
    ? tableComponentMember({ comparison, side: "base", tableIndex: table.tableIndex })
    : null;
};

const tableComponentMembersForStructure = (
  comparison: ResolvedDocxStoryComparison,
  payload: ResolvedDocxTableStructureOperandPayload,
): readonly ResolvedDocxTableComponentMember[] => {
  const members: ResolvedDocxTableComponentMember[] = [];
  const add = (member: ResolvedDocxTableComponentMember | null): void => {
    if (member && !members.includes(member)) members.push(member);
  };
  const addTarget = (tableIndex: number): void =>
    add(tableComponentMember({ comparison, side: "target", tableIndex }));
  switch (payload.type) {
    case "insertTable":
      add(tableComponentMemberForSource(comparison, payload.anchor.source));
      addTarget(payload.change.tableIndex);
      break;
    case "deleteTable":
    case "deleteTableRow":
    case "deleteTableColumn":
      add(tableComponentMemberForSource(comparison, payload.source));
      break;
    case "replaceTable":
      add(tableComponentMemberForSource(comparison, payload.source));
      addTarget(
        payload.owner.type === "canonical-replacement"
          ? payload.owner.replacement.revisedTableIndex
          : payload.owner.inserted.tableIndex,
      );
      break;
    case "insertTableRow":
    case "insertTableColumn":
      add(tableComponentMemberForSource(comparison, payload.anchor.source));
      addTarget(payload.change.tableIndex);
      break;
    default: {
      const unreachable: never = payload;
      return panic("Unhandled table operand while binding component ownership", {
        payload: unreachable,
      });
    }
  }
  const carrier =
    payload.type === "insertTable" || payload.type === "replaceTable"
      ? payload.terminalCarrier
      : undefined;
  if (carrier) {
    add(
      comparisonIndexOf(comparison).terminalComponentMemberByBlock.get(carrier.event.block) ??
        panic("A terminal table carrier lost its canonical component ownership"),
    );
  }
  if (members.length === 0) {
    return panic("A DOCX table operand has no canonical component ownership");
  }
  return Object.freeze(members);
};

const issueTableStructureOperand = (
  comparison: ResolvedDocxStoryComparison,
  input: ResolvedDocxTableStructureOperandPayload,
): ResolvedDocxTableStructureOperand => {
  const payload = ownedTableStructurePayload(input);
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
    value: payload,
  });
  componentMembersByTableOperand.set(operand, {
    comparison,
    value: tableComponentMembersForStructure(comparison, payload),
  });
  return operand;
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
      break;
    }
    case "insertTableRow":
    case "insertTableColumn": {
      if (!index.structuralChanges.has(structuralChangeOf(input))) {
        return panic("A DOCX table-structure operand must name an exact canonical change");
      }
      requireExactStructuralAnchor(input, comparison, payload);
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
      requireExactStructuralAnchor(input, comparison, payload);
      break;
    }
    default: {
      const unreachable: never = input;
      return panic("Unhandled DOCX table-structure operand input", { input: unreachable });
    }
  }
  switch (input.type) {
    case "insertTable":
      return issueTableStructureOperand(comparison, {
        type: "insertTable",
        anchor: input.anchor,
        change: input.change,
      });
    case "deleteTable":
      return issueTableStructureOperand(comparison, {
        type: "deleteTable",
        source: input.source,
        change: input.change,
      });
    case "replaceTable":
      return issueTableStructureOperand(comparison, {
        type: "replaceTable",
        source: input.source,
        owner: input.owner,
      });
    case "insertTableRow":
      return issueTableStructureOperand(comparison, {
        type: "insertTableRow",
        anchor: input.anchor,
        change: input.change,
      });
    case "deleteTableRow":
      return issueTableStructureOperand(comparison, {
        type: "deleteTableRow",
        source: input.source,
        change: input.change,
      });
    case "insertTableColumn":
      return issueTableStructureOperand(comparison, {
        type: "insertTableColumn",
        anchor: input.anchor,
        change: input.change,
      });
    case "deleteTableColumn":
      return issueTableStructureOperand(comparison, {
        type: "deleteTableColumn",
        source: input.source,
        change: input.change,
      });
    default: {
      const unreachable: never = input;
      return panic("Unhandled validated DOCX table-structure operand", { input: unreachable });
    }
  }
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

const ownedTableCoordinate = (
  coordinate: TableGeometryPairing["base"],
): Readonly<TableGeometryPairing["base"]> =>
  Object.freeze({
    tableIndex: coordinate.tableIndex,
    rowIndex: coordinate.rowIndex,
    cellIndex: coordinate.cellIndex,
  });

const ownedTableGeometryUnsupportedIssue = (
  issue: TableGeometryUnsupportedIssue,
): TableGeometryUnsupportedIssue => {
  switch (issue.reason) {
    case "invalid-limit":
      return Object.freeze({ reason: "invalid-limit", limit: issue.limit, actual: issue.actual });
    case "limit-exceeded":
      return Object.freeze({
        reason: "limit-exceeded",
        limit: issue.limit,
        maximum: issue.maximum,
        actual: issue.actual,
      });
    case "invalid-coordinate":
      return Object.freeze({
        reason: "invalid-coordinate",
        side: issue.side,
        coordinate: ownedTableCoordinate(issue.coordinate),
      });
    case "invalid-table-index":
    case "duplicate-table-index":
      return Object.freeze({
        reason: issue.reason,
        side: issue.side,
        tableIndex: issue.tableIndex,
      });
    case "invalid-table-position":
    case "duplicate-table-position":
      return Object.freeze({ reason: issue.reason, position: issue.position });
    case "duplicate-pairing":
      return Object.freeze({
        reason: "duplicate-pairing",
        side: issue.side,
        coordinate: ownedTableCoordinate(issue.coordinate),
      });
    case "conflicting-table-pairing":
    case "conflicting-row-pairing":
      return Object.freeze({
        reason: issue.reason,
        side: issue.side,
        coordinate: ownedTableCoordinate(issue.coordinate),
      });
    case "missing-table":
    case "missing-row":
    case "missing-cell":
    case "unexpected-node-role":
      return Object.freeze({
        reason: issue.reason,
        side: issue.side,
        scope: issue.scope,
        coordinate: ownedTableCoordinate(issue.coordinate),
      });
    case "structural-cell-mismatch":
      return Object.freeze({
        reason: "structural-cell-mismatch",
        property: issue.property,
        base: ownedTableCoordinate(issue.base),
        target: ownedTableCoordinate(issue.target),
      });
    case "non-reconstructable-structure-change":
      return Object.freeze({
        reason: "non-reconstructable-structure-change",
        scope: issue.scope,
        property: issue.property,
        base: ownedTableCoordinate(issue.base),
        target: ownedTableCoordinate(issue.target),
      });
    case "non-reconstructable-property-change":
      return Object.freeze({
        reason: "non-reconstructable-property-change",
        reconstruction: issue.reconstruction,
        scope: issue.scope,
        base: ownedTableCoordinate(issue.base),
        target: ownedTableCoordinate(issue.target),
      });
    case "invalid-property-payload":
      return Object.freeze({
        reason: "invalid-property-payload",
        scope: issue.scope,
        base: ownedTableCoordinate(issue.base),
        target: ownedTableCoordinate(issue.target),
      });
    default: {
      const unreachable: never = issue;
      return panic("Unhandled table geometry issue while capturing comparison ownership", {
        issue: unreachable,
      });
    }
  }
};

const ownedTableFormatPayload = (
  payload: ResolvedDocxTableFormatOperandPayload,
): ResolvedDocxTableFormatOperandPayload => {
  switch (payload.status) {
    case "ready":
      return Object.freeze({
        status: "ready",
        program: payload.program,
        changes: Object.freeze(
          payload.changes.map(({ change, sequence }) => Object.freeze({ change, sequence })),
        ),
      });
    case "unsupported":
      if (payload.issue.reason === "unprojected-table-structure") {
        return Object.freeze({
          status: "unsupported",
          issue: Object.freeze({
            reason: "unprojected-table-structure",
            side: payload.issue.side,
          }),
        });
      }
      return Object.freeze({
        status: "unsupported",
        issue: Object.freeze({
          reason: "unrepresentable-table-geometry",
          issue: ownedTableGeometryUnsupportedIssue(payload.issue.issue),
        }),
      });
    default: {
      const unreachable: never = payload;
      return panic("Unhandled table-format payload while capturing ownership", {
        payload: unreachable,
      });
    }
  }
};

const issueTableFormatOperand = (
  comparison: ResolvedDocxStoryComparison,
  payload: ResolvedDocxTableFormatOperandPayload,
  members: readonly ResolvedDocxTableComponentMember[],
): ResolvedDocxTableFormatOperand => {
  const operand = Object.freeze({
    type: "matchTableFormatting" as const,
    [RESOLVED_DOCX_TABLE_FORMAT_OPERAND_BRAND]: true as const,
  });
  payloadByTableFormatOperand.set(operand, {
    comparison,
    value: ownedTableFormatPayload(payload),
  });
  componentMembersByTableOperand.set(operand, {
    comparison,
    value: Object.freeze([...members]),
  });
  return operand;
};

/**
 * Resolve paired table properties once. Unchanged tables produce no semantic
 * operation; changed and unsupported projections retain one exact owner.
 */
export const resolvedDocxTableFormatOperands = (
  comparison: ResolvedDocxStoryComparison,
): readonly ResolvedDocxTableFormatOperand[] => {
  const { baseSnapshot, targetSnapshot } = resolvedDocxStoryComparisonPayload(comparison);
  const pairings = resolvedDocxTableGeometryPairings(comparison);
  if (pairings.length === 0) return Object.freeze([]);
  const index = comparisonIndexOf(comparison);
  type FormatComponent = {
    readonly members: ResolvedDocxTableComponentMember[];
    readonly pairings: TableGeometryPairing[];
    issue?: Extract<
      ResolvedDocxTableFormatOperandPayload,
      { readonly status: "unsupported" }
    >["issue"];
  };
  const roots = pairings.map((_pairing, pairingIndex) => pairingIndex);
  const rootOf = (pairingIndex: number): number => {
    let root = pairingIndex;
    while (roots[root] !== root) root = roots[root] ?? root;
    let current = pairingIndex;
    while (roots[current] !== root) {
      const next = roots[current] ?? root;
      roots[current] = root;
      current = next;
    }
    return root;
  };
  const join = (left: number, right: number): void => {
    const leftRoot = rootOf(left);
    const rightRoot = rootOf(right);
    if (leftRoot === rightRoot) return;
    roots[Math.max(leftRoot, rightRoot)] = Math.min(leftRoot, rightRoot);
  };
  const firstPairingByMember = new Map<ResolvedDocxTableComponentMember, number>();
  const membersByPairing: readonly (readonly ResolvedDocxTableComponentMember[])[] = pairings.map(
    (pairing, pairingIndex) => {
      const baseMember = tableComponentMember({
        comparison,
        side: "base",
        tableIndex: pairing.base.tableIndex,
      });
      const targetMember = tableComponentMember({
        comparison,
        side: "target",
        tableIndex: pairing.target.tableIndex,
      });
      const members = Object.freeze([baseMember, targetMember]);
      for (const member of members) {
        const firstPairing = firstPairingByMember.get(member);
        if (firstPairing === undefined) firstPairingByMember.set(member, pairingIndex);
        else join(firstPairing, pairingIndex);
      }
      return members;
    },
  );
  const componentsByRoot = new Map<number, FormatComponent>();
  for (const [pairingIndex, pairing] of pairings.entries()) {
    const root = rootOf(pairingIndex);
    const component = componentsByRoot.get(root) ?? { members: [], pairings: [] };
    for (const member of membersByPairing[pairingIndex] ?? []) {
      if (!component.members.includes(member)) component.members.push(member);
    }
    component.pairings.push(pairing);
    componentsByRoot.set(root, component);
  }
  const components = [...componentsByRoot.values()];
  const baseTables = storyTablesOf(resolvedDocxOperationSnapshot(baseSnapshot));
  const baseByIndex = new Map(baseTables.map((table) => [table.index, table.node]));
  const targetTables = resolvedDocxTableNodes(targetSnapshot);
  for (const component of components) {
    const inspectedPairs = new Set<string>();
    for (const { base, target } of component.pairings) {
      const pair = `${String(base.tableIndex)}:${String(target.tableIndex)}`;
      if (inspectedPairs.has(pair)) continue;
      inspectedPairs.add(pair);
      const baseTable = baseByIndex.get(base.tableIndex);
      const targetTable = targetTables.get(target.tableIndex);
      if (baseTable && tableHasHiddenRows(baseTable)) {
        component.issue = { reason: "unprojected-table-structure", side: "source" };
        break;
      }
      if (targetTable && tableHasHiddenRows(targetTable)) {
        component.issue = { reason: "unprojected-table-structure", side: "target" };
        break;
      }
    }
  }

  const preflight = preflightTableGeometryComponents({
    baseTables,
    targetTables,
    components: components.map((component) => ({
      component,
      pairings: component.pairings,
    })),
    tableGridChanges: "defer-to-table-structure",
  });
  if (preflight.status === "unsupported") {
    return Object.freeze(
      components.map((component) =>
        issueTableFormatOperand(
          comparison,
          {
            status: "unsupported",
            issue: { reason: "unrepresentable-table-geometry", issue: preflight.issue },
          },
          component.members,
        ),
      ),
    );
  }
  const operands: ResolvedDocxTableFormatOperand[] = [];
  for (const { component, result } of preflight.components) {
    if (component.issue) {
      operands.push(
        issueTableFormatOperand(
          comparison,
          { status: "unsupported", issue: component.issue },
          component.members,
        ),
      );
      continue;
    }
    if (result.status === "unsupported") {
      operands.push(
        issueTableFormatOperand(
          comparison,
          {
            status: "unsupported",
            issue: { reason: "unrepresentable-table-geometry", issue: result.issue },
          },
          component.members,
        ),
      );
      continue;
    }
    const semanticOccurrences = tableGeometryProgramSemanticChangeOccurrences(result.program);
    if (
      semanticOccurrences.length === 0 &&
      tableGeometryProgramTableGridTransitions(result.program).length === 0
    ) {
      continue;
    }
    const changes = semanticOccurrences.map(({ change, owner }) => {
      const sequence = index.tableGeometrySequenceByPairing.get(tableGeometryPairingKey(owner));
      if (sequence === undefined) {
        return panic("A table-format change lost its canonical paired-cell event");
      }
      return Object.freeze({ change, sequence });
    });
    operands.push(
      issueTableFormatOperand(
        comparison,
        { status: "ready", program: result.program, changes },
        component.members,
      ),
    );
  }
  return Object.freeze(operands);
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

/** Resolve the comparison-owned scopes used to partition table execution. */
const resolvedDocxTableOperandComponentMembers = (
  operand: ResolvedDocxTableStructureOperand | ResolvedDocxTableFormatOperand,
  comparison: ResolvedDocxStoryComparison,
): readonly ResolvedDocxTableComponentMember[] =>
  requireComparisonOwned(
    componentMembersByTableOperand.get(operand),
    comparison,
    "DOCX table component membership",
  );

/** Derive the total connected-component partition from comparison-owned scopes. */
export const resolvedDocxTableComponents = ({
  comparison,
  operands,
}: {
  readonly comparison: ResolvedDocxStoryComparison;
  readonly operands: readonly (
    | ResolvedDocxTableStructureOperand
    | ResolvedDocxTableFormatOperand
  )[];
}): readonly ResolvedDocxTableComponent[] => {
  if (new Set(operands).size !== operands.length) {
    return panic("A DOCX table operand cannot occur twice in a component partition");
  }
  const parents = operands.map((_operand, operandIndex) => operandIndex);
  const rootOf = (operandIndex: number): number => {
    let root = operandIndex;
    while (parents[root] !== root) root = parents[root] ?? root;
    let current = operandIndex;
    while (parents[current] !== root) {
      const next = parents[current] ?? root;
      parents[current] = root;
      current = next;
    }
    return root;
  };
  const join = (left: number, right: number): void => {
    const leftRoot = rootOf(left);
    const rightRoot = rootOf(right);
    if (leftRoot === rightRoot) return;
    parents[Math.max(leftRoot, rightRoot)] = Math.min(leftRoot, rightRoot);
  };
  const firstOperandByMember = new Map<ResolvedDocxTableComponentMember, number>();
  for (const [operandIndex, operand] of operands.entries()) {
    for (const member of resolvedDocxTableOperandComponentMembers(operand, comparison)) {
      const firstOperand = firstOperandByMember.get(member);
      if (firstOperand === undefined) firstOperandByMember.set(member, operandIndex);
      else join(firstOperand, operandIndex);
    }
  }
  const operandsByRoot = new Map<
    number,
    (ResolvedDocxTableStructureOperand | ResolvedDocxTableFormatOperand)[]
  >();
  for (const [operandIndex, operand] of operands.entries()) {
    const root = rootOf(operandIndex);
    const owned = operandsByRoot.get(root) ?? [];
    owned.push(operand);
    operandsByRoot.set(root, owned);
  }
  return Object.freeze(
    [...operandsByRoot.values()].map((ownedOperands) => {
      const component = Object.freeze({
        [RESOLVED_DOCX_TABLE_COMPONENT_BRAND]: true as const,
      });
      operandsByTableComponent.set(component, {
        comparison,
        value: Object.freeze(ownedOperands),
      });
      return component;
    }),
  );
};

/** Resolve only a canonical component issued for this exact comparison. */
export const resolvedDocxTableComponentOperands = (
  component: ResolvedDocxTableComponent,
  comparison: ResolvedDocxStoryComparison,
): readonly (ResolvedDocxTableStructureOperand | ResolvedDocxTableFormatOperand)[] =>
  requireComparisonOwned(
    operandsByTableComponent.get(component),
    comparison,
    "DOCX table component",
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
  if (
    carrier.kind !== "paragraph" ||
    carrier.text.length !== 0 ||
    comparisonIndexOf(comparison).targetTerminalOuterTableIndex !==
      tableReplacementIndexes(payload.owner).target
  ) {
    return panic("A terminal table replacement must own an empty final carrier and target table");
  }
  const resolved = issueTableStructureOperand(comparison, {
    type: "replaceTable",
    source: payload.source,
    owner: payload.owner,
    terminalCarrier: { source: terminalCarrier, event },
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
  if (
    carrier.kind !== "paragraph" ||
    carrier.text.length !== 0 ||
    comparisonIndexOf(comparison).targetTerminalOuterTableIndex !== payload.change.tableIndex
  ) {
    return panic("A terminal table insertion must own an empty final carrier and target table");
  }
  const resolved = issueTableStructureOperand(comparison, {
    type: "insertTable",
    anchor: { source: terminalCarrier, position: "after" },
    change: payload.change,
    terminalCarrier: { source: terminalCarrier, event },
  });
  if (resolved.type !== "insertTable") {
    return panic("A terminal table insertion lost its structural discriminator");
  }
  return resolved;
};

/** Exact compiler-owned boundary for one canonical structural insertion. */
export const resolvedDocxStructuralInsertionBoundary = (
  comparison: ResolvedDocxStoryComparison,
  change: ResolvedDocxInsertionStructuralChange,
): ResolvedDocxStructuralInsertionBoundary | null => {
  const index = comparisonIndexOf(comparison);
  if (!index.structuralChanges.has(change)) {
    return panic("A structural insertion boundary must name an exact canonical change");
  }
  return index.structuralInsertionBoundaryByChange.get(change) ?? null;
};

/** Whether canonical structure or relocation owns this relation's table placement delta. */
export const resolvedDocxTablePlacementChangeIsOwned = (
  comparison: ResolvedDocxStoryComparison,
  relation: FolioContentPairRelation,
): boolean => {
  const index = comparisonIndexOf(comparison);
  if (!index.relations.has(relation)) {
    return panic("A table-placement ownership query must name an exact canonical relation");
  }
  return index.structurallyOwnedTablePlacementRelations.has(relation);
};

/** Canonical base-container ownership without exposing the mutable index map. */
export const resolvedDocxBaseContainerAlignmentForBlockId = (
  comparison: ResolvedDocxStoryComparison,
  blockId: string,
): FolioContentBaseContainerAlignment | null =>
  comparisonIndexOf(comparison).baseContainerByBlockId.get(blockId) ?? null;

/** Canonical base block lookup without exposing the mutable index map. */
export const resolvedDocxBaseBlockForId = (
  comparison: ResolvedDocxStoryComparison,
  blockId: string,
): FolioContentBlock | null => comparisonIndexOf(comparison).baseBlockById.get(blockId) ?? null;

/** Last canonical base block of every indexed container, in document order. */
export const resolvedDocxBaseTerminalBlocks = (
  comparison: ResolvedDocxStoryComparison,
): readonly ResolvedDocxBaseTerminalBlock[] => comparisonIndexOf(comparison).baseTerminalBlocks;

/** Deterministic work accounting for the one-time comparison-index build. @internal */
export const resolvedDocxComparisonIndexWork = (
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxComparisonIndexWork => comparisonIndexOf(comparison).work;

/** Exact cell pairings derived once from this comparison's canonical relation graph. */
export const resolvedDocxTableGeometryPairings = (
  comparison: ResolvedDocxStoryComparison,
): readonly TableGeometryPairing[] => comparisonIndexOf(comparison).tableGeometryPairings;
