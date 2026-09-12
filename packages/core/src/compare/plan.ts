/**
 * Turn one completed neutral story comparison into the change list a caller
 * reads and the closed DOCX instruction program that realizes it.
 *
 * Pure: no parsing, no serialization, no clock. Alignment, moves, split/merge,
 * text segments, and formatting classification have already been completed by
 * the representation-neutral comparison core. This module only lowers that
 * immutable semantic product into tracked-document operations.
 */

import { panic, Result } from "better-result";

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import { canonicalJson } from "../utils/canonicalJson";
import type { FolioAIBlock, FolioAIEditSnapshot } from "../ai-edits/types";
import {
  docxParagraphPropertiesEqual,
  docxParagraphPropertiesFromBlock,
  docxParagraphPropertyChangeIsLowerable,
} from "../internal/compare/docx-paragraph-transport";
import {
  resolvedDocxBaseBlockForId,
  resolvedDocxBaseContainerAlignmentForBlockId,
  resolvedDocxBaseTerminalBlocks,
  resolvedDocxDeletedEventOperand,
  resolvedDocxDeletedEventOperandPayload,
  resolvedDocxInsertedEventOperand,
  resolvedDocxInsertedEventOperandPayload,
  resolvedDocxMergeEventOperand,
  resolvedDocxMergeEventOperandPayload,
  resolvedDocxMoveEventOperand,
  resolvedDocxMoveEventOperandPayload,
  resolvedDocxPairedEventOperand,
  resolvedDocxSplitEventOperand,
  resolvedDocxSplitEventOperandPayload,
  resolvedDocxStoryComparisonPayload,
  resolvedDocxStructuralInsertionBoundary,
  resolvedDocxTablePlacementChangeIsOwned,
  resolvedDocxTerminalReplacementOperand,
  resolvedDocxTableFormatOperand,
  resolvedDocxTableStructureOperand,
  resolvedDocxTableStructureOperandPayload,
  resolvedDocxTerminalTableInsertionOperand,
  resolvedDocxTerminalTableReplacementOperand,
  resolvedDocxTrailingDeletionOperand,
  type ResolvedDocxStoryComparison,
} from "../internal/compare/resolved-docx-story-comparison";
import {
  resolvedDocxHasExactAuthoredRuns,
  resolvedDocxOperationSnapshot,
  resolvedDocxSourceOperand,
  resolvedDocxSourceOperandBlock,
  resolvedDocxSourceOperandForBlockId,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "../internal/compare/resolved-docx-story-snapshot";
import {
  DocxComparisonProgram,
  type DocxComparisonOperationInput,
} from "../internal/compare/docx-program";
import {
  type FolioContentComparisonEvent,
  type FolioContentPairRelation,
  type FolioContentRangePairRelation,
  type FolioContentWholePairRelation,
} from "./content";
import type {
  FolioContentBaseContainerAlignment,
  FolioContentBlock,
  FolioContentContainerAlignment,
} from "./content-types";
import {
  CompareDocxLoweringError,
  CompareDocxOperationLimitError,
  type CompareUnsupportedPart,
} from "./types";

const isEmptyParagraphNode = (block: FolioAIBlock, snapshot: FolioAIEditSnapshot): boolean => {
  const anchor =
    snapshot.anchors[block.id] ??
    panic("A comparison snapshot block has no matching anchor", { blockId: block.id });
  return block.text === "" && block.table === undefined && anchor.to - anchor.from === 2;
};

type CompareUnsupportedEventType = Extract<
  CompareUnsupportedPart,
  { readonly eventType: unknown }
>["eventType"];

type TrailingDeletionOptions = {
  comparison: ResolvedDocxStoryComparison;
  baseResolvedSnapshot: ResolvedDocxStorySnapshot;
  baseSnapshot: FolioAIEditSnapshot;
  targetSnapshot: FolioAIEditSnapshot;
  unavailableTerminalCarrierIds: ReadonlySet<string>;
  /** The plan so far, rewritten around each container it empties. */
  operations: readonly DocxComparisonOperationInput[];
};

type TrailingDeletionPlan = {
  readonly operations: DocxComparisonOperationInput[];
};

/**
 * Where the plan's inserts land for one container's trailing deleted run, and
 * which paragraph the merge chain starts from.
 */
type TrailingRun = {
  /** The deleted blocks that end the container, carrier first. */
  blockIds: readonly string[];
  /** The last SURVIVING paragraph of the container, when its mark is free. */
  chainStart: FolioAIBlock | null;
  /** Indexes into the plan of the paragraph insertions placed in the run. */
  insertIndexes: readonly number[];
  /** Indexes into the plan of tables placed inside the run. */
  tableInsertIndexes: readonly number[];
};

/**
 * Rewrite the plan so a container's final paragraph mark survives the removal
 * of the paragraphs that end it.
 *
 * A deleted paragraph mark means "merge this paragraph into the following
 * one". A container's final paragraph has no following one, so the mark cannot
 * say it: the applier leaves that mark alone, and the removal has to be
 * expressed one paragraph earlier. The carrier is that final paragraph, and
 * every trailing removal resolves onto it.
 *
 * Nothing is inserted where the run was: the chain runs from the last
 * SURVIVING paragraph forward. Its mark goes, each removed paragraph's mark
 * goes with it, and the carrier loses its words and keeps its mark as the
 * paragraph the merged text lands in. A carrier that has no words to lose is
 * not deleted at all: its removal is entirely the marks in front of it, and an
 * operation that would write no revision belongs nowhere in the plan.
 *
 * Something IS inserted there: the inserted paragraphs stand where the removed
 * ones did, and the last of them lands INSIDE the carrier as inserted runs
 * before its kept mark, so the carrier is the target's last paragraph rather
 * than a blank line after it. The other inserted paragraphs carry inserted
 * marks of their own and sit before the carrier. The mark count then works out
 * on its own — one mark deleted for every removed paragraph but the carrier,
 * one inserted for every added paragraph but the last — so the chain is not
 * rotated as well.
 *
 * The carrier keeps its own mark either way, and a paragraph's properties live
 * on its mark, so the merged paragraph ends up with the carrier's. Those
 * properties therefore become the target's, written as `w:pPrChange` so
 * rejecting restores what the carrier had. That is bookkeeping for the merge
 * rather than an edit of its own, so it adds no entry to the change list: what
 * the reader is told is that paragraphs were removed and paragraphs added.
 */
const withTrailingDeletionRules = ({
  comparison,
  baseResolvedSnapshot,
  baseSnapshot,
  targetSnapshot,
  unavailableTerminalCarrierIds,
  operations,
}: TrailingDeletionOptions): TrailingDeletionPlan => {
  const plan = [...operations];
  const sourceBlockId = (source: ResolvedDocxSourceOperand): string =>
    resolvedDocxSourceOperandBlock(source, baseResolvedSnapshot).identity.id;
  const deletionIndexByBlockId = new Map<string, number>();
  const insertionsByAlignment = new Map<
    FolioContentContainerAlignment,
    {
      readonly index: number;
      readonly boundary: Extract<
        FolioContentComparisonEvent,
        { readonly type: "inserted" }
      >["boundary"];
    }[]
  >();
  const tableInsertIndexesByAnchor = new Map<string, number[]>();
  const tableInsertIndexes: number[] = [];
  const tableDeleteIndexes: number[] = [];
  const tableReplacementIndexes: number[] = [];
  // A block whose own mark the plan already moves is not a chain start: a
  // split has put a second paragraph after it, and a merge has spent its mark.
  const markedBlockIds = new Set<string>();
  for (const [index, operation] of plan.entries()) {
    switch (operation.type) {
      case "deleteParagraph": {
        const { event } = resolvedDocxDeletedEventOperandPayload(operation.event, comparison);
        deletionIndexByBlockId.set(event.block.identity.id, index);
        break;
      }
      case "insertParagraph": {
        const { event } = resolvedDocxInsertedEventOperandPayload(operation.event, comparison);
        if (event.boundary.type === "unanchoredContainer") {
          return panic("An unanchored insertion reached terminal-deletion rewriting");
        }
        const alignment = event.boundary.containerAlignment;
        const placed = insertionsByAlignment.get(alignment) ?? [];
        placed.push({ index, boundary: event.boundary });
        insertionsByAlignment.set(alignment, placed);
        break;
      }
      case "tableStructure": {
        const instruction = operation.operation;
        const payload = resolvedDocxTableStructureOperandPayload(instruction, comparison);
        switch (instruction.type) {
          case "insertTable": {
            if (payload.type !== "insertTable") {
              return panic("A table insertion operand lost its canonical payload");
            }
            const anchorId = sourceBlockId(payload.anchor.source);
            const placed = tableInsertIndexesByAnchor.get(anchorId) ?? [];
            placed.push(index);
            tableInsertIndexesByAnchor.set(anchorId, placed);
            tableInsertIndexes.push(index);
            break;
          }
          case "deleteTable":
            tableDeleteIndexes.push(index);
            break;
          case "replaceTable":
            tableReplacementIndexes.push(index);
            break;
          case "insertTableRow":
          case "deleteTableRow":
          case "insertTableColumn":
          case "deleteTableColumn":
            break;
          default: {
            const unreachable: never = instruction;
            return panic("Unhandled table operation in terminal-deletion rewriting", {
              instruction: unreachable,
            });
          }
        }
        break;
      }
      case "splitParagraph": {
        const { event } = resolvedDocxSplitEventOperandPayload(operation.event, comparison);
        markedBlockIds.add(event.relations[0].base.block.identity.id);
        break;
      }
      case "mergeParagraphs": {
        const { event } = resolvedDocxMergeEventOperandPayload(operation.event, comparison);
        markedBlockIds.add(event.relations[0].base.block.identity.id);
        break;
      }
      case "moveParagraph": {
        const { event } = resolvedDocxMoveEventOperandPayload(operation.event, comparison);
        markedBlockIds.add(event.move.relation.base.block.identity.id);
        if (event.move.sourceRemovalBoundary.type === "terminalPredecessor") {
          markedBlockIds.add(event.move.sourceRemovalBoundary.predecessor.identity.id);
        }
        break;
      }
      case "pairedBlock":
      case "deleteTrailingParagraphs":
      case "replaceTerminalParagraph":
      case "tableFormat":
        break;
      default: {
        const unreachable: never = operation;
        return panic("Unhandled operation in terminal-deletion rewriting", {
          operation: unreachable,
        });
      }
    }
  }
  if (deletionIndexByBlockId.size === 0) {
    return { operations: plan };
  }

  const blocks = baseSnapshot.blocks;
  const indexById = new Map(blocks.map((block, index) => [block.id, index]));
  /**
   * Walk back from a container's last block over the run of deleted
   * paragraphs it ends with. A block of another container in between — a
   * table between two body paragraphs, a table nested in a cell — ends the
   * walk: no mark joins across it.
   */
  const trailingRunOf = (
    alignment: FolioContentBaseContainerAlignment,
    carrier: FolioAIBlock,
  ): TrailingRun => {
    const blockIds: string[] = [];
    const insertIndexes: number[] = [];
    const runTableInsertIndexes: number[] = [];
    let chainStart: FolioAIBlock | null = null;
    let index =
      indexById.get(carrier.id) ??
      panic("A container's last block is not in the snapshot it came from", {
        blockId: carrier.id,
      });
    for (; index >= 0; index--) {
      const block = blocks[index];
      if (
        block === undefined ||
        resolvedDocxBaseContainerAlignmentForBlockId(comparison, block.id) !== alignment
      ) {
        break;
      }
      if (!deletionIndexByBlockId.has(block.id)) {
        chainStart = markedBlockIds.has(block.id) ? null : block;
        break;
      }
      blockIds.push(block.id);
      runTableInsertIndexes.push(...(tableInsertIndexesByAnchor.get(block.id) ?? []));
    }
    const trailingIds = new Set(blockIds);
    for (const insertion of insertionsByAlignment.get(alignment) ?? []) {
      if (insertion.boundary.type === "unanchoredContainer") {
        return panic("An unanchored insertion reached a trailing paragraph run");
      }
      const boundaryId = insertion.boundary.paragraph.identity.id;
      if (
        trailingIds.has(boundaryId) ||
        (insertion.boundary.type === "afterParagraph" && boundaryId === chainStart?.id)
      ) {
        insertIndexes.push(insertion.index);
      }
    }
    // The plan emits operations in target order, so the plan's own order is
    // the order the inserted paragraphs have to end up in. Numerically: the
    // walk collects them container-block by container-block, and the default
    // comparator would sort index 10 in front of index 2.
    return {
      blockIds,
      chainStart,
      insertIndexes: insertIndexes.toSorted((left, right) => left - right),
      tableInsertIndexes: runTableInsertIndexes.toSorted((left, right) => left - right),
    };
  };

  const terminalTargetTableIndex = targetSnapshot.blocks.at(-1)?.table?.outerTableIndex;
  const dropped = new Set<number>();
  const appended: DocxComparisonOperationInput[] = [];
  for (const { alignment, block: canonicalCarrier } of resolvedDocxBaseTerminalBlocks(comparison)) {
    const carrierIndex =
      indexById.get(canonicalCarrier.identity.id) ??
      panic("A canonical terminal block is absent from its operation snapshot", {
        blockId: canonicalCarrier.identity.id,
      });
    const carrier =
      blocks[carrierIndex] ??
      panic("A canonical terminal block lost its operation-snapshot position", {
        blockId: canonicalCarrier.identity.id,
      });
    if (alignment.base.end !== "paragraph") continue;
    const carrierDeletionIndex = deletionIndexByBlockId.get(carrier.id);
    if (carrierDeletionIndex === undefined) {
      continue;
    }
    if (unavailableTerminalCarrierIds.has(carrier.id)) {
      continue;
    }
    const run = trailingRunOf(alignment, carrier);
    if (run.tableInsertIndexes.length > 0) {
      const tableInsertIndex =
        run.tableInsertIndexes.length === 1 ? run.tableInsertIndexes.at(0) : undefined;
      const tableInsertPlan = tableInsertIndex === undefined ? undefined : plan[tableInsertIndex];
      const tableInsert =
        tableInsertPlan?.type === "tableStructure" &&
        tableInsertPlan.operation.type === "insertTable"
          ? tableInsertPlan.operation
          : undefined;
      let insertedTargetTableIndex: number | null = null;
      if (tableInsert?.type === "insertTable") {
        const payload = resolvedDocxTableStructureOperandPayload(tableInsert, comparison);
        if (payload.type !== "insertTable") {
          return panic("A terminal table insertion lost its canonical payload");
        }
        insertedTargetTableIndex = payload.change.tableIndex;
      }
      if (
        alignment.base.type !== "body" ||
        !isEmptyParagraphNode(carrier, baseSnapshot) ||
        tableInsertIndex === undefined ||
        tableInsertPlan?.type !== "tableStructure" ||
        terminalTargetTableIndex === undefined ||
        tableInsert?.type !== "insertTable" ||
        insertedTargetTableIndex !== terminalTargetTableIndex
      ) {
        // A table within a trailing paragraph run otherwise interrupts the
        // deletion chain. Keep the conservative plan and let verification
        // report that its ownership could not be reproduced.
        continue;
      }
      // The body schema permits the target table to precede `sectPr`
      // directly. Put the base's otherwise-final carrier immediately before
      // that table: its deleted mark can then remove the empty paragraph,
      // while rejecting the table insertion restores the original ending.
      const carrierSource = resolvedDocxSourceOperandForBlockId(baseResolvedSnapshot, carrier.id);
      plan[tableInsertIndex] = {
        type: "tableStructure",
        operation: resolvedDocxTerminalTableInsertionOperand(
          comparison,
          tableInsert,
          carrierSource,
        ),
      };
      dropped.add(carrierDeletionIndex);
      continue;
    }
    if (
      alignment.base.type === "body" &&
      run.blockIds.length === 1 &&
      run.blockIds.at(0) === carrier.id &&
      isEmptyParagraphNode(carrier, baseSnapshot) &&
      terminalTargetTableIndex !== undefined
    ) {
      const beforeCarrier = blocks.at((indexById.get(carrier.id) ?? 0) - 1);
      const baseTerminalTableIndex = beforeCarrier?.table?.outerTableIndex;
      const replacements = tableReplacementIndexes.flatMap((index) => {
        const candidate = plan[index];
        if (candidate?.type !== "tableStructure" || candidate.operation.type !== "replaceTable") {
          return [];
        }
        const instruction = candidate.operation;
        const payload = resolvedDocxTableStructureOperandPayload(instruction, comparison);
        if (payload.type !== "replaceTable") {
          return panic("A table replacement lost its canonical payload");
        }
        const indexes =
          payload.owner.type === "canonical-replacement"
            ? {
                base: payload.owner.replacement.baseTableIndex,
                target: payload.owner.replacement.revisedTableIndex,
              }
            : {
                base: payload.owner.deleted.tableIndex,
                target: payload.owner.inserted.tableIndex,
              };
        return indexes?.base === baseTerminalTableIndex &&
          indexes.target === terminalTargetTableIndex
          ? [{ index, instruction }]
          : [];
      });
      const replacement = replacements.length === 1 ? replacements.at(0) : undefined;
      if (replacement) {
        plan[replacement.index] = {
          type: "tableStructure",
          operation: resolvedDocxTerminalTableReplacementOperand(
            comparison,
            replacement.instruction,
            resolvedDocxSourceOperandForBlockId(baseResolvedSnapshot, carrier.id),
          ),
        };
        dropped.add(carrierDeletionIndex);
        continue;
      }
      const terminalInsertions = tableInsertIndexes.flatMap((index) => {
        const candidate = plan[index];
        if (candidate?.type !== "tableStructure" || candidate.operation.type !== "insertTable") {
          return [];
        }
        const instruction = candidate.operation;
        const payload = resolvedDocxTableStructureOperandPayload(instruction, comparison);
        return payload.type === "insertTable" &&
          payload.change.tableIndex === terminalTargetTableIndex
          ? [{ index, payload }]
          : [];
      });
      const terminalDeletions = tableDeleteIndexes.flatMap((index) => {
        const candidate = plan[index];
        if (candidate?.type !== "tableStructure" || candidate.operation.type !== "deleteTable") {
          return [];
        }
        const instruction = candidate.operation;
        const payload = resolvedDocxTableStructureOperandPayload(instruction, comparison);
        return payload.type === "deleteTable" &&
          payload.change.tableIndex === baseTerminalTableIndex
          ? [{ index, payload }]
          : [];
      });
      const inserted = terminalInsertions.length === 1 ? terminalInsertions.at(0) : undefined;
      const deleted = terminalDeletions.length === 1 ? terminalDeletions.at(0) : undefined;
      if (inserted && deleted) {
        const carrierSource = resolvedDocxSourceOperandForBlockId(baseResolvedSnapshot, carrier.id);
        const composed = resolvedDocxTableStructureOperand(comparison, {
          type: "replaceTable",
          source: deleted.payload.source,
          owner: {
            type: "structural-pair",
            deleted: deleted.payload.change,
            inserted: inserted.payload.change,
          },
        });
        if (composed.type !== "replaceTable") {
          return panic("A composed table replacement lost its structural discriminator");
        }
        plan[inserted.index] = {
          type: "tableStructure",
          operation: resolvedDocxTerminalTableReplacementOperand(
            comparison,
            composed,
            carrierSource,
          ),
        };
        dropped.add(deleted.index);
        dropped.add(carrierDeletionIndex);
        continue;
      }
    }
    const lastInsertIndex = run.insertIndexes.at(-1);
    if (lastInsertIndex === undefined) {
      if (run.chainStart) {
        const chainStart = resolvedDocxBaseBlockForId(comparison, run.chainStart.id);
        if (!chainStart) {
          return panic("A terminal merge carrier lost its canonical base block", {
            blockId: run.chainStart.id,
          });
        }
        const deletedEvents = run.blockIds.toReversed().map((blockId) => {
          const deletionIndex =
            deletionIndexByBlockId.get(blockId) ??
            panic("A trailing deletion run lost its deletion instruction", { blockId });
          const deletion = plan[deletionIndex];
          if (deletion?.type !== "deleteParagraph") {
            return panic("A trailing deletion run lost its exact deleted event", {
              deletionIndex,
            });
          }
          return deletion.event;
        });
        const firstDeleted =
          deletedEvents.at(0) ?? panic("A trailing deletion run has no deleted event");
        appended.push({
          type: "deleteTrailingParagraphs",
          operation: resolvedDocxTrailingDeletionOperand(comparison, {
            events: [firstDeleted, ...deletedEvents.slice(1)],
          }),
        });
        for (const blockId of run.blockIds) {
          const deletionIndex =
            deletionIndexByBlockId.get(blockId) ??
            panic("A trailing deletion run lost its deletion instruction", { blockId });
          dropped.add(deletionIndex);
        }
      } else if (isEmptyParagraphNode(carrier, baseSnapshot)) {
        dropped.add(carrierDeletionIndex);
      }
    } else {
      const lastInsertPlan = plan[lastInsertIndex];
      if (lastInsertPlan?.type !== "insertParagraph") {
        // A relocation is one atomic move instruction, so it is not rewritten
        // as a terminal insertion without its source half.
        continue;
      }
      const carrierDeletionPlan = plan[carrierDeletionIndex];
      if (carrierDeletionPlan?.type !== "deleteParagraph") {
        continue;
      }
      dropped.add(lastInsertIndex);
      const baseContent = resolvedDocxBaseBlockForId(comparison, carrier.id);
      if (!baseContent) {
        return panic("A terminal carrier lost its canonical base block", {
          blockId: carrier.id,
        });
      }
      const targetContent = resolvedDocxInsertedEventOperandPayload(
        lastInsertPlan.event,
        comparison,
      ).event.block;
      const propertiesEqual = docxParagraphPropertiesEqual(
        docxParagraphPropertiesFromBlock(baseContent),
        docxParagraphPropertiesFromBlock(targetContent),
      );
      if (
        carrier.text === targetContent.text &&
        canonicalJson(baseContent.runs) === canonicalJson(targetContent.runs) &&
        propertiesEqual
      ) {
        // The carrier already holds the words the last inserted paragraph
        // brings with the same authored runs, so neither half is a revision.
        dropped.add(carrierDeletionIndex);
      } else {
        plan[carrierDeletionIndex] = {
          type: "replaceTerminalParagraph",
          operation: resolvedDocxTerminalReplacementOperand(comparison, {
            deleted: carrierDeletionPlan.event,
            inserted: lastInsertPlan.event,
          }),
        };
      }
    }
  }
  return {
    operations: [...plan.filter((_, index) => !dropped.has(index)), ...appended],
  };
};

export type CompareStoryPlan = {
  readonly unsupported: readonly CompareUnsupportedPart[];
  /** The sole transport semantics for this story, consumed once by apply. */
  readonly program: DocxComparisonProgram;
};

export type PlanStoryCompareOptions = {
  /** Canonical semantics already bound to the exact base and target capsules. */
  comparison: ResolvedDocxStoryComparison;
  /** Cap on generated operations, separate from semantic comparison budgets. */
  maxOperations: number;
};

type PlanStoryCompareError = CompareDocxLoweringError | CompareDocxOperationLimitError;

const operationLimit = (maxOperations: number): CompareDocxOperationLimitError =>
  new CompareDocxOperationLimitError({
    message: "The comparison needs more operations than the engine generates.",
    limit: maxOperations,
  });

const loweringError = ({
  story,
  reason,
  message,
  baseBlockId,
  targetBlockId,
  tableIndex,
}: {
  story: FolioDocumentStoryHandle;
  reason: ConstructorParameters<typeof CompareDocxLoweringError>[0]["reason"];
  message: string;
  baseBlockId?: string;
  targetBlockId?: string;
  tableIndex?: number;
}): CompareDocxLoweringError =>
  new CompareDocxLoweringError({
    message,
    reason,
    story,
    ...(baseBlockId !== undefined && { baseBlockId }),
    ...(targetBlockId !== undefined && { targetBlockId }),
    ...(tableIndex !== undefined && { tableIndex }),
  });

const textChanged = (relation: FolioContentPairRelation): boolean =>
  relation.segments.some(({ type }) => type !== "equal");

type UnsupportedRelationField = {
  readonly reason:
    | "block-semantics"
    | "container-change"
    | "effective-inline-formatting"
    | "effective-paragraph-formatting"
    | "structural-boundary-change";
  readonly field: string;
};

const unsupportedRelationFields = (
  comparison: ResolvedDocxStoryComparison,
  relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
): readonly UnsupportedRelationField[] => {
  const unsupported: UnsupportedRelationField[] = [];
  for (const change of relation.blockChanges) {
    switch (change.field) {
      case "kind":
      case "blockProperties":
        unsupported.push({ reason: "block-semantics", field: change.field });
        break;
      case "table":
        if (!resolvedDocxTablePlacementChangeIsOwned(comparison, relation)) {
          unsupported.push({ reason: "block-semantics", field: change.field });
        }
        break;
      case "containerPath":
        unsupported.push({ reason: "container-change", field: change.field });
        break;
      case "structuralBoundaries":
        unsupported.push({ reason: "structural-boundary-change", field: change.field });
        break;
      default: {
        const unreachable: never = change;
        return panic("Unhandled canonical block change", { change: unreachable });
      }
    }
  }
  const formatting = relation.formatting;
  if (!formatting) return Object.freeze(unsupported);
  const authoredParagraphKeys = new Set(formatting.paragraph.authored.map(({ key }) => key));
  for (const change of formatting.paragraph.authored) {
    if (!docxParagraphPropertyChangeIsLowerable(change)) {
      unsupported.push({
        reason: "block-semantics",
        field: `paragraph.${change.key}`,
      });
    }
  }
  for (const change of formatting.paragraph.effective) {
    if (!authoredParagraphKeys.has(change.key)) {
      unsupported.push({
        reason: "effective-paragraph-formatting",
        field: `paragraph.${change.key}`,
      });
    }
  }
  for (const range of formatting.ranges) {
    const authoredKeys = new Set(range.formatting.authored.map(({ key }) => key));
    for (const change of range.formatting.effective) {
      if (!authoredKeys.has(change.key)) {
        unsupported.push({
          reason: "effective-inline-formatting",
          field: `runs.${change.key}`,
        });
      }
    }
  }
  return Object.freeze(unsupported);
};

/** Exhaustively lower one completed semantic story comparison. */
export const planStoryCompare = ({
  comparison,
  maxOperations,
}: PlanStoryCompareOptions): Result<CompareStoryPlan, PlanStoryCompareError> => {
  const {
    baseStory: story,
    baseSnapshot,
    targetSnapshot,
    comparison: contentComparison,
  } = resolvedDocxStoryComparisonPayload(comparison);
  const baseOperationSnapshot = resolvedDocxOperationSnapshot(baseSnapshot);
  const targetOperationSnapshot = resolvedDocxOperationSnapshot(targetSnapshot);
  const events = contentComparison.events;
  const unsupported: CompareUnsupportedPart[] = [];
  const operations: DocxComparisonOperationInput[] = [];

  const recordUnavailableRuns = (
    eventType: CompareUnsupportedEventType,
    blocks: readonly {
      readonly block: FolioContentBlock;
      readonly snapshot: ResolvedDocxStorySnapshot;
      readonly side: "base" | "target";
    }[],
  ): boolean => {
    let unavailable = false;
    for (const { block, snapshot, side } of blocks) {
      if (resolvedDocxHasExactAuthoredRuns(snapshot, block)) continue;
      unavailable = true;
      unsupported.push({
        reason: "block-semantics",
        story,
        eventType,
        field: "runs.authoredProjection",
        ...(side === "base"
          ? { baseBlockId: block.identity.id }
          : { targetBlockId: block.identity.id }),
      });
    }
    return unavailable;
  };

  const recordUnsupportedRelation = (
    eventType: CompareUnsupportedEventType,
    relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
  ): boolean => {
    const fields = unsupportedRelationFields(comparison, relation);
    for (const field of fields) {
      unsupported.push({
        reason: field.reason,
        story,
        eventType,
        field: field.field,
        baseBlockId: relation.base.block.identity.id,
        targetBlockId: relation.revised.block.identity.id,
      });
    }
    return fields.length > 0;
  };
  const insertOperation = (
    event: Extract<FolioContentComparisonEvent, { readonly type: "inserted" }>,
  ):
    | Extract<DocxComparisonOperationInput, { readonly type: "insertParagraph" }>
    | CompareDocxLoweringError => {
    if (event.boundary.type === "unanchoredContainer") {
      return loweringError({
        story,
        reason: "missing-insertion-anchor",
        message: "The target adds content to a container with no surviving paragraph boundary.",
        targetBlockId: event.block.identity.id,
      });
    }
    return {
      type: "insertParagraph",
      event: resolvedDocxInsertedEventOperand(comparison, event),
    };
  };

  for (const event of events) {
    switch (event.type) {
      case "unchanged":
        break;
      case "modified":
      case "formatting": {
        const { relation } = event;
        recordUnsupportedRelation(event.type, relation);
        const changesText = textChanged(relation);
        const hasAuthoredRunChanges =
          relation.formatting?.ranges.some(({ formatting }) => formatting.authored.length > 0) ??
          false;
        const unavailableRuns =
          (changesText || hasAuthoredRunChanges) &&
          recordUnavailableRuns(event.type, [
            { block: relation.base.block, snapshot: baseSnapshot, side: "base" },
            { block: relation.revised.block, snapshot: targetSnapshot, side: "target" },
          ]);
        const hasLowerableParagraphChanges =
          relation.formatting?.paragraph.authored.some(docxParagraphPropertyChangeIsLowerable) ??
          false;
        if (
          !unavailableRuns &&
          (changesText || hasAuthoredRunChanges || hasLowerableParagraphChanges)
        ) {
          operations.push({
            type: "pairedBlock",
            event: resolvedDocxPairedEventOperand(comparison, event),
          });
        }
        break;
      }
      case "deleted":
        operations.push({
          type: "deleteParagraph",
          event: resolvedDocxDeletedEventOperand(comparison, event),
        });
        break;
      case "inserted": {
        if (event.block.structuralBoundaries.length > 0) {
          unsupported.push({
            reason: "structural-boundary-change",
            story,
            eventType: "inserted",
            field: "structuralBoundaries",
            targetBlockId: event.block.identity.id,
          });
        }
        if (event.block.containerPath.length > 0) {
          unsupported.push({
            reason: "container-change",
            story,
            eventType: "inserted",
            field: "containerPath",
            targetBlockId: event.block.identity.id,
          });
        }
        const unavailableRuns = recordUnavailableRuns("inserted", [
          { block: event.block, snapshot: targetSnapshot, side: "target" },
        ]);
        if (!unavailableRuns) {
          const lowered = insertOperation(event);
          if (lowered instanceof CompareDocxLoweringError) return Result.err(lowered);
          operations.push(lowered);
        }
        break;
      }
      case "movedFrom":
        break;
      case "movedTo": {
        const relation = event.move.relation;
        recordUnsupportedRelation("moved", relation);
        const unavailableRuns = recordUnavailableRuns("moved", [
          { block: relation.revised.block, snapshot: targetSnapshot, side: "target" },
        ]);
        if (unavailableRuns) break;
        if (event.move.destinationBoundary.type === "unanchoredContainer") {
          return Result.err(
            loweringError({
              story,
              reason: "missing-insertion-anchor",
              message:
                "The moved paragraph has no surviving destination boundary in its container.",
              baseBlockId: relation.base.block.identity.id,
              targetBlockId: relation.revised.block.identity.id,
            }),
          );
        }
        const sourceRemovalBoundary = event.move.sourceRemovalBoundary;
        switch (sourceRemovalBoundary.type) {
          case "successorParagraph":
          case "terminalPredecessor":
            operations.push({
              type: "moveParagraph",
              event: resolvedDocxMoveEventOperand(comparison, event),
            });
            break;
          case "unanchoredContainer":
            return Result.err(
              loweringError({
                story,
                reason: "missing-removal-boundary",
                message:
                  "The source paragraph has no same-container boundary that can carry its removal.",
                baseBlockId: relation.base.block.identity.id,
                targetBlockId: relation.revised.block.identity.id,
              }),
            );
          default: {
            const unreachable: never = sourceRemovalBoundary;
            return panic("Unhandled neutral paragraph removal boundary", {
              boundary: unreachable,
            });
          }
        }
        break;
      }
      case "split": {
        const [first, second] = event.relations;
        recordUnsupportedRelation("split", first);
        recordUnsupportedRelation("split", second);
        const baseBlock = first.base.block;
        const firstBlock = first.revised.block;
        const secondBlock = second.revised.block;
        const unavailableRuns = recordUnavailableRuns("split", [
          { block: baseBlock, snapshot: baseSnapshot, side: "base" },
          { block: firstBlock, snapshot: targetSnapshot, side: "target" },
          { block: secondBlock, snapshot: targetSnapshot, side: "target" },
        ]);
        if (unavailableRuns) break;
        operations.push({
          type: "splitParagraph",
          event: resolvedDocxSplitEventOperand(comparison, event),
        });
        break;
      }
      case "merge": {
        const [first, second] = event.relations;
        recordUnsupportedRelation("merge", first);
        recordUnsupportedRelation("merge", second);
        const firstBlock = first.base.block;
        const secondBlock = second.base.block;
        const targetBlock = first.revised.block;
        const unavailableRuns = recordUnavailableRuns("merge", [
          { block: firstBlock, snapshot: baseSnapshot, side: "base" },
          { block: secondBlock, snapshot: baseSnapshot, side: "base" },
          { block: targetBlock, snapshot: targetSnapshot, side: "target" },
        ]);
        if (unavailableRuns) break;
        operations.push({
          type: "mergeParagraphs",
          event: resolvedDocxMergeEventOperand(comparison, event),
        });
        break;
      }
      case "tableReplacement": {
        const { replacement } = event;
        const baseBlock = replacement.baseBlocks[0];
        const unavailableRuns = recordUnavailableRuns("tableReplacement", [
          ...replacement.baseBlocks.map((block) => ({
            block,
            snapshot: baseSnapshot,
            side: "base" as const,
          })),
          ...replacement.revisedBlocks.map((block) => ({
            block,
            snapshot: targetSnapshot,
            side: "target" as const,
          })),
        ]);
        if (unavailableRuns) break;
        operations.push({
          type: "tableStructure",
          operation: resolvedDocxTableStructureOperand(comparison, {
            type: "replaceTable",
            source: resolvedDocxSourceOperand(baseSnapshot, baseBlock),
            owner: { type: "canonical-replacement", replacement },
          }),
        });
        break;
      }
      case "structural": {
        if (event.memberIndex !== 0) break;
        const structural = event.change;
        const firstBlock = structural.blocks[0];
        switch (structural.type) {
          case "table-row-delete":
            operations.push({
              type: "tableStructure",
              operation: resolvedDocxTableStructureOperand(comparison, {
                type: "deleteTableRow",
                source: resolvedDocxSourceOperand(baseSnapshot, firstBlock),
                change: structural,
              }),
            });
            break;
          case "table-column-delete":
            operations.push({
              type: "tableStructure",
              operation: resolvedDocxTableStructureOperand(comparison, {
                type: "deleteTableColumn",
                source: resolvedDocxSourceOperand(baseSnapshot, firstBlock),
                change: structural,
              }),
            });
            break;
          case "table-delete":
            operations.push({
              type: "tableStructure",
              operation: resolvedDocxTableStructureOperand(comparison, {
                type: "deleteTable",
                source: resolvedDocxSourceOperand(baseSnapshot, firstBlock),
                change: structural,
              }),
            });
            break;
          case "table-insert": {
            if (
              recordUnavailableRuns(
                "structural",
                structural.blocks.map((block) => ({
                  block,
                  snapshot: targetSnapshot,
                  side: "target" as const,
                })),
              )
            ) {
              break;
            }
            const anchor = resolvedDocxStructuralInsertionBoundary(comparison, structural);
            if (!anchor) {
              return Result.err(
                loweringError({
                  story,
                  reason: "missing-insertion-anchor",
                  message: "The target adds a table to a story with no tracked insertion anchor.",
                  tableIndex: structural.tableIndex,
                }),
              );
            }
            operations.push({
              type: "tableStructure",
              operation: resolvedDocxTableStructureOperand(comparison, {
                type: "insertTable",
                anchor,
                change: structural,
              }),
            });
            break;
          }
          case "table-row-insert": {
            const anchor = resolvedDocxStructuralInsertionBoundary(comparison, structural);
            if (!anchor) {
              return Result.err(
                loweringError({
                  story,
                  reason: "table-row-anchor",
                  message: "The inserted table row has no surviving table anchor.",
                  tableIndex: structural.tableIndex,
                }),
              );
            }
            if (
              recordUnavailableRuns(
                "structural",
                structural.blocks.map((block) => ({
                  block,
                  snapshot: targetSnapshot,
                  side: "target" as const,
                })),
              )
            ) {
              break;
            }
            operations.push({
              type: "tableStructure",
              operation: resolvedDocxTableStructureOperand(comparison, {
                type: "insertTableRow",
                anchor,
                change: structural,
              }),
            });
            break;
          }
          case "table-column-insert": {
            unsupported.push({
              reason: "table-column-content",
              story,
              eventType: "structural",
              tableIndex: structural.tableIndex,
            });
            if (
              recordUnavailableRuns(
                "structural",
                structural.blocks.map((block) => ({
                  block,
                  snapshot: targetSnapshot,
                  side: "target" as const,
                })),
              )
            ) {
              break;
            }
            const anchor = resolvedDocxStructuralInsertionBoundary(comparison, structural);
            if (!anchor) {
              return panic("A canonical table-column insertion lost its structural anchor", {
                tableIndex: structural.tableIndex,
              });
            }
            operations.push({
              type: "tableStructure",
              operation: resolvedDocxTableStructureOperand(comparison, {
                type: "insertTableColumn",
                anchor,
                change: structural,
              }),
            });
            break;
          }
          default: {
            const unreachable: never = structural;
            return panic("Unhandled canonical structural change", { change: unreachable });
          }
        }
        break;
      }
      default: {
        const unreachable: never = event;
        return panic("Unhandled canonical comparison event", { event: unreachable });
      }
    }
    if (operations.length > maxOperations) {
      return Result.err(operationLimit(maxOperations));
    }
  }

  const tableFormat = resolvedDocxTableFormatOperand(comparison);
  if (tableFormat) operations.push({ type: "tableFormat", operation: tableFormat });

  const deletedBlockIds = new Set(
    operations.flatMap((operation) =>
      operation.type === "deleteParagraph"
        ? [
            resolvedDocxDeletedEventOperandPayload(operation.event, comparison).event.block.identity
              .id,
          ]
        : [],
    ),
  );
  const unavailableTerminalCarrierIds = new Set<string>();
  for (const { block: carrier } of resolvedDocxBaseTerminalBlocks(comparison)) {
    if (!deletedBlockIds.has(carrier.id)) continue;
    if (resolvedDocxHasExactAuthoredRuns(baseSnapshot, carrier)) continue;
    unavailableTerminalCarrierIds.add(carrier.id);
    unsupported.push({
      reason: "block-semantics",
      story,
      eventType: "deleted",
      field: "runs.authoredProjection",
      baseBlockId: carrier.id,
    });
  }

  const planned = withTrailingDeletionRules({
    comparison,
    baseResolvedSnapshot: baseSnapshot,
    baseSnapshot: baseOperationSnapshot,
    targetSnapshot: targetOperationSnapshot,
    unavailableTerminalCarrierIds,
    operations,
  });

  if (planned.operations.length > maxOperations) {
    return Result.err(operationLimit(maxOperations));
  }
  const program = DocxComparisonProgram.create(comparison, planned.operations);
  if (program.size > maxOperations) return Result.err(operationLimit(maxOperations));
  return Result.ok({
    unsupported: Object.freeze(unsupported),
    program,
  });
};
