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

import { docxParagraphPropertyChangeIsLowerable } from "../internal/compare/docx-paragraph-transport";
import {
  resolvedDocxDeletedEventOperand,
  resolvedDocxInsertedEventOperand,
  resolvedDocxMergeEventOperand,
  resolvedDocxMoveEventOperand,
  resolvedDocxPairedEventOperand,
  resolvedDocxSplitEventOperand,
  resolvedDocxStoryComparisonPayload,
  resolvedDocxStructuralInsertionBoundary,
  resolvedDocxTablePlacementChangeIsOwned,
  resolvedDocxTerminalTransitionForEvent,
  resolvedDocxTerminalTransitionOperandPayload,
  resolvedDocxTableFormatOperands,
  resolvedDocxTableStructureOperand,
  resolvedDocxTableStructureOperandPayload,
  type ResolvedDocxStoryComparison,
} from "../internal/compare/resolved-docx-story-comparison";
import {
  resolvedDocxHasExactAuthoredRuns,
  resolvedDocxUnsupportedProjectionFields,
  resolvedDocxSourceOperand,
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
import type { FolioContentBlock } from "./content-types";
import { CompareDocxOperationLimitError, type CompareUnsupportedPart } from "./types";

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

type PlanStoryCompareError = CompareDocxOperationLimitError;

type CompareUnsupportedEventType = Extract<
  CompareUnsupportedPart,
  { readonly eventType: unknown }
>["eventType"];

type ComparePlanUnsupportedPart = Extract<CompareUnsupportedPart, { readonly eventType: unknown }>;

type PlannedOperationDisposition =
  | {
      readonly type: "emitted";
      readonly operation: DocxComparisonOperationInput;
    }
  | {
      readonly type: "omitted";
      readonly unsupported: readonly [ComparePlanUnsupportedPart, ...ComparePlanUnsupportedPart[]];
    };

const emittedOperation = (
  operation: DocxComparisonOperationInput,
): PlannedOperationDisposition => ({ type: "emitted", operation });

const omittedOperation = (
  first: ComparePlanUnsupportedPart,
  ...rest: ComparePlanUnsupportedPart[]
): PlannedOperationDisposition => ({
  type: "omitted",
  unsupported: Object.freeze([first, ...rest]),
});

const operationLimit = (maxOperations: number): CompareDocxOperationLimitError =>
  new CompareDocxOperationLimitError({
    message: "The comparison needs more operations than the engine generates.",
    limit: maxOperations,
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
  const events = contentComparison.events;
  const unsupported: CompareUnsupportedPart[] = [];
  const operations: DocxComparisonOperationInput[] = [];

  const appendOperationDisposition = (disposition: PlannedOperationDisposition): void => {
    switch (disposition.type) {
      case "emitted":
        operations.push(disposition.operation);
        break;
      case "omitted":
        unsupported.push(...disposition.unsupported);
        break;
      default: {
        const unreachable: never = disposition;
        return panic("Unhandled DOCX lowering disposition", { disposition: unreachable });
      }
    }
  };

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
      const blockIdentity =
        side === "base" ? { baseBlockId: block.identity.id } : { targetBlockId: block.identity.id };
      if (!resolvedDocxHasExactAuthoredRuns(snapshot, block)) {
        unavailable = true;
        unsupported.push({
          reason: "block-semantics",
          story,
          eventType,
          field: "runs.authoredProjection",
          ...blockIdentity,
        });
      }
      for (const field of resolvedDocxUnsupportedProjectionFields(snapshot, block)) {
        unavailable = true;
        unsupported.push({
          reason: "block-semantics",
          story,
          eventType,
          field,
          ...blockIdentity,
        });
      }
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
  ): PlannedOperationDisposition => {
    if (event.boundary.type === "unanchoredContainer") {
      return omittedOperation({
        reason: "missing-insertion-anchor",
        story,
        eventType: "inserted",
        targetBlockId: event.block.identity.id,
      });
    }
    return emittedOperation({
      type: "insertParagraph",
      event: resolvedDocxInsertedEventOperand(comparison, event),
    });
  };

  const moveOperation = (
    event: Extract<FolioContentComparisonEvent, { readonly type: "movedTo" }>,
  ): PlannedOperationDisposition => {
    const { relation, destinationBoundary, sourceRemovalBoundary } = event.move;
    const boundaryUnsupported: ComparePlanUnsupportedPart[] = [];
    if (destinationBoundary.type === "unanchoredContainer") {
      boundaryUnsupported.push({
        reason: "missing-insertion-anchor",
        story,
        eventType: "moved",
        baseBlockId: relation.base.block.identity.id,
        targetBlockId: relation.revised.block.identity.id,
      });
    }
    switch (sourceRemovalBoundary.type) {
      case "successorParagraph":
      case "successorTable":
        break;
      case "terminalPredecessor":
      case "unanchoredContainer":
        boundaryUnsupported.push({
          reason: "missing-removal-boundary",
          story,
          eventType: "moved",
          baseBlockId: relation.base.block.identity.id,
          targetBlockId: relation.revised.block.identity.id,
        });
        break;
      default: {
        const unreachable: never = sourceRemovalBoundary;
        return panic("Unhandled neutral paragraph removal boundary", {
          boundary: unreachable,
        });
      }
    }
    const [first, ...rest] = boundaryUnsupported;
    return first
      ? omittedOperation(first, ...rest)
      : emittedOperation({
          type: "moveParagraph",
          event: resolvedDocxMoveEventOperand(comparison, event),
        });
  };

  for (const event of events) {
    const terminal = resolvedDocxTerminalTransitionForEvent(comparison, event);
    if (terminal) {
      if (!terminal.isOwner) continue;
      const transition = resolvedDocxTerminalTransitionOperandPayload(
        terminal.operation,
        comparison,
      );
      let unavailable = false;
      const boundaryUnsupported: ComparePlanUnsupportedPart[] = [];
      if (transition.type === "paragraph") {
        if (transition.chainStartRewrite !== null) {
          const { event: rewrite } = transition.chainStartRewrite;
          const eventType = rewrite.type;
          unavailable = recordUnsupportedRelation(eventType, rewrite.relation) || unavailable;
          unavailable =
            recordUnavailableRuns(eventType, [
              {
                block: rewrite.relation.base.block,
                snapshot: baseSnapshot,
                side: "base",
              },
              {
                block: rewrite.relation.revised.block,
                snapshot: targetSnapshot,
                side: "target",
              },
            ]) || unavailable;
        }
        const recordedMoves = new Set<
          Extract<FolioContentComparisonEvent, { readonly type: "movedTo" }>["move"]
        >();
        for (const member of transition.sourceMembers) {
          if (member.type !== "movedFrom") continue;
          recordedMoves.add(member.occurrence.event.move);
        }
        for (const member of transition.targetMembers) {
          if (member.type === "inserted") {
            const inserted = member.occurrence.event;
            if (inserted.block.structuralBoundaries.length > 0) {
              unsupported.push({
                reason: "structural-boundary-change",
                story,
                eventType: "inserted",
                field: "structuralBoundaries",
                targetBlockId: inserted.block.identity.id,
              });
              unavailable = true;
            }
            if (inserted.block.containerPath.length > 0) {
              unsupported.push({
                reason: "container-change",
                story,
                eventType: "inserted",
                field: "containerPath",
                targetBlockId: inserted.block.identity.id,
              });
              unavailable = true;
            }
            unavailable =
              recordUnavailableRuns("inserted", [
                { block: inserted.block, snapshot: targetSnapshot, side: "target" },
              ]) || unavailable;
          } else {
            recordedMoves.add(member.occurrence.event.move);
          }
        }
        for (const move of recordedMoves) {
          unavailable = recordUnsupportedRelation("moved", move.relation) || unavailable;
          unavailable =
            recordUnavailableRuns("moved", [
              { block: move.relation.revised.block, snapshot: targetSnapshot, side: "target" },
            ]) || unavailable;
          if (
            move.destinationBoundary.type === "unanchoredContainer" &&
            transition.targetCarrier !== move.relation.revised.block
          ) {
            boundaryUnsupported.push({
              reason: "missing-insertion-anchor",
              story,
              eventType: "moved",
              baseBlockId: move.relation.base.block.identity.id,
              targetBlockId: move.relation.revised.block.identity.id,
            });
          }
          if (
            move.sourceRemovalBoundary.type === "unanchoredContainer" ||
            (move.sourceRemovalBoundary.type === "terminalPredecessor" &&
              !transition.sourceMembers.some(
                (member) => member.type === "movedFrom" && member.occurrence.event.move === move,
              ))
          ) {
            boundaryUnsupported.push({
              reason: "missing-removal-boundary",
              story,
              eventType: "moved",
              baseBlockId: move.relation.base.block.identity.id,
              targetBlockId: move.relation.revised.block.identity.id,
            });
          }
        }
      } else if (transition.type === "tableAppend") {
        for (const suffixMember of transition.targetSuffix) {
          if (suffixMember.type === "terminalCarrier") {
            const carrierBlock = suffixMember.occurrence.event.block;
            if (carrierBlock.structuralBoundaries.length > 0) {
              unsupported.push({
                reason: "structural-boundary-change",
                story,
                eventType: "inserted",
                field: "structuralBoundaries",
                targetBlockId: carrierBlock.identity.id,
              });
              unavailable = true;
            }
            if (carrierBlock.containerPath.length > 0) {
              unsupported.push({
                reason: "container-change",
                story,
                eventType: "inserted",
                field: "containerPath",
                targetBlockId: carrierBlock.identity.id,
              });
              unavailable = true;
            }
            unavailable =
              recordUnavailableRuns("inserted", [
                { block: carrierBlock, snapshot: targetSnapshot, side: "target" },
              ]) || unavailable;
            continue;
          }
          if (suffixMember.type === "table") {
            const table = resolvedDocxTableStructureOperandPayload(
              suffixMember.operation,
              comparison,
            );
            if (table.type !== "insertTable") {
              return panic("A terminal target suffix lost its table insertion role");
            }
            unavailable =
              recordUnavailableRuns(
                "structural",
                table.change.blocks.map((block) => ({
                  block,
                  snapshot: targetSnapshot,
                  side: "target" as const,
                })),
              ) || unavailable;
            continue;
          }
          const { member } = suffixMember;
          if (member.type === "inserted") {
            const inserted = member.occurrence.event;
            if (inserted.block.structuralBoundaries.length > 0) {
              unsupported.push({
                reason: "structural-boundary-change",
                story,
                eventType: "inserted",
                field: "structuralBoundaries",
                targetBlockId: inserted.block.identity.id,
              });
              unavailable = true;
            }
            if (inserted.block.containerPath.length > 0) {
              unsupported.push({
                reason: "container-change",
                story,
                eventType: "inserted",
                field: "containerPath",
                targetBlockId: inserted.block.identity.id,
              });
              unavailable = true;
            }
            unavailable =
              recordUnavailableRuns("inserted", [
                { block: inserted.block, snapshot: targetSnapshot, side: "target" },
              ]) || unavailable;
            continue;
          }
          const move = member.occurrence.event.move;
          unavailable = recordUnsupportedRelation("moved", move.relation) || unavailable;
          unavailable =
            recordUnavailableRuns("moved", [
              { block: move.relation.revised.block, snapshot: targetSnapshot, side: "target" },
            ]) || unavailable;
          if (
            move.sourceRemovalBoundary.type === "terminalPredecessor" ||
            move.sourceRemovalBoundary.type === "unanchoredContainer"
          ) {
            boundaryUnsupported.push({
              reason: "missing-removal-boundary",
              story,
              eventType: "moved",
              baseBlockId: move.relation.base.block.identity.id,
              targetBlockId: move.relation.revised.block.identity.id,
            });
          }
        }
      } else {
        const table = resolvedDocxTableStructureOperandPayload(transition.operation, comparison);
        let targetBlocks: readonly FolioContentBlock[];
        switch (table.type) {
          case "insertTable":
            targetBlocks = table.change.blocks;
            break;
          case "replaceTable":
            targetBlocks =
              table.owner.type === "canonical-replacement"
                ? table.owner.replacement.revisedBlocks
                : table.owner.inserted.blocks;
            break;
          default:
            return panic("A terminal table transition lost its insert or replacement role", {
              operationType: table.type,
            });
        }
        unavailable = recordUnavailableRuns(
          table.type === "replaceTable" ? "tableReplacement" : "structural",
          targetBlocks.map((block) => ({
            block,
            snapshot: targetSnapshot,
            side: "target" as const,
          })),
        );
        for (const member of transition.sourceMembers.slice(0, -1)) {
          if (member.type !== "movedFrom") continue;
          const move = member.occurrence.event.move;
          unavailable = recordUnsupportedRelation("moved", move.relation) || unavailable;
          unavailable =
            recordUnavailableRuns("moved", [
              { block: move.relation.revised.block, snapshot: targetSnapshot, side: "target" },
            ]) || unavailable;
          if (move.destinationBoundary.type === "unanchoredContainer") {
            boundaryUnsupported.push({
              reason: "missing-insertion-anchor",
              story,
              eventType: "moved",
              baseBlockId: move.relation.base.block.identity.id,
              targetBlockId: move.relation.revised.block.identity.id,
            });
          }
        }
      }
      const [firstBoundaryUnsupported, ...remainingBoundaryUnsupported] = boundaryUnsupported;
      if (firstBoundaryUnsupported) {
        appendOperationDisposition(
          omittedOperation(firstBoundaryUnsupported, ...remainingBoundaryUnsupported),
        );
      } else if (!unavailable) {
        appendOperationDisposition(
          emittedOperation({ type: "terminalTransition", operation: terminal.operation }),
        );
      }
      if (operations.length > maxOperations) return Result.err(operationLimit(maxOperations));
      continue;
    }
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
        const disposition = insertOperation(event);
        if (disposition.type === "omitted" || !unavailableRuns) {
          appendOperationDisposition(disposition);
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
        const disposition = moveOperation(event);
        if (disposition.type === "omitted" || !unavailableRuns) {
          appendOperationDisposition(disposition);
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
            const unavailableRuns = recordUnavailableRuns(
              "structural",
              structural.blocks.map((block) => ({
                block,
                snapshot: targetSnapshot,
                side: "target" as const,
              })),
            );
            const anchor = resolvedDocxStructuralInsertionBoundary(comparison, structural);
            if (!anchor) {
              appendOperationDisposition(
                omittedOperation({
                  reason: "missing-insertion-anchor",
                  story,
                  eventType: "structural",
                  tableIndex: structural.tableIndex,
                }),
              );
              break;
            }
            if (!unavailableRuns) {
              appendOperationDisposition(
                emittedOperation({
                  type: "tableStructure",
                  operation: resolvedDocxTableStructureOperand(comparison, {
                    type: "insertTable",
                    anchor,
                    change: structural,
                  }),
                }),
              );
            }
            break;
          }
          case "table-row-insert": {
            const anchor = resolvedDocxStructuralInsertionBoundary(comparison, structural);
            const unavailableRuns = recordUnavailableRuns(
              "structural",
              structural.blocks.map((block) => ({
                block,
                snapshot: targetSnapshot,
                side: "target" as const,
              })),
            );
            if (!anchor) {
              appendOperationDisposition(
                omittedOperation({
                  reason: "table-row-anchor",
                  story,
                  eventType: "structural",
                  tableIndex: structural.tableIndex,
                }),
              );
              break;
            }
            if (!unavailableRuns) {
              appendOperationDisposition(
                emittedOperation({
                  type: "tableStructure",
                  operation: resolvedDocxTableStructureOperand(comparison, {
                    type: "insertTableRow",
                    anchor,
                    change: structural,
                  }),
                }),
              );
            }
            break;
          }
          case "table-column-insert": {
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

  for (const tableFormat of resolvedDocxTableFormatOperands(comparison)) {
    operations.push({ type: "tableFormat", operation: tableFormat });
  }
  if (operations.length > maxOperations) return Result.err(operationLimit(maxOperations));
  const program = DocxComparisonProgram.create(comparison, operations);
  if (program.size > maxOperations) return Result.err(operationLimit(maxOperations));
  return Result.ok({
    unsupported: Object.freeze(unsupported),
    program,
  });
};
