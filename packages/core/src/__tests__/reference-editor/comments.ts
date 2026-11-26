/** ProseMirror review and suggestion resolution for the reference editor. */
import { panic } from "better-result";
import { closeHistory } from "prosemirror-history";
import {
  Fragment,
  type Mark,
  type MarkType,
  type Node as PMNode,
  type ResolvedPos,
} from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";
import { Mapping } from "prosemirror-transform";

import {
  type LockedControlViolation,
  type ParagraphMarkChange,
  REVIEW_CARRIERS,
} from "@stll/docx-core/model";
import { paragraphMarkOutcome } from "@stll/docx-core/ops";

import { markParagraphPropertySourceTransfers } from "../../docx/paragraphPropertySource";
import type { RemovedSectionReference } from "../../internal/sectionEndpointResolution";
import {
  deleteTextBoxesAnchoredAt,
  droppedTextBoxAnchorIds,
  moveTextBoxesPastNextParagraph,
} from "../../prosemirror/anchoredTextBoxes";
import {
  expectFieldAttrs,
  expectParagraphAttrs,
  expectRunPropertyChangeMarkAttrs,
  expectTrackedChangeMarkAttrs,
} from "../../prosemirror/attrs";
import {
  collectSuggestions,
  findAIEditRevisionRange,
  getSuggestions,
  isPPrMarkAttr,
  isSuggestionMark,
  isTableCellRevisionAttr,
  isTableRowRevisionAttr,
  mergeSegments,
  paragraphMarkWasAdded,
  readStructuralSuggestion,
  readSuggestedMarker,
  readSuggestedParagraphMark,
  readSuggestedParagraphPropertyChange,
  suggestedCommentOf,
  suggestionIdOfMarkRecord,
  type TableRowRevisionAttr,
  trackedChangeEntryRevisionIds,
} from "../../prosemirror/commands/comments";
import {
  hasSerializableParagraphPropertyChange,
  paragraphPropertiesSnapshot,
} from "../../prosemirror/commands/propertyChangeScope";
import { resolveParagraphChangeAttrs } from "../../prosemirror/commands/resolveParagraphProperties";
import {
  removeCommentAnchors,
  retireThreadsResolvedAway,
} from "../../prosemirror/commentRecords";
import {
  addedBreakCarrierBefore,
  finalParagraphsOf,
  paragraphEndsItsContainer,
} from "../../prosemirror/containerFinalParagraph";
import {
  resolutionRemovesControl,
  withoutResolvedEnclosures,
} from "../../prosemirror/contentControlRevisions";
import { copyOriginBoundaryAttrsAfterResolution } from "../../prosemirror/copyOriginResolution";
import { resolveEmptyFieldResultRuns } from "../../prosemirror/emptyFieldResultRuns";
import { JOINED_RUNS_RESTYLED_META } from "../../prosemirror/extensions/features/JoinedRunStyleExtension";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "../../prosemirror/extensions/nodes/SdtExtension";
import { canonicalFieldNode } from "../../prosemirror/fieldRepresentation";
import {
  isBlockBoundaryNode,
  normalizeKeptRangeBoundaries,
  paragraphRemainsOf,
  type RemovedSubtree,
  removeResolvedSubtree,
  revisionLayerRemovesContent,
  type RevisionSelection,
} from "../../prosemirror/keptRangeBoundaries";
import {
  lockedControlSpans,
  lockedControlViolationOfTransform,
} from "../../prosemirror/lockedControls";
import {
  markChangedParagraphRanges,
  markStructuralChange,
  markTrackedSectionEndpointRemoval,
} from "../../prosemirror/paragraphChangeTrackerMeta";
import {
  carryParagraphProperties,
  paragraphLeftAfter,
} from "../../prosemirror/paragraphPropertyCarry";
import { getDocumentNumbering } from "../../prosemirror/plugins/documentNumbering";
import { getDocumentStyleResolver } from "../../prosemirror/plugins/documentStyleState";
import { resolvedPreservedXmlAttrs } from "../../prosemirror/preservedXmlReview";
import { nodePropertyRevisionSites } from "../../prosemirror/revisionCarriers";
import { RUN_FORMATTING_MARK_NAMES } from "../../prosemirror/runFormattingMarkNames";
import { paragraphRunStyleContextAt } from "../../prosemirror/runStyleFormatting";
import type { TrackedRevisionAncestor } from "../../prosemirror/schema/marks";
import type { ParagraphPropertyChangeAttrs } from "../../prosemirror/schema/nodes";
import { sectionRecordCopyPositions } from "../../prosemirror/sectionCarrier";
import { isTableCellRetainedInReviewView } from "../../prosemirror/tableCellRevisionVisibility";
import {
  markRowEmptiedInBatch,
  reconcileOmittedGridSlots,
  reconcileTableGridAfterColumnRemoval,
  rowHoldsNoCells,
  type RowsEmptiedInBatch,
} from "../../prosemirror/tableGridMutation";
import { isTableRowRevisionCompanion } from "../../prosemirror/tableRowRevisionOwnership";
import {
  runInCellPending,
  runParagraphIntoTable,
} from "../../prosemirror/tableRunIn";
import {
  retainedTrackedRevisionMark,
  trackedRevisionPathOf,
} from "../../prosemirror/trackedRevisionPath";
import { extractTrackedChanges } from "../../prosemirror/utils/extractTrackedChanges";
import { paragraphContentOf } from "../../prosemirror/zeroWidthAnchors";
import type { RunPropertyChange } from "../../types/document";
import { resolveStateStory } from "./markupViewProjection";
import {
  joinAtParagraphMark,
  requestDeterministicParaIds,
  setParagraphAttrsWithRebasedRunFormatting,
} from "./paragraphRevision";
import { rejoinRunsAt } from "./rejoinRunCarriers";
import { resolvedCommentPoints } from "./resolvedCommentPoints";
import { resolveNodePropertyChangeAttrs } from "./resolveNodePropertyChangeAttrs";
import { RevisionResolutionStep } from "./revisionResolutionStep";
import { finalRevisionParagraphRanges } from "./revisionResolutionTracking";
import { rejectedRunFormatting } from "./runPropertyChangeResolution";
import { hasMatchingCollapsedTableCellMerge } from "./tableCellMergeProjection";
import {
  resolveCollapsedTableCellMerge,
  resolveVisibleTableCellMerge,
} from "./tableCellMergeResolution";
import {
  deleteTableRowAt,
  removeRowsWithoutCellsAfterBatch,
} from "./tableRevisionCleanup";

export {
  addCommentMark,
  removeCommentMark,
  type ParagraphBoundaryChange,
  findParagraphBoundaryChangeAtPosition,
  trackedChangeEntryRevisionIds,
  findAIEditRevisionRange,
  suggestionCommentIds,
  type SuggestionKind,
  type SuggestionAppliedAs,
  type FolioSuggestion,
  getSuggestions,
  pendingSuggestionRevisionIds,
  suggestionOwnersByRevision,
  suggestionIdOfRevision,
  findSuggestionRange,
  findChangeAtPosition,
  findNextChange,
  findPreviousChange,
} from "../../prosemirror/commands/comments";

type ResolveMode = "accept" | "reject";

/**
 * What every review resolution does before it is dispatched: it retires the
 * comment threads it left without an anchor, and it is an undo step of its
 * own, never merged into the edit it resolves.
 */
const finishResolution = (state: EditorState, tr: Transaction): void => {
  retireThreadsResolvedAway(state, tr);
  closeHistory(tr);
};

/**
 * Resolve a tracked change: accept or reject.
 * - Accept: keep insertions (remove mark), delete deletions (remove text)
 * - Reject: keep deletions (remove mark), delete insertions (remove text)
 * - Run formatting: accept the current formatting, or restore the previous formatting
 *
 * Pass `revisionId` to scope the operation to one specific
 * revision — otherwise overlapping marks for other revisions get
 * processed too, silently consuming pending work. Without an id
 * the operation matches every revision mark or property-change entry in the
 * range (the bulk accept-all/reject-all path).
 */
/**
 * A range a resolution takes out. Only {@link removeResolvedSubtree} answers
 * for what it holds, so no range can drop the boundaries inside it.
 */
type RemovedRange = {
  from: number;
  to: number;
  removals: readonly RemovedSubtree[];
};

const keptIn = (range: RemovedRange): PMNode[] =>
  range.removals.flatMap(({ kept }) => kept);

function resolveChangeUnguarded(
  from: number,
  to: number,
  mode: "accept" | "reject",
  revisionIds?: readonly number[],
): Command {
  return (state, dispatch) => {
    const insertionType = state.schema.marks["insertion"];
    const deletionType = state.schema.marks["deletion"];
    const styleResolver = getDocumentStyleResolver(state);
    const numbering = getDocumentNumbering(state);

    const keepType = mode === "accept" ? insertionType : deletionType;
    const removeType = mode === "accept" ? deletionType : insertionType;
    const revisionSet =
      revisionIds === undefined ? null : new Set<number>(revisionIds);
    // A range-wide removal lets ProseMirror coalesce one revision split by
    // inline formatting. The id-scoped path must keep matching each mark.
    const removeKeptMarksInBulk =
      revisionSet === null && keepType !== undefined;
    const matchesRevision = (mark: { attrs: Record<string, unknown> }) =>
      revisionSet === null ||
      (typeof mark.attrs["revisionId"] === "number" &&
        revisionSet.has(mark.attrs["revisionId"]));

    if (dispatch) {
      const points =
        mode === "accept"
          ? resolvedCommentPoints({
              doc: state.doc,
              retiresStructure: (node, position) => {
                if (node.type.name === "paragraph")
                  return (
                    collectPPrMarkOp(
                      node,
                      position,
                      from,
                      to,
                      mode,
                      revisionSet,
                    )?.action === "join"
                  );
                if (!rangeCoversNode(from, to, position, node)) return false;
                if (node.type.name === "tableRow")
                  return (
                    collectTableRowStructuralOp(
                      node,
                      position,
                      mode,
                      revisionSet,
                    )?.action === "remove"
                  );
                return collectTableCellStructuralOps(
                  node,
                  position,
                  mode,
                  revisionSet,
                ).some(
                  (op) => op.type === "membership" && op.action === "remove",
                );
              },
              removes: (node, position) =>
                position >= from &&
                position + node.nodeSize <= to &&
                node.marks.some(
                  (mark) => mark.type === removeType && matchesRevision(mark),
                ),
            })
          : new Map<number, number>();
      const tr = state.tr;
      // Joins below re-read the runs they move themselves.
      tr.setMeta(JOINED_RUNS_RESTYLED_META, true);
      /** Content this resolution takes out, and what each removal keeps. */
      const deleteRanges: RemovedRange[] = [];
      /** The boundaries removed content left standing, for the range repair. */
      const keptBoundaries = new Set<PMNode>();
      /** What this resolution resolves, for what removed content keeps. */
      const selection: RevisionSelection = {
        mode,
        resolves: (revisionId) =>
          revisionSet === null || revisionSet.has(revisionId),
      };
      const removeRange = (
        rangeFrom: number,
        rangeTo: number,
        removed: RemovedSubtree,
      ): void => {
        deleteRanges.push({
          from: rangeFrom,
          to: rangeTo,
          removals: [removed],
        });
      };
      const resolveCopyOriginAttrs = (source: PMNode, position: number) => {
        const current = tr.doc.nodeAt(position);
        if (!current) return;
        const attrs = copyOriginBoundaryAttrsAfterResolution({
          source,
          current,
          baseAttrs: current.attrs,
          selection,
        });
        if (attrs === null) {
          panic("A copied boundary owner cannot be rejected outside removal.");
        }
        if (attrs !== current.attrs) {
          tr.setNodeMarkup(position, undefined, attrs);
        }
      };
      /** Where resolved inline content began and ended, in `state.doc`. */
      const resolvedBoundaries: number[] = [];
      const pPrMarkOps: PPrMarkOp[] = [];
      const tableRowStructuralOps: TableRowStructuralOp[] = [];
      const tableCellStructuralOps: TableCellStructuralOp[] = [];
      let removedSectionEndpointCount = 0;
      const removedSectionReferences: RemovedSectionReference[] = [];
      const countedSectionEndpoints = new Set<PMNode>();
      const countRemovedSectionEndpoint = (paragraph: PMNode): void => {
        if (
          !ownsSectionEndpoint(paragraph) ||
          countedSectionEndpoints.has(paragraph)
        )
          return;
        countedSectionEndpoints.add(paragraph);
        removedSectionEndpointCount++;
        removedSectionReferences.push(...sectionReferencesOf(paragraph));
      };
      // A selective rejection restores properties on their current carrier.
      // If its break still goes, the eventual survivor must read them too.
      const carryRestoredProperties = (position: number): void => {
        if (mode !== "reject" || revisionSet === null) return;
        const source = tr.doc.nodeAt(position);
        const mark = source ? expectParagraphAttrs(source).pPrMark : null;
        if (
          !source ||
          (mark?.kind !== "del" && mark?.kind !== "moveFrom") ||
          revisionSet.has(mark.info.id)
        )
          return;
        const survivor = paragraphLeftAfter({
          doc: tr.doc,
          paragraphPos: position,
        });
        if (survivor === null) return;
        carryParagraphProperties({
          tr,
          position: survivor,
          source,
          styleResolver,
          numbering,
          revision: mark.info,
        });
      };

      state.doc.nodesBetween(from, to, (node, pos): boolean => {
        if (node.type.name === "paragraph") {
          const op = collectPPrMarkOp(node, pos, from, to, mode, revisionSet);
          if (op) {
            pPrMarkOps.push(op);
          }

          const boundaryCovered = rangeCoversParagraphBoundary(
            from,
            to,
            pos,
            node,
          );
          const nextAttrs = resolveParagraphChangeAttrs({
            node,
            boundaryCovered,
            mode,
            revisionSet,
            styleResolver,
            numbering,
          });

          if (nextAttrs) {
            if (nextAttrs["_sectionProperties"] == null) {
              countRemovedSectionEndpoint(node);
            }
            const styleChanged = nextAttrs["styleId"] !== node.attrs["styleId"];
            if (styleChanged && styleResolver) {
              setParagraphAttrsWithRebasedRunFormatting({
                nextAttrs,
                paragraphPosition: pos,
                styleResolver,
                tr,
              });
            } else {
              tr.setNodeMarkup(pos, undefined, nextAttrs);
            }
            carryRestoredProperties(pos);
          }

          return true;
        }

        if (
          node.type.name === "tableRow" &&
          rangeCoversNode(from, to, pos, node)
        ) {
          const op = collectTableRowStructuralOp(node, pos, mode, revisionSet);
          if (op) {
            tableRowStructuralOps.push(op);
          }
        }
        if (
          (node.type.name === "tableCell" ||
            node.type.name === "tableHeader") &&
          rangeCoversNode(from, to, pos, node)
        ) {
          tableCellStructuralOps.push(
            ...collectTableCellStructuralOps(node, pos, mode, revisionSet),
          );
        }

        // Property revisions carried on the table / row / cell node attrs. A
        // row carries two of them (`w:trPrChange` and `w:tblPrExChange`), so
        // the sites are a list, not one attr per node type.
        const nodeAttrSites = nodePropertyRevisionSites(node.type.name).filter(
          (site) => site.resolution === "node-attrs",
        );
        if (nodeAttrSites.length > 0) {
          if (rangeCoversNode(from, to, pos, node)) {
            for (const site of nodeAttrSites) {
              const nextAttrs = resolveNodePropertyChangeAttrs({
                node: tr.doc.nodeAt(pos) ?? node,
                site,
                mode,
                revisionSet,
              });
              if (nextAttrs) {
                tr.setNodeMarkup(pos, undefined, nextAttrs);
                markStructuralChange(tr);
              }
            }
          }
          return true;
        }
        // Text and tracked run carriers (including zero-width structural atoms) can carry
        // tracked-change marks; widen the visitor so rejecting an inserted
        // picture removes it like inserted text. eigenpal #641.
        if (!node.isInline) {
          return true;
        }
        const nodeEnd = pos + node.nodeSize;
        const rangeFrom = Math.max(from, pos);
        const rangeTo = Math.min(to, nodeEnd);

        // The control goes with a revision that encloses it; one over its
        // content leaves it standing, emptied.
        if (node.type.name === INLINE_CONTENT_CONTROL_NODE_NAME) {
          if (!rangeCoversNode(from, to, pos, node)) {
            return true;
          }
          const resolvesRevision = (revisionId: number): boolean =>
            revisionSet === null || revisionSet.has(revisionId);
          if (
            resolutionRemovesControl({
              control: node,
              removeType,
              resolves: resolvesRevision,
            })
          ) {
            removeRange(pos, nodeEnd, removeResolvedSubtree(node, selection));
            return false;
          }
          const released = withoutResolvedEnclosures(node, resolvesRevision);
          if (released) {
            tr.setNodeMarkup(pos, undefined, released);
          }
          return true;
        }

        const runPropertyChangeMark = node.marks.find(
          (mark) => mark.type.name === "runPropertyChange",
        );
        const removesNode =
          removeType !== undefined &&
          node.marks.some(
            (mark) => mark.type === removeType && matchesRevision(mark),
          );
        const resolvesNode =
          removesNode ||
          runPropertyChangeMark !== undefined ||
          node.marks.some(
            (mark) =>
              (mark.type === insertionType || mark.type === deletionType) &&
              matchesRevision(mark),
          );
        if (resolvesNode) {
          resolvedBoundaries.push(rangeFrom, rangeTo);
        }
        if (runPropertyChangeMark) {
          resolveRunPropertyChange({
            tr,
            node,
            from: rangeFrom,
            to: rangeTo,
            mark: runPropertyChangeMark,
            mode,
            revisionSet,
            styleResolver,
          });
        }

        const nestedMark = node.marks.find(
          (mark) =>
            (mark.type === insertionType || mark.type === deletionType) &&
            (expectTrackedChangeMarkAttrs(mark)._docxRevisionAncestors
              ?.length ?? 0) > 0,
        );
        if (nestedMark) {
          const path = trackedRevisionPathOf(nestedMark);
          const selected = (layer: TrackedRevisionAncestor): boolean =>
            revisionSet === null || revisionSet.has(layer.revisionId);
          let nestedContentRemoved = false;
          if (path.some(selected)) {
            if (
              path.some(
                (layer) =>
                  selected(layer) && revisionLayerRemovesContent(layer, mode),
              )
            ) {
              removeRange(
                rangeFrom,
                rangeTo,
                removeResolvedSubtree(node, selection),
              );
              nestedContentRemoved = true;
              // The removal answers for everything inside the node.
              if (!node.isLeaf) return false;
            } else {
              tr.removeMark(rangeFrom, rangeTo, nestedMark);
              const retained = retainedTrackedRevisionMark(
                nestedMark,
                (layer) => !selected(layer),
              );
              if (retained) tr.addMark(rangeFrom, rangeTo, retained);
            }
          }
          const current = tr.doc.nodeAt(pos);
          if (current) {
            const attrs = resolvedPreservedXmlAttrs(node, current.marks);
            if (attrs !== node.attrs) tr.setNodeMarkup(pos, undefined, attrs);
          }
          if (!nestedContentRemoved) resolveCopyOriginAttrs(node, pos);
          return true;
        }

        if (removesNode) {
          removeRange(
            rangeFrom,
            rangeTo,
            removeResolvedSubtree(node, selection),
          );
          // The removal answers for everything inside the node.
          if (!node.isLeaf) return false;
        }

        if (
          !removesNode &&
          keepType &&
          node.marks.some(
            (mark) => mark.type === keepType && matchesRevision(mark),
          )
        ) {
          const retained = node.marks.filter(
            (mark) => !(mark.type === keepType && matchesRevision(mark)),
          );
          const attrs = resolvedPreservedXmlAttrs(node, retained);
          if (attrs !== node.attrs) tr.setNodeMarkup(pos, undefined, attrs);
        }

        if (!removeKeptMarksInBulk) {
          for (const mark of node.marks) {
            if (keepType && mark.type === keepType && matchesRevision(mark)) {
              tr.removeMark(rangeFrom, rangeTo, mark);
            }
          }
        }
        if (!removesNode) resolveCopyOriginAttrs(node, pos);
        return true;
      });

      if (removeKeptMarksInBulk) {
        tr.removeMark(from, to, keepType);
      }

      let rangesToDelete = deleteRanges;
      if (revisionSet === null) {
        // Adjacent inline ranges have no paragraph boundary between them, so
        // one replacement has the same mapping outside the deleted content.
        const coalescedDeleteRanges: RemovedRange[] = [];
        for (const range of deleteRanges) {
          const previous = coalescedDeleteRanges.at(-1);
          if (previous && range.from <= previous.to) {
            previous.to = Math.max(previous.to, range.to);
            previous.removals = [...previous.removals, ...range.removals];
            continue;
          }
          coalescedDeleteRanges.push({ ...range });
        }
        rangesToDelete = coalescedDeleteRanges;
      }
      for (const range of rangesToDelete.toReversed()) {
        const start = tr.doc.resolve(range.from);
        const field = start.parent;
        if (
          keptIn(range).length === 0 &&
          field.type.name === "structuredField" &&
          start.parentOffset === 0 &&
          range.to === start.end()
        ) {
          // Canonicalize before the replacement fitter supplies a synthetic
          // child for a required inline carrier whose result was removed.
          const position = start.before();
          const empty = canonicalFieldNode(field.copy(Fragment.empty));
          tr.replaceWith(position, position + field.nodeSize, empty);
          continue;
        }
        const kept = keptIn(range);
        for (const boundary of kept) keptBoundaries.add(boundary);
        if (kept.length === 0) tr.delete(range.from, range.to);
        else tr.replaceWith(range.from, range.to, kept);
      }
      // A text box goes with its anchor, and before any paragraph mark joins:
      // the box stands between its paragraph and the next one.
      if (
        deleteTextBoxesAnchoredAt(
          tr,
          droppedTextBoxAnchorIds(state.doc, tr.doc),
        )
      ) {
        markStructuralChange(tr);
      }

      // Where a node this resolution reads stands now, or null once the
      // removal of a block around it (a table, a row) took it: mapped on, its
      // position would name whatever follows the removed block.
      const standingAt = (position: number): number | null => {
        const mapped = tr.mapping.mapResult(position, 1);
        return mapped.deletedAcross ? null : mapped.pos;
      };

      // Invariant: table rows and cells resolve BEFORE paragraph marks. A join
      // must see the sibling the resolved document has, so a table this
      // resolution removes cannot stand between a break and the paragraph it
      // joins (a rejected split around a rejected inserted table).
      tableRowStructuralOps.sort((left, right) => right.rowPos - left.rowPos);
      let resolvedTableRowStructure = false;
      for (const op of tableRowStructuralOps) {
        const mappedPos = standingAt(op.rowPos);
        if (mappedPos === null) continue;
        const row = tr.doc.nodeAt(mappedPos);
        if (!row || row.type.name !== "tableRow") {
          continue;
        }
        if (op.action === "clear") {
          tr.setNodeAttribute(mappedPos, op.attrName, null);
          // The row keeps its content, so the run-level half of the same
          // revision has to go with the row attribute: `keepType` is exactly
          // the mark kind whose row marker resolves by clearing.
          clearTableRowContentMarks({
            tr,
            rowPos: mappedPos,
            markType: keepType,
            revision: op.revision,
          });
          resolvedTableRowStructure = true;
          continue;
        }
        deleteTableRowAt(tr, mappedPos);
        resolvedTableRowStructure = true;
      }
      if (resolvedTableRowStructure) {
        markStructuralChange(tr);
      }

      tableCellStructuralOps.sort(
        (left, right) => right.cellPos - left.cellPos,
      );
      let resolvedTableCellStructure = false;
      let failedTableCellMergeResolution = false;
      const emptiedTables: RowsEmptiedInBatch = { pending: false };
      for (const op of tableCellStructuralOps) {
        const mappedPos = standingAt(op.cellPos);
        if (mappedPos === null) continue;
        const cell = tr.doc.nodeAt(mappedPos);
        if (
          !cell ||
          (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader")
        ) {
          continue;
        }
        if (op.type === "merge") {
          const resolved =
            op.source === "collapsed"
              ? resolveCollapsedTableCellMerge(
                  tr,
                  mappedPos,
                  op.mode,
                  op.revisionSet,
                )
              : resolveVisibleTableCellMerge({
                  tr,
                  cellPos: mappedPos,
                  mode: op.mode,
                  revisionSet: op.revisionSet,
                });
          if (!resolved) {
            failedTableCellMergeResolution = true;
            break;
          }
          resolvedTableCellStructure ||= resolved;
          continue;
        }
        if (op.action === "clear") {
          tr.setNodeAttribute(mappedPos, "cellMarker", null);
          resolvedTableCellStructure = true;
          continue;
        }
        deleteTableCellAt(tr, mappedPos, emptiedTables);
        resolvedTableCellStructure = true;
      }
      if (failedTableCellMergeResolution) {
        return false;
      }
      removeRowsWithoutCellsAfterBatch(tr, emptiedTables);
      if (resolvedTableCellStructure) {
        markStructuralChange(tr);
      }

      // Process paragraph-mark ops from end → start so earlier positions stay
      // valid as later paragraphs collapse. Map every position through the
      // accumulated transaction so the inline deletes above don't desync the
      // attr writes or joins below.
      pPrMarkOps.sort((a, b) => b.paragraphPos - a.paragraphPos);
      // Tables whose run-in cell was pending when this resolution began: a
      // paragraph before one keeps its place, as the bulk resolver decides
      // before it resolves any table.
      const pendingTables: number[] = [];
      state.doc.descendants((node, position) => {
        if (node.type.spec["tableRole"] === "table" && runInCellPending(node)) {
          pendingTables.push(position);
        }
        return !node.isTextblock;
      });
      // A mark that goes must not leave a copy of its section record behind.
      if (pPrMarkOps.some(({ action }) => action === "join")) {
        for (const position of sectionRecordCopyPositions(tr.doc)) {
          tr.setNodeAttribute(position, "_sectionProperties", null);
        }
      }
      // Where two paragraphs' runs meet after a join: their runs stay apart.
      const joinSeams: { position: number; step: number }[] = [];
      // Each joined paragraph, with the suggested breaks the join displaced.
      const joinedParagraphs: {
        position: number;
        step: number;
        displaced: AddedParagraphMark[];
      }[] = [];
      for (const op of pPrMarkOps) {
        const mappedPos = standingAt(op.paragraphPos);
        if (mappedPos === null) continue;
        const paragraph = tr.doc.nodeAt(mappedPos);
        if (!paragraph || paragraph.type.name !== "paragraph") {
          continue;
        }
        if (op.action === "clear") {
          tr.setNodeAttribute(mappedPos, "pPrMark", null);
          continue;
        }
        if (op.action === "settle") {
          tr.setNodeAttribute(mappedPos, "pPrMark", op.mark ?? null);
          continue;
        }
        // The text boxes drawn in the paragraph are blocks between it and the
        // next one; they follow the joined paragraph instead.
        if (moveTextBoxesPastNextParagraph(tr, mappedPos)) {
          markStructuralChange(tr);
        }
        const joinPos = mappedPos + paragraph.nodeSize;
        let nextPos = joinPos;
        const boundaries: PMNode[] = [];
        while (nextPos < tr.doc.content.size) {
          const boundary = tr.doc.nodeAt(nextPos);
          // A whole suggested block this rejection takes out is gone before
          // the break joins: the words run into the block that stays, under
          // its mark, never into the one rejecting removes.
          if (
            boundary !== null &&
            boundary.attrs["pPrMark"] == null &&
            rejectionRemovesWholeNode(boundary, {
              mode,
              resolves: (revisionId) => revisionSet?.has(revisionId) === true,
            }) &&
            !endsItsContainer(tr.doc.resolve(nextPos))
          ) {
            tr.delete(nextPos, nextPos + boundary.nodeSize);
            continue;
          }
          if (boundary === null || !isBlockBoundaryNode(boundary)) break;
          boundaries.push(boundary);
          nextPos += boundary.nodeSize;
        }
        const nextNode =
          nextPos < tr.doc.content.size ? tr.doc.nodeAt(nextPos) : null;
        const joinable = nextNode?.type.name === paragraph.type.name;
        let follows: "paragraph" | "table" | "block" | "nothing" = "nothing";
        if (joinable) follows = "paragraph";
        else if (nextNode?.type.spec["tableRole"] === "table")
          follows = "table";
        else if (nextNode !== null && nextNode.type.name !== "textBox")
          follows = "block";
        const resolved = tr.doc.resolve(mappedPos);
        const preceding = resolved.parent.maybeChild(resolved.index() - 1);
        const retainedBoundaries = paragraphRemainsOf(paragraph);
        const outcome = paragraphMarkOutcome({
          follows,
          holdsContent: paragraphContentOf(paragraph),
          markWasAdded: op.markWasAdded,
          // Before a standing block, the container's actual final child must
          // remain a paragraph. At its end, removal makes the preceding child
          // final; only a content control may end with a table.
          containerStillEndsWithParagraph:
            follows !== "nothing"
              ? resolved.parent.lastChild?.type.name === "paragraph"
              : preceding?.type.name === "paragraph" ||
                (preceding !== null &&
                  resolved.parent.type.name === "blockSdt"),
          followingTableCanReceiveRunIn:
            follows === "table" &&
            boundaries.length === 0 &&
            !pendingTables.some(
              (position) => tr.mapping.map(position) === nextPos,
            ) &&
            !ownsSectionEndpoint(paragraph),
          needsFollowingCommentHost:
            follows === "table" &&
            preceding?.type.name !== "paragraph" &&
            retainedBoundaries.length > 0,
        });
        const ranIn =
          outcome.kind === "run-in"
            ? runParagraphIntoTable({
                tr,
                paragraphPos: mappedPos,
                styleResolver,
              })
            : null;
        if (ranIn) {
          tr.setMeta(JOINED_RUNS_RESTYLED_META, true);
          markParagraphPropertySourceTransfers(tr, [ranIn.transfer]);
          continue;
        }
        if (!joinable) {
          // Nothing to join with: the next sibling is a table or opaque block,
          // or the paragraph ends its container — a body, cell, header, note or text box
          // each end with one, and the position then lands on a boundary where
          // `join` would merge the containers and silently lose rows.
          //
          // A paragraph before a TABLE is still gone once its break and its
          // content are both resolved away: it is simply not there any more, so
          // it is removed rather than left blank.
          //
          // At the container's END the two directions part. A break that was
          // ADDED there added the paragraph it ends, so taking the addition
          // back removes the paragraph and the container ends where it did
          // before. A break that was REMOVED there cannot be honoured at all:
          // it says "join with the paragraph after this one" and there is
          // none, so the paragraph keeps its place and loses only the mark's
          // revision. That is also what makes a redline that deleted such a
          // mark fail its own round trip rather than resolve into a document
          // no consumer would have reached from it.
          //
          // A container must contain a paragraph either way, so its parent's
          // only child stays blank whatever its mark says.
          //
          // Section properties live on the removed paragraph mark and go
          // with it, as they do for a join.
          //
          // What the paragraph held that is no content goes to the
          // paragraph before it (the engine's `paragraphRemains`), as
          // deleting it directly leaves it. With no paragraph there to take
          // it, the paragraph stays.
          const remains =
            outcome.kind === "remove" ? paragraphRemainsOf(paragraph) : null;
          const toPrevious = remains ?? [];
          if (
            remains !== null &&
            (toPrevious.length === 0 || preceding?.type.name === "paragraph")
          ) {
            countRemovedSectionEndpoint(paragraph);
            tr.delete(mappedPos, mappedPos + paragraph.nodeSize);
            if (toPrevious.length > 0) tr.insert(mappedPos - 1, toPrevious);
            continue;
          }
          tr.setNodeAttribute(mappedPos, "pPrMark", null);
          // A section break an added mark ended never stood in the original:
          // it goes with the mark, even where the paragraph stays.
          if (op.markWasAdded && ownsSectionEndpoint(paragraph)) {
            countRemovedSectionEndpoint(paragraph);
            tr.setNodeAttribute(mappedPos, "_sectionProperties", null);
          }
          continue;
        }
        // The paragraph the reader is left with is the NEXT one, whether this
        // one still has words or not: its mark ends the joined paragraph, and
        // in OOXML a paragraph's properties live on its mark. PM's `join`
        // keeps the first node's attrs, so the next one's are set explicitly.
        //
        // The next paragraph's own `pPrMark` travels with its attrs: it is a
        // different revision, and resolving this one must not resolve it.
        //
        // Section properties live on the removed paragraph mark and go with
        // it: the joined paragraph ends with the next one's section break.
        try {
          const destination = tr.doc.nodeAt(nextPos);
          const next =
            destination?.type.name === "paragraph" ? destination : nextNode;
          const displaced = suggestedBreaksAJoinDisplaces(paragraph, next);
          joinAtParagraphMark({
            tr,
            paragraphPos: mappedPos,
            paragraph,
            next,
            styleResolver,
          });
          joinSeams.push({
            position: mappedPos + 1 + paragraph.content.size,
            step: tr.steps.length,
          });
          joinedParagraphs.push({
            position: mappedPos,
            step: tr.steps.length,
            displaced,
          });
          countRemovedSectionEndpoint(paragraph);
        } catch {
          // PM rejects the join if the two blocks aren't structurally
          // compatible (e.g. paragraph followed by a table). Leaving the
          // marker is the safe fallback.
          continue;
        }
        if (op.markWasAdded) carryRestoredProperties(mappedPos);
      }
      if (!placeSuggestedBreaksOfJoins(tr, joinedParagraphs)) {
        return false;
      }

      if (removedSectionEndpointCount > 0) {
        markTrackedSectionEndpointRemoval(tr, {
          sourceDoc: state.doc,
          removedEndpointCount: removedSectionEndpointCount,
          removedReferences: removedSectionReferences,
        });
      }

      if (!terminalTableDeletionIsPending(tr)) {
        resolveTerminalTableReviewCarrier(tr, mode);
      }

      // The pieces a revision split off its run are one run again; runs that
      // only meet because their paragraphs joined are not pieces of one run.
      const seams = new Set(
        joinSeams.map(({ position, step }) =>
          tr.mapping.slice(step).map(position),
        ),
      );
      rejoinRunsAt({
        tr,
        boundaries: resolvedBoundaries
          .map((position) => tr.mapping.map(position))
          .filter((position) => !seams.has(position)),
        styleResolver,
      });

      const resolvedFields: { node: PMNode; position: number }[] = [];
      tr.doc.descendants((node, position) => {
        const canonical = canonicalFieldNode(node);
        if (canonical !== node)
          resolvedFields.push({ node: canonical, position });
      });
      for (const { node, position } of resolvedFields.toReversed()) {
        const field = tr.doc.nodeAt(position);
        if (!field) continue;
        if (field.type === node.type)
          tr.setNodeMarkup(position, undefined, node.attrs, node.marks);
        else tr.replaceWith(position, position + field.nodeSize, node);
      }
      // A range that loses all covered content is removed rather than
      // converted to a point. Point-authored comments never enter `points`.
      if (points.size > 0) {
        const doomed = new Set(points.keys());
        removeCommentAnchors(tr, doomed);
      }
      normalizeKeptRangeBoundaries(tr, keptBoundaries);
      finishResolution(state, tr);
      if (tr.steps.length > 0) {
        dispatch(tr);
      }
    }
    return true;
  };
}

/**
 * The receiver belongs to the deleted terminal table, not every revision in
 * the document. Keep it through independently resolved changes, then resolve
 * it as soon as that table's deletion has been decided.
 */
const terminalTableDeletionIsPending = (tr: Transaction): boolean => {
  let terminalTable: PMNode | null = null;
  for (let index = tr.doc.childCount - 1; index >= 0; index--) {
    const node = tr.doc.child(index);
    if (node.type.name === "table") {
      terminalTable = node;
      break;
    }
  }
  if (terminalTable === null) {
    return false;
  }
  let pending = false;
  terminalTable.descendants((node) => {
    pending ||=
      (node.type.name === "tableRow" &&
        isTableRowRevisionAttr(node.attrs["trDel"])) ||
      (node.type.name !== "tableRow" &&
        node.marks.some((mark) => mark.type.name === "deletion"));
    return !pending;
  });
  return pending;
};

/** Resolve the Folio-only receiver Word needs after a terminal deleted table. */
const resolveTerminalTableReviewCarrier = (
  tr: Transaction,
  mode: ResolveMode,
): void => {
  let position = tr.doc.content.size;
  let carrier: PMNode | null = null;
  // Text boxes after it belong to paragraphs before it, wherever they stand.
  for (let index = tr.doc.childCount - 1; index >= 0; index--) {
    carrier = tr.doc.child(index);
    position -= carrier.nodeSize;
    if (carrier.type.name !== "textBox") break;
  }
  if (
    carrier?.type.name !== "paragraph" ||
    carrier.attrs["reviewCarrier"] !== REVIEW_CARRIERS.TERMINAL_TABLE
  ) {
    return;
  }
  // Its own changes still pending, it is not yet what they leave it as.
  let pending = carrier.attrs["pPrMark"] != null;
  carrier.forEach((child) => {
    pending ||= child.marks.some(({ type }) =>
      ["insertion", "deletion"].includes(type.name),
    );
  });
  if (pending) return;
  if (mode === "accept" || carrier.content.size > 0) {
    tr.setNodeAttribute(position, "reviewCarrier", undefined);
    return;
  }
  tr.delete(position, position + carrier.nodeSize);
};

type ResolveRunPropertyChangeOptions = {
  tr: Transaction;
  node: PMNode;
  from: number;
  to: number;
  mark: Mark;
  mode: "accept" | "reject";
  revisionSet: Set<number> | null;
  styleResolver: ReturnType<typeof getDocumentStyleResolver>;
};

const resolveRunPropertyChange = ({
  tr,
  node,
  from,
  to,
  mark,
  mode,
  revisionSet,
  styleResolver,
}: ResolveRunPropertyChangeOptions): void => {
  const { changes } = expectRunPropertyChangeMarkAttrs(mark);
  const matches = changes.filter(
    (change) => revisionSet === null || revisionSet.has(change.info.id),
  );
  if (matches.length === 0) {
    return;
  }

  tr.removeMark(from, to, mark);
  const remaining = changes.filter(
    (change) => revisionSet !== null && !revisionSet.has(change.info.id),
  );
  if (remaining.length > 0) {
    tr.addMark(from, to, mark.type.create({ changes: remaining }));
  }
  const emptyRuns =
    node.type.name === "field" || node.type.name === "structuredField"
      ? expectFieldAttrs(node)._docxEmptyResultRuns
      : undefined;
  const resolvedRuns = emptyRuns
    ? resolveEmptyFieldResultRuns({
        runs: emptyRuns,
        mode,
        revisionIds: revisionSet,
      })
    : null;
  if (resolvedRuns) {
    const current = tr.doc.nodeAt(from);
    if (!current) panic("A resolved field has no node at its position");
    tr.setNodeMarkup(from, undefined, {
      ...current.attrs,
      _docxEmptyResultRuns: resolvedRuns.runs,
    });
    if (resolvedRuns.visibleFormattingChanged) {
      const styleContext = paragraphRunStyleContextAt({
        doc: tr.doc,
        pos: from,
        styleResolver,
      });
      for (const currentMark of node.marks) {
        if (RUN_FORMATTING_MARK_NAMES.has(currentMark.type.name)) {
          tr.removeMark(from, to, currentMark.type);
        }
      }
      // A field keeps no run formatting record: only its marks rebuild.
      for (const nextMark of rejectedRunFormatting({
        node,
        paragraphContext: styleContext,
        previousFormatting: resolvedRuns.visibleFormatting,
        styleResolver,
      }).marks ?? []) {
        tr.addMark(from, to, nextMark);
      }
    }
    return;
  }
  if (mode === "accept") {
    return;
  }

  const rejected = rejectedRunFormatting({
    node,
    paragraphContext: paragraphRunStyleContextAt({
      doc: tr.doc,
      pos: from,
      styleResolver,
    }),
    previousFormatting: matches.at(0)?.previousFormatting,
    styleResolver,
  });
  if (rejected.atomRunFormatting !== null) {
    const current = tr.doc.nodeAt(from);
    if (!current) panic("A resolved run atom has no node at its position");
    tr.setNodeMarkup(from, undefined, {
      ...current.attrs,
      _docxRunFormatting: rejected.atomRunFormatting,
    });
  }
  if (rejected.marks === null) return;
  for (const currentMark of node.marks) {
    if (RUN_FORMATTING_MARK_NAMES.has(currentMark.type.name)) {
      tr.removeMark(from, to, currentMark.type);
    }
  }
  for (const previousMark of rejected.marks) {
    tr.addMark(from, to, previousMark);
  }
};

type PPrMarkOp = {
  paragraphPos: number;
  /** `settle` leaves the one record of a two-record mark still pending. */
  action: "clear" | "join" | "settle";
  /** The mark `settle` leaves. */
  mark?: ParagraphMarkChange | undefined;
  /**
   * Whether the mark said the break was ADDED. A break that was added can be
   * taken back out at a container's edge, because the paragraph it ends was
   * not there before it; one that was removed cannot, because that paragraph
   * is what the container ends with.
   */
  markWasAdded: boolean;
};

type TableRowStructuralOp = {
  rowPos: number;
  attrName: "trIns" | "trDel";
  action: "clear" | "remove";
  revision: TableRowRevisionAttr;
};

type TableCellStructuralOp =
  | {
      type: "membership";
      cellPos: number;
      action: "clear" | "remove";
    }
  | {
      type: "merge";
      cellPos: number;
      source: "visible" | "collapsed";
      mode: "accept" | "reject";
      revisionSet: Set<number> | null;
    };

function collectTableRowStructuralOp(
  node: PMNode,
  rowPos: number,
  mode: "accept" | "reject",
  revisionSet: Set<number> | null,
): TableRowStructuralOp | null {
  for (const attrName of ["trIns", "trDel"] as const) {
    const marker = node.attrs[attrName];
    if (!isTableRowRevisionAttr(marker)) {
      continue;
    }
    if (revisionSet !== null && !revisionSet.has(marker.revisionId)) {
      continue;
    }
    const keepsRow = (attrName === "trIns") === (mode === "accept");
    return {
      rowPos,
      attrName,
      action: keepsRow ? "clear" : "remove",
      revision: marker,
    };
  }
  return null;
}

type ClearTableRowContentMarksOptions = {
  tr: Transaction;
  rowPos: number;
  markType: MarkType | undefined;
  revision: TableRowRevisionAttr;
};

/**
 * Clear the run-level revision marks a row's structural revision wrote.
 *
 * A tracked row insertion or deletion is marked twice — on the row and around
 * every run in its cells — so resolving one half and leaving the other would
 * hand back a row whose text still reads as inserted (or struck through) after
 * the change was accepted. Both halves resolve in the caller's transaction.
 */
function clearTableRowContentMarks({
  tr,
  rowPos,
  markType,
  revision,
}: ClearTableRowContentMarksOptions): void {
  const row = tr.doc.nodeAt(rowPos);
  if (!row || !markType) {
    return;
  }
  const contentFrom = rowPos + 1;
  const rowRevision = {
    id: revision.revisionId,
    author: revision.author,
    date: revision.date,
  };
  const removals: { from: number; to: number; mark: Mark }[] = [];
  row.descendants((node, offset) => {
    if (!node.isInline) {
      return true;
    }
    for (const mark of node.marks) {
      if (
        mark.type === markType &&
        isTableRowRevisionCompanion({
          row: rowRevision,
          candidate: {
            id: mark.attrs["revisionId"],
            author: mark.attrs["author"],
            date: mark.attrs["date"],
          },
        })
      ) {
        const from = contentFrom + offset;
        removals.push({ from, to: from + node.nodeSize, mark });
      }
    }
    return false;
  });
  for (const removal of removals) {
    tr.removeMark(removal.from, removal.to, removal.mark);
  }
}

function collectTableCellStructuralOps(
  node: PMNode,
  cellPos: number,
  mode: "accept" | "reject",
  revisionSet: Set<number> | null,
): TableCellStructuralOp[] {
  const operations: TableCellStructuralOp[] = [];
  const marker = node.attrs["cellMarker"];
  if (
    isTableCellRevisionAttr(marker) &&
    (revisionSet === null || revisionSet.has(marker.info.revisionId))
  ) {
    if (marker.kind === "merge") {
      operations.push({
        type: "merge",
        cellPos,
        source: "visible",
        mode,
        revisionSet,
      });
    } else {
      const view = mode === "accept" ? "final" : "original";
      operations.push({
        type: "membership",
        cellPos,
        action: isTableCellRetainedInReviewView(marker.kind, view)
          ? "clear"
          : "remove",
      });
    }
  }
  if (hasMatchingCollapsedTableCellMerge(node, revisionSet)) {
    operations.push({
      type: "merge",
      cellPos,
      source: "collapsed",
      mode,
      revisionSet,
    });
  }
  return operations;
}

/** Record rows for a final sweep after every cell revision has resolved. */
function deleteTableCellAt(
  tr: Transaction,
  cellPos: number,
  emptied: RowsEmptiedInBatch,
): void {
  const cell = tr.doc.nodeAt(cellPos);
  if (
    !cell ||
    (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader")
  ) {
    return;
  }
  const resolved = tr.doc.resolve(cellPos);
  const row = resolved.parent;
  if (row.type.spec["tableRole"] !== "row") {
    return;
  }
  const tableDepth = resolved.depth - 1;
  const table = resolved.node(tableDepth);
  if (table.type.spec["tableRole"] !== "table") {
    return;
  }
  const tablePosition = resolved.before(tableDepth);
  const cellOffset = cellPos - resolved.start(tableDepth);
  const removedColumn = TableMap.get(table).findCell(cellOffset).left;
  // A cell that spans down is not its row's to take with it: removing the row
  // would move the cell into the row below and keep it. The cell goes alone,
  // and the row it leaves empty is removed once the rest of the column has
  // gone too, when the table is rectangular again.
  const spansDown = Number(cell.attrs["rowspan"]) > 1;
  const holdsOtherCells = !rowHoldsNoCells(row, cell);
  if (holdsOtherCells || spansDown) {
    if (!holdsOtherCells) {
      markRowEmptiedInBatch({ tr, rowPosition: resolved.start() - 1, emptied });
    }
    tr.delete(cellPos, cellPos + cell.nodeSize);
    reconcileOmittedGridSlots(tr, tablePosition, table, removedColumn);
    reconcileTableGridAfterColumnRemoval({
      tr,
      tablePosition,
      previousTable: table,
      removedColumn,
    });
    return;
  }
  if (table.childCount === 1) {
    deleteTableRowAt(tr, resolved.start() - 1);
    return;
  }
  deleteTableRowAt(tr, resolved.start() - 1);
  reconcileOmittedGridSlots(tr, tablePosition, table, removedColumn);
  reconcileTableGridAfterColumnRemoval({
    tr,
    tablePosition,
    previousTable: table,
    removedColumn,
  });
}

function collectPPrMarkOp(
  node: { attrs: Record<string, unknown>; nodeSize: number },
  pos: number,
  from: number,
  to: number,
  mode: "accept" | "reject",
  revisionSet: Set<number> | null,
): PPrMarkOp | null {
  if (!rangeCoversParagraphBoundary(from, to, pos, node)) {
    return null;
  }
  const pPrMark = node.attrs["pPrMark"];
  if (!isPPrMarkAttr(pPrMark)) {
    return null;
  }
  const deletion = pPrMark.deletion;
  if (deletion !== undefined) {
    // An added mark another author then deleted: accepting the deletion or
    // rejecting the insertion takes the break; the other decision on one
    // record leaves the other pending.
    const chosen = (id: number) => revisionSet === null || revisionSet.has(id);
    const insertionChosen = chosen(pPrMark.info.id);
    const deletionChosen = chosen(deletion.info.id);
    if (!insertionChosen && !deletionChosen) return null;
    if (mode === "accept" ? deletionChosen : insertionChosen)
      return { paragraphPos: pos, action: "join", markWasAdded: true };
    const { deletion: _resolved, ...added } = pPrMark;
    return {
      paragraphPos: pos,
      action: "settle",
      mark: mode === "accept" ? { kind: "del", info: deletion.info } : added,
      markWasAdded: true,
    };
  }
  if (revisionSet !== null && !revisionSet.has(pPrMark.info.id)) {
    return null;
  }
  // Accepting an added break, or rejecting a removed one, keeps the break
  // (clear the attr). The other two remove it (join with the next paragraph).
  const markWasAdded = paragraphMarkWasAdded(pPrMark.kind);
  const action: PPrMarkOp["action"] =
    markWasAdded === (mode === "accept") ? "clear" : "join";
  return { paragraphPos: pos, action, markWasAdded };
}

function rangeCoversParagraphBoundary(
  from: number,
  to: number,
  pos: number,
  node: { nodeSize: number },
): boolean {
  const boundaryFrom = pos + node.nodeSize - 1;
  const boundaryTo = pos + node.nodeSize;
  return from <= boundaryFrom && to >= boundaryTo;
}

/** Whether [from, to] fully covers the node — the range property-change cards
 * and accept-all / reject-all sweeps supply for table-level records. */
function rangeCoversNode(
  from: number,
  to: number,
  pos: number,
  node: { nodeSize: number },
): boolean {
  return from <= pos && to >= pos + node.nodeSize;
}

const ownsSectionEndpoint = (paragraph: PMNode): boolean =>
  expectParagraphAttrs(paragraph)._sectionProperties != null;

const sectionReferencesOf = (paragraph: PMNode): RemovedSectionReference[] => {
  const sectionProperties = expectParagraphAttrs(paragraph)._sectionProperties;
  if (!sectionProperties) {
    return [];
  }
  return [
    ...(sectionProperties.headerReferences ?? []).map(({ type, rId }) => ({
      part: "header" as const,
      type,
      relationshipId: rId,
    })),
    ...(sectionProperties.footerReferences ?? []).map(({ type, rId }) => ({
      part: "footer" as const,
      type,
      relationshipId: rId,
    })),
  ];
};

const isAddedPPrMarkAttr = (value: unknown): value is AddedParagraphMark =>
  isPPrMarkAttr(value) && (value.kind === "ins" || value.kind === "moveTo");

/** The transaction `command` would dispatch on `state`, or `null` when it would not apply. */
const attemptedTransaction = (
  command: Command,
  state: EditorState,
  view?: Parameters<Command>[2],
): Transaction | null => {
  let attempted: Transaction | null = null;
  const applies = command(
    state,
    (transaction) => {
      attempted = transaction;
    },
    view,
  );
  return applies ? attempted : null;
};

/**
 * `command`, refused when what it would do breaks a content control's lock
 * (`lockedControlViolationOfTransform`). A document without a locked control
 * runs it as it is.
 */
const honouringContentControlLocks =
  (command: Command): Command =>
  (state, dispatch, view) => {
    if (lockedControlSpans(state.doc).length === 0)
      return command(state, dispatch, view);
    const attempted = attemptedTransaction(command, state, view);
    if (attempted === null) return false;
    if (lockedControlViolationOfTransform(attempted) !== null) return false;
    dispatch?.(attempted);
    return true;
  };

function resolveChange(
  from: number,
  to: number,
  mode: "accept" | "reject",
  revisionIds?: readonly number[],
): Command {
  return honouringContentControlLocks(
    resolveChangeUnguarded(from, to, mode, revisionIds),
  );
}

/**
 * What the content-control locks say about resolving the tracked changes
 * between `from` and `to` (only those of `revisionIds`, when given): `null`
 * when they allow it or there is nothing there to resolve. A revision whose
 * resolution would change a content-locked control's content, or remove a
 * control locked against removal, stays pending.
 */
export const changeResolutionLockViolation = (
  state: EditorState,
  {
    from,
    to,
    mode,
    revisionIds,
  }: {
    from: number;
    to: number;
    mode: "accept" | "reject";
    revisionIds?: readonly number[];
  },
): LockedControlViolation | null => {
  if (lockedControlSpans(state.doc).length === 0) return null;
  const attempted = attemptedTransaction(
    resolveChangeUnguarded(from, to, mode, revisionIds),
    state,
  );
  return attempted === null
    ? null
    : lockedControlViolationOfTransform(attempted);
};

/** {@link changeResolutionLockViolation} for the revisions `revisionIds`, wherever they are. */
export const revisionResolutionLockViolation = (
  state: EditorState,
  revisionIds: number | readonly number[],
  mode: "accept" | "reject",
): LockedControlViolation | null => {
  const range = findAIEditRevisionRange(state, revisionIds);
  return range === null
    ? null
    : changeResolutionLockViolation(state, {
        ...range,
        mode,
        revisionIds:
          typeof revisionIds === "number" ? [revisionIds] : revisionIds,
      });
};

/**
 * Accept a tracked change at the given range.
 * - Insertion: remove mark, keep text
 * - Deletion: remove mark AND text
 */
export function acceptChange(from: number, to: number): Command {
  return resolveChange(from, to, "accept");
}

/**
 * Reject a tracked change at the given range.
 * - Insertion: remove mark AND text
 * - Deletion: remove mark, keep text
 */
export function rejectChange(from: number, to: number): Command {
  return resolveChange(from, to, "reject");
}

const resolveAllChanges =
  (mode: ResolveMode): Command =>
  (state, dispatch) => {
    if (!dispatch) return true;
    const result = resolveStateStory(state, mode);
    if (result.failed) return false;
    if (result.resolved.eq(state.doc)) return true;
    const tr = state.tr.step(
      new RevisionResolutionStep({
        beforeDoc: state.doc,
        afterDoc: result.resolved,
        steps: result.steps,
        positionMap: result.positionMap,
      }),
    );
    // The resolved ranges are positions in the resolved document: the range
    // repair's steps map them too.
    const resolvedAt = tr.steps.length;
    normalizeKeptRangeBoundaries(tr, result.keptBoundaries);
    // Resolution is a reproducible projection. Schema fillers, including cells
    // a table-repair plugin adds, need the same identities on every reading.
    requestDeterministicParaIds(tr, `resolve-all:${mode}`);
    markChangedParagraphRanges(tr, {
      ranges: finalRevisionParagraphRanges(result),
      mappingFrom: resolvedAt,
      replacesStepMapAt: 0,
    });
    markParagraphPropertySourceTransfers(tr, result.transfers);
    if (result.structural) markStructuralChange(tr);
    if (result.removedEndpointCount > 0)
      markTrackedSectionEndpointRemoval(tr, {
        sourceDoc: state.doc,
        removedEndpointCount: result.removedEndpointCount,
        removedReferences: result.removedReferences,
      });
    finishResolution(state, tr);
    dispatch(tr);
    return true;
  };

/**
 * Accept all tracked changes in the document. Returns `false` while pending
 * suggestions are present because this command cannot restage them.
 */
export function acceptAllChanges(): Command {
  const resolve = resolveAllChangesHonouringLocks("accept");
  return (state, dispatch, view) => {
    // This generic command has no host operation queue to replay. A mixed
    // story must use the headless reviewer's acceptAll path instead.
    if (getSuggestions(state).length > 0) return false;
    return resolve(state, dispatch, view);
  };
}

/**
 * Reject all tracked changes in the document.
 */
export function rejectAllChanges(): Command {
  return resolveAllChangesHonouringLocks("reject");
}

/**
 * The revisions of `state` that a content control's lock keeps from being
 * resolved in `mode`, each with the lock's verdict: see
 * {@link changeResolutionLockViolation}. Empty without a locked control.
 */
export const lockHeldRevisions = (
  state: EditorState,
  mode: "accept" | "reject",
): Map<number, LockedControlViolation> => {
  const held = new Map<number, LockedControlViolation>();
  if (lockedControlSpans(state.doc).length === 0) return held;
  for (const id of new Set(
    extractTrackedChanges(state).entries.flatMap(trackedChangeEntryRevisionIds),
  )) {
    const violation = revisionResolutionLockViolation(state, id, mode);
    if (violation !== null) held.set(id, violation);
  }
  return held;
};

/**
 * Resolve every tracked change a content control's lock allows. Without a
 * change the locks hold, the whole story resolves in one pass, as before;
 * otherwise each other change resolves on its own, in an editor view (a
 * transaction each), and the held ones stay pending as they are. Without a
 * view nothing resolves when any change is held.
 */
const resolveAllChangesHonouringLocks =
  (mode: ResolveMode): Command =>
  (state, dispatch, view) => {
    const resolveAll = resolveAllChanges(mode);
    if (lockedControlSpans(state.doc).length === 0)
      return resolveAll(state, dispatch, view);
    const attempted = attemptedTransaction(resolveAll, state, view);
    if (attempted === null) return false;
    if (lockedControlViolationOfTransform(attempted) === null) {
      dispatch?.(attempted);
      return true;
    }
    if (!dispatch) return true;
    if (!view) return false;
    const held = lockHeldRevisions(state, mode);
    const ids = new Set(
      extractTrackedChanges(state).entries.flatMap(
        trackedChangeEntryRevisionIds,
      ),
    );
    for (const id of [...ids].toReversed()) {
      if (held.has(id)) continue;
      const resolveOne =
        mode === "accept" ? acceptAIEditRevision(id) : rejectAIEditRevision(id);
      resolveOne(view.state, view.dispatch, view);
    }
    return true;
  };

/**
 * Resolve a complete state synchronously for a headless reader or writer.
 *
 * Uses the same single-step resolver as the interactive editor command.
 */
export function resolveAllChangesInHeadlessState(
  state: EditorState,
  mode: ResolveMode,
): EditorState {
  return projectAllChangesInHeadlessState(state, mode) ?? state;
}

/** The review transaction, or null when its content cannot be resolved. */
export const projectAllChangesInHeadlessState = (
  state: EditorState,
  mode: ResolveMode,
): EditorState | null => {
  let resolvedState = state;
  const resolved = resolveAllChanges(mode)(state, (transaction) => {
    resolvedState = state.apply(transaction);
  });
  return resolved ? resolvedState : null;
};

/**
 * Resolve a story and retain its position map for callers addressing the tracked source.
 */
export const resolveAllChangesInHeadlessStateWithMapping = (
  state: EditorState,
  mode: ResolveMode,
): { state: EditorState; mapping: Mapping } => {
  let resolvedState = state;
  let mapping = new Mapping();
  resolveAllChanges(mode)(state, (transaction) => {
    resolvedState = state.apply(transaction);
    mapping = transaction.mapping;
  });
  return { state: resolvedState, mapping };
};

/**
 * Accept the tracked-change marks belonging to an AI-edit operation.
 * Pass a single revisionId for inserts, standalone deletions, or formatting changes; or the
 * full id list for a replace (one id per side). Returns false when
 * none of the ids match anything in the doc.
 */
export function acceptAIEditRevision(
  revisionIds: number | readonly number[],
): Command {
  return (state, dispatch) => {
    const range = findAIEditRevisionRange(state, revisionIds);
    if (!range) {
      return false;
    }
    const ids = typeof revisionIds === "number" ? [revisionIds] : revisionIds;
    return resolveChange(range.from, range.to, "accept", ids)(state, dispatch);
  };
}

/**
 * Reject the tracked-change marks belonging to an AI-edit operation.
 * See {@link acceptAIEditRevision} for the id semantics.
 */
export function rejectAIEditRevision(
  revisionIds: number | readonly number[],
): Command {
  return (state, dispatch) => {
    const range = findAIEditRevisionRange(state, revisionIds);
    if (!range) {
      return false;
    }
    const ids = typeof revisionIds === "number" ? [revisionIds] : revisionIds;
    return resolveChange(range.from, range.to, "reject", ids)(state, dispatch);
  };
}

/**
 * Remove the comment marks of the suggestions `matches` selects, or, with
 * `keep`, clear their suggestion so they stay as ordinary comments. Mark
 * steps do not shift positions.
 */
const resolveSuggestedCommentMarks = (
  tr: Transaction,
  matches: (suggestionId: string) => boolean,
  keep: boolean,
): boolean => {
  const found: { from: number; to: number; mark: Mark }[] = [];
  tr.doc.descendants((node, pos) => {
    if (!node.isInline) return undefined;
    for (const mark of node.marks) {
      const suggestionId = suggestedCommentOf(mark);
      if (suggestionId !== null && matches(suggestionId))
        found.push({ from: pos, to: pos + node.nodeSize, mark });
    }
    return undefined;
  });
  for (const { from, to, mark } of found) {
    tr.removeMark(from, to, mark);
    if (keep)
      tr.addMark(
        from,
        to,
        mark.type.create({ ...mark.attrs, suggestionId: null }),
      );
  }
  return found.length > 0;
};

/**
 * Whether a rejection takes `node` out whole: a wholly suggested paragraph or
 * table whose insertion it rejects. The suggestion rejections remove it, and
 * a paragraph break the same rejection takes out joins past it.
 */
const rejectionRemovesWholeNode = (
  node: PMNode,
  selection: { mode: ResolveMode; resolves: (revisionId: number) => boolean },
): boolean => {
  const structural = readStructuralSuggestion(node);
  return (
    selection.mode === "reject" &&
    structural?.isNodeInsert === true &&
    selection.resolves(structural.revisionId)
  );
};

/** Whether the node at `at` is its parent's last child. */
const endsItsContainer = (at: ResolvedPos): boolean =>
  at.index() >= at.parent.childCount - 1;

export type AcceptSuggestionOptions = {
  author: string;
  /**
   * Date stamped on the resulting user tracked change (serialized as `w:date`);
   * defaults to now. A malformed / non-parseable value is normalized to now, and
   * a valid value is canonicalized to ISO 8601, so an invalid date can never be
   * written into the document.
   */
  date?: string;
};

/**
 * Normalize an optional acceptance date to a canonical ISO 8601 string. Absent
 * or unparseable input falls back to the current time — a malformed `w:date`
 * must never reach the serialized document (fail-safe boundary validation).
 */
const normalizeAcceptDate = (date: string | undefined): string => {
  if (date === undefined) {
    return new Date().toISOString();
  }
  const parsed = new Date(date);
  return Number.isNaN(parsed.getTime())
    ? new Date().toISOString()
    : parsed.toISOString();
};

type AddedParagraphMark = {
  kind: "ins" | "moveTo";
  info: {
    id: number;
    author?: unknown;
    date?: unknown;
    initials?: unknown;
    provenance?: unknown;
    suggestionId?: unknown;
  };
};

const acceptedParagraphMark = (
  node: PMNode,
  author: string,
  date: string,
): AddedParagraphMark | null => {
  if (node.type.name !== "paragraph") {
    return null;
  }
  const marker: unknown = node.attrs["_suggestedInsert"];
  if (
    typeof marker !== "object" ||
    marker === null ||
    !("revisionId" in marker)
  ) {
    return null;
  }
  if (typeof marker.revisionId !== "number") {
    return null;
  }
  const initials = "initials" in marker ? marker.initials : undefined;
  return {
    kind: "ins",
    info: {
      id: marker.revisionId,
      author,
      date,
      ...(typeof initials === "string" ? { initials } : {}),
    },
  };
};

/**
 * The node-attr patch that converts a suggested block/table revision into a
 * normal (`"user"`) one authored by the accepting user. Returns `null` when the
 * node carries no matching suggested revision. Position-stable (attr writes
 * only), so callers can apply it during a single descendants walk.
 */
const convertStructuralSuggestionAttrs = (
  node: PMNode,
  author: string,
  date: string,
): Record<string, unknown> | null => {
  const attrs = node.attrs;
  const userInfo = (marker: { revisionId?: unknown; info?: unknown }) => {
    const source = (marker.info ?? marker) as {
      revisionId?: unknown;
      initials?: unknown;
    };
    return {
      revisionId: source.revisionId,
      author,
      date,
      ...(typeof source.initials === "string"
        ? { initials: source.initials }
        : {}),
    };
  };

  // `_suggestedInsert` is only meaningful on paragraphs and tables (see
  // `readStructuralSuggestion`); ignore it on any other node type.
  const name = node.type.name;
  const insertMarker = attrs["_suggestedInsert"];
  if (
    (name === "paragraph" || name === "table") &&
    typeof insertMarker === "object" &&
    insertMarker !== null
  ) {
    if (name === "table") {
      // Whole inserted table has no OOXML tracked form → accept applies it
      // directly by clearing the suggestion marker (the table stays as content).
      return { ...attrs, _suggestedInsert: null };
    }
    // Inserted paragraph → real inserted-paragraph tracked change: mark the
    // paragraph break as `w:ins` (the inline runs are re-authored by the mark
    // pass). An accepted following suggestion may already have rotated its
    // own break onto this paragraph; keep that boundary while the caller
    // places this paragraph's break on the preceding free carrier.
    const paragraphMark = acceptedParagraphMark(node, author, date);
    if (paragraphMark) {
      return {
        ...attrs,
        _suggestedInsert: null,
        pPrMark: attrs["pPrMark"] ?? paragraphMark,
      };
    }
  }
  if (node.type.name === "tableRow") {
    if (readSuggestedMarker(attrs["trIns"])) {
      return {
        ...attrs,
        trIns: userInfo(attrs["trIns"] as { revisionId?: unknown }),
      };
    }
    if (readSuggestedMarker(attrs["trDel"])) {
      return {
        ...attrs,
        trDel: userInfo(attrs["trDel"] as { revisionId?: unknown }),
      };
    }
  }
  if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
    const cellMarker = attrs["cellMarker"] as
      { kind?: unknown; info?: unknown } | null | undefined;
    // Merge markers never carry suggestion provenance (cell merge/split is
    // `unsupportedMode`), so only ins/del markers can convert to user changes.
    if (
      cellMarker &&
      (cellMarker.kind === "ins" || cellMarker.kind === "del") &&
      readSuggestedMarker(cellMarker.info)
    ) {
      return {
        ...attrs,
        cellMarker: {
          kind: cellMarker.kind,
          info: userInfo(cellMarker as { info?: unknown }),
        },
      };
    }
  }
  return null;
};

type ReauthorParagraphPropertySuggestionsOptions = {
  node: PMNode;
  matchesSuggestion: (suggestionId: string) => boolean;
  author: string;
  date: string;
};

const reauthorParagraphPropertySuggestions = ({
  node,
  matchesSuggestion,
  author,
  date,
}: ReauthorParagraphPropertySuggestionsOptions):
  ParagraphPropertyChangeAttrs[] | null => {
  if (node.type.name !== "paragraph") {
    return null;
  }
  const propertyChanges = expectParagraphAttrs(node)._propertyChanges;
  if (!Array.isArray(propertyChanges)) {
    return null;
  }
  let changed = false;
  const next = propertyChanges.map((change) => {
    const suggested = readSuggestedParagraphPropertyChange(change);
    if (!suggested || !matchesSuggestion(suggested.suggestionId)) {
      return change;
    }
    changed = true;
    return Object.assign({}, change, {
      info: {
        ...change.info,
        author,
        date,
        provenance: "user" as const,
        suggestionId: null,
      },
    });
  });
  return changed ? next : null;
};

const canReauthorParagraphPropertySuggestions = (
  doc: PMNode,
  matchesSuggestion: (suggestionId: string) => boolean,
): boolean => {
  let valid = true;
  doc.descendants((node) => {
    if (!valid || node.type.name !== "paragraph") {
      return undefined;
    }
    const propertyChanges = expectParagraphAttrs(node)._propertyChanges;
    if (!Array.isArray(propertyChanges)) {
      return undefined;
    }
    let serializableCount = 0;
    for (const change of propertyChanges) {
      const suggested = readSuggestedParagraphPropertyChange(change);
      if (!suggested || matchesSuggestion(suggested.suggestionId)) {
        serializableCount++;
      }
    }
    if (serializableCount > 1) {
      valid = false;
      return false;
    }
    return undefined;
  });
  return valid;
};

type FinalParagraphRotationOptions = {
  tr: Transaction;
  shouldRotate: (node: PMNode, position: number) => boolean;
  fallbackAuthor?: string;
};

type AppendRotatedParagraphPropertyChangeOptions = {
  tr: Transaction;
  currentPosition: number;
  previous: PMNode;
  mark: {
    info: {
      id: number;
      author?: unknown;
      date?: unknown;
      initials?: unknown;
      provenance?: unknown;
      suggestionId?: unknown;
    };
  };
  fallbackAuthor?: string;
};

const appendRotatedParagraphPropertyChange = ({
  tr,
  currentPosition,
  previous,
  mark,
  fallbackAuthor,
}: AppendRotatedParagraphPropertyChangeOptions): boolean => {
  const current = tr.doc.nodeAt(currentPosition);
  if (!current || current.type.name !== "paragraph") {
    return false;
  }
  const author =
    typeof mark.info.author === "string" ? mark.info.author : fallbackAuthor;
  if (author === undefined) {
    return false;
  }
  const existing = expectParagraphAttrs(current)._propertyChanges;
  const suggestionId =
    mark.info.provenance === "suggested" &&
    typeof mark.info.suggestionId === "string"
      ? mark.info.suggestionId
      : null;
  // One saved property change per paragraph; a suggested one is not saved.
  if (
    suggestionId === null &&
    hasSerializableParagraphPropertyChange(existing)
  ) {
    return false;
  }
  tr.setNodeMarkup(currentPosition, undefined, {
    ...current.attrs,
    _propertyChanges: [
      ...(Array.isArray(existing) ? existing : []),
      {
        type: "paragraphPropertyChange",
        info: {
          id: mark.info.id,
          author,
          ...(typeof mark.info.date === "string"
            ? { date: mark.info.date }
            : {}),
          ...(typeof mark.info.initials === "string"
            ? { initials: mark.info.initials }
            : {}),
          // A suggested break's snapshot is part of the suggestion: held back
          // from the save and resolved with it.
          ...(suggestionId === null
            ? {}
            : { provenance: "suggested" as const, suggestionId }),
        },
        previousFormatting: paragraphPropertiesSnapshot(previous),
      } satisfies ParagraphPropertyChangeAttrs,
    ],
  });
  return true;
};

/**
 * Shift a final run of inserted paragraph marks one paragraph to the left.
 * The container-final paragraph stays markless, while every inserted
 * paragraph keeps a same-revision property snapshot that makes targeted
 * rejection restore the paragraph whose mark now carries its break.
 */
const rotateAddedFinalParagraphBreaks = ({
  tr,
  shouldRotate,
  fallbackAuthor,
}: FinalParagraphRotationOptions): boolean => {
  const paragraphTypeName =
    tr.doc.type.schema.nodes["paragraph"]?.name ?? "paragraph";
  const finals = finalParagraphsOf(tr.doc, paragraphTypeName).filter(
    ({ node, position }) => shouldRotate(node, position),
  );

  for (const { position: finalPosition } of finals) {
    const final = tr.doc.nodeAt(finalPosition);
    const finalMark = final?.attrs["pPrMark"];
    if (!final || !isAddedPPrMarkAttr(finalMark)) {
      continue;
    }
    tr.setNodeAttribute(finalPosition, "pPrMark", null);
    if (
      !shiftAddedParagraphBreakBefore({
        tr,
        paragraphPosition: finalPosition,
        mark: finalMark,
        ...(fallbackAuthor !== undefined && { fallbackAuthor }),
      })
    ) {
      return false;
    }
  }
  return true;
};

type PlaceAcceptedParagraphBreakOptions = {
  tr: Transaction;
  paragraphPosition: number;
  mark: AddedParagraphMark;
  fallbackAuthor: string;
};

type FirstRotatedPropertyChange =
  { type: "append" } | { type: "rebase"; revisionId: number };

type ShiftAddedParagraphBreakBeforeOptions = {
  tr: Transaction;
  paragraphPosition: number;
  mark: AddedParagraphMark;
  fallbackAuthor?: string;
  firstPropertyChange?: FirstRotatedPropertyChange;
};

type RebaseRotatedParagraphPropertyChangeOptions = {
  tr: Transaction;
  currentPosition: number;
  previous: PMNode;
  revisionId: number;
};

const rebaseRotatedParagraphPropertyChange = ({
  tr,
  currentPosition,
  previous,
  revisionId,
}: RebaseRotatedParagraphPropertyChangeOptions): boolean => {
  const current = tr.doc.nodeAt(currentPosition);
  if (!current || current.type.name !== "paragraph") {
    return false;
  }
  const existing = expectParagraphAttrs(current)._propertyChanges;
  if (!Array.isArray(existing)) {
    return false;
  }
  const matching = existing.filter(({ info }) => info.id === revisionId);
  if (matching.length !== 1) {
    return false;
  }
  tr.setNodeMarkup(currentPosition, undefined, {
    ...current.attrs,
    _propertyChanges: existing.map((change) =>
      change.info.id === revisionId
        ? Object.assign({}, change, {
            previousFormatting: paragraphPropertiesSnapshot(previous),
          })
        : change,
    ),
  });
  return true;
};

type RebaseAddedParagraphBreakChainOptions = {
  tr: Transaction;
  carrierPosition: number;
  previous: PMNode;
};

/**
 * Rebase every property change in one rotated paragraph-break chain to the
 * formatting of its original, leftmost carrier. A later targeted suggestion
 * may extend an already-rotated chain, so normalizing only the newly shifted
 * prefix would leave the serialized history dependent on acceptance order.
 */
const rebaseAddedParagraphBreakChain = ({
  tr,
  carrierPosition,
  previous,
}: RebaseAddedParagraphBreakChainOptions): boolean => {
  const carrier = tr.doc.nodeAt(carrierPosition);
  if (!carrier || carrier.type.name !== "paragraph") {
    return false;
  }
  const resolved = tr.doc.resolve(carrierPosition);
  let current = carrier;
  let nextPosition = carrierPosition + carrier.nodeSize;

  for (
    let index = resolved.index() + 1;
    index < resolved.parent.childCount;
    index++
  ) {
    const currentMark = current.attrs["pPrMark"];
    if (!isAddedPPrMarkAttr(currentMark)) {
      return true;
    }
    const sibling = resolved.parent.child(index);
    if (sibling.type.name !== carrier.type.name) {
      if (sibling.type.spec["tableRole"] === "table") {
        return false;
      }
      nextPosition += sibling.nodeSize;
      continue;
    }

    if (
      !rebaseRotatedParagraphPropertyChange({
        tr,
        currentPosition: nextPosition,
        previous,
        revisionId: currentMark.info.id,
      })
    ) {
      return false;
    }
    const next = tr.doc.nodeAt(nextPosition);
    if (!next || next.type.name !== carrier.type.name) {
      return false;
    }
    current = next;
    nextPosition += sibling.nodeSize;
  }

  return !isAddedPPrMarkAttr(current.attrs["pPrMark"]);
};

/**
 * Insert an added paragraph break immediately before `paragraphPosition`.
 * Existing added marks shift left as one chain until the first free carrier;
 * each shifted mark keeps a property snapshot on the paragraph to its right.
 */
const shiftAddedParagraphBreakBefore = ({
  tr,
  paragraphPosition,
  mark,
  fallbackAuthor,
  firstPropertyChange = { type: "append" },
}: ShiftAddedParagraphBreakBeforeOptions): boolean => {
  const paragraph = tr.doc.nodeAt(paragraphPosition);
  if (!paragraph || paragraph.type.name !== "paragraph") {
    return false;
  }
  const resolved = tr.doc.resolve(paragraphPosition);
  const carrier = addedBreakCarrierBefore({
    at: resolved,
    paragraphTypeName: paragraph.type.name,
  });
  if (!carrier) {
    return false;
  }

  const path = [{ node: paragraph, position: paragraphPosition }];
  let position = paragraphPosition;
  for (let index = resolved.index() - 1; index >= 0; index--) {
    const sibling = resolved.parent.child(index);
    position -= sibling.nodeSize;
    if (sibling.type.name !== paragraph.type.name) {
      if (sibling.type.spec["tableRole"] === "table") {
        break;
      }
      continue;
    }
    path.push({ node: sibling, position });
    if (position === carrier.position) {
      break;
    }
  }
  if (path.at(-1)?.position !== carrier.position) {
    return false;
  }

  const propertyChangeBaseline = carrier.node;
  let carriedMark = mark;
  for (let index = 1; index < path.length; index++) {
    const current = path[index - 1];
    const previous = path[index];
    if (!current || !previous) {
      return false;
    }
    const displacedMark = previous.node.attrs["pPrMark"];
    tr.setNodeAttribute(previous.position, "pPrMark", carriedMark);

    const propertyChangeApplied =
      index === 1 && firstPropertyChange.type === "rebase"
        ? rebaseRotatedParagraphPropertyChange({
            tr,
            currentPosition: current.position,
            previous: propertyChangeBaseline,
            revisionId: firstPropertyChange.revisionId,
          })
        : appendRotatedParagraphPropertyChange({
            tr,
            currentPosition: current.position,
            previous: propertyChangeBaseline,
            mark: carriedMark,
            ...(fallbackAuthor !== undefined && { fallbackAuthor }),
          });
    if (!propertyChangeApplied) {
      return false;
    }

    if (displacedMark == null) {
      return rebaseAddedParagraphBreakChain({
        tr,
        carrierPosition: carrier.position,
        previous: propertyChangeBaseline,
      });
    }
    if (!isAddedPPrMarkAttr(displacedMark)) {
      return false;
    }
    carriedMark = displacedMark;
  }
  return false;
};

/** A wholly suggested paragraph's break, as the inserted mark that records it. */
const suggestedParagraphBreakOf = (
  paragraph: PMNode,
): AddedParagraphMark | null => {
  const structural = readStructuralSuggestion(paragraph);
  if (structural?.kind !== "insertBlock") return null;
  const marker = paragraph.attrs["_suggestedInsert"] as {
    author?: unknown;
    date?: unknown;
    initials?: unknown;
  };
  return {
    kind: "ins",
    info: {
      id: structural.revisionId,
      author: marker.author,
      ...(typeof marker.date === "string" ? { date: marker.date } : {}),
      ...(typeof marker.initials === "string"
        ? { initials: marker.initials }
        : {}),
      provenance: "suggested",
      suggestionId: structural.suggestionId,
    },
  };
};

/**
 * The suggested paragraph breaks a join of `paragraph` into `next` leaves
 * with no mark to stand on.
 *
 * A wholly suggested paragraph's break is its own mark while that mark is
 * free. Where another revision's mark stands there (the break of a paragraph
 * inserted after it, or a deleted break it took over), the suggestion's break
 * is the free mark before it, which is where accepting the suggestion writes
 * it. A join that leaves the paragraph holding other words ends its
 * whole-paragraph marker, so such a break must then be written there: the
 * joined paragraph's mark is another revision's (the next side's own), or is
 * gone (the side whose mark the join removes). The next side's break with a
 * free mark is not among them: the joined mark records it (`joinedParagraphAttrs`).
 */
const suggestedBreaksAJoinDisplaces = (
  paragraph: PMNode,
  next: PMNode,
): AddedParagraphMark[] => {
  const left = readStructuralSuggestion(paragraph);
  const right = readStructuralSuggestion(next);
  if (
    left?.kind === "insertBlock" &&
    right?.kind === "insertBlock" &&
    left.suggestionId === right.suggestionId
  )
    return [];
  const displaced: AddedParagraphMark[] = [];
  for (const side of [paragraph, next]) {
    const suggested = suggestedParagraphBreakOf(side);
    if (suggested === null || side.attrs["pPrMark"] == null) continue;
    if (
      readSuggestedParagraphMark(side)?.suggestionId !==
      suggested.info.suggestionId
    )
      displaced.push(suggested);
  }
  return displaced;
};

/** The paragraph a join left at `position`, once `step` and every later step ran. */
const joinedParagraphAt = (
  tr: Transaction,
  position: number,
  step: number,
): number | null => {
  const at = tr.mapping.slice(step).map(position);
  if (tr.doc.nodeAt(at)?.type.name === "paragraph") return at;
  // A later join ran an earlier paragraph into this one.
  const resolved = tr.doc.resolve(at);
  return resolved.parent.type.name === "paragraph" ? resolved.before() : null;
};

/**
 * Write the suggested breaks the joins of one resolution left without a mark.
 *
 * A wholly suggested paragraph whose own mark another revision holds has its
 * break on the free mark before it: accepting the suggestion writes it there
 * ({@link placeAcceptedParagraphBreakBefore}), and rejecting the suggestion
 * takes that break out. A join that ends the paragraph's whole-paragraph
 * marker has to write the break there, or rejecting the suggestion leaves its
 * paragraph standing. A container's last paragraph carries no added break
 * either, so a suggested one a join wrote there moves back the same way.
 * After every join of the resolution, so no break lands on a mark it still
 * resolves.
 */
const placeSuggestedBreaksOfJoins = (
  tr: Transaction,
  joins: readonly {
    position: number;
    step: number;
    displaced: readonly AddedParagraphMark[];
  }[],
): boolean => {
  for (const { position: joinedAt, step, displaced } of joins) {
    const position = joinedParagraphAt(tr, joinedAt, step);
    const joined = position === null ? null : tr.doc.nodeAt(position);
    if (position === null || !joined) continue;
    const breaks: AddedParagraphMark[] = [...displaced];
    const own: unknown = joined.attrs["pPrMark"];
    const at = tr.doc.resolve(position);
    if (
      isAddedPPrMarkAttr(own) &&
      readSuggestedParagraphMark(joined) !== null &&
      paragraphEndsItsContainer(at, joined.type.name) &&
      addedBreakCarrierBefore({ at, paragraphTypeName: joined.type.name })
    ) {
      tr.setNodeAttribute(position, "pPrMark", null);
      breaks.unshift(own);
    }
    for (const mark of breaks) {
      if (
        !shiftAddedParagraphBreakBefore({
          tr,
          paragraphPosition: position,
          mark,
        })
      ) {
        return false;
      }
    }
  }
  return true;
};

/**
 * Place an accepted paragraph's break when its own break slot is taken: a
 * following accepted insertion has rotated its break onto the paragraph, or
 * the paragraph carries the break of the one it was inserted after, pending
 * deletion. Either way that boundary stays, and the accepted break goes on
 * the preceding free carrier.
 */
const placeAcceptedParagraphBreakBefore = ({
  tr,
  paragraphPosition,
  mark,
  fallbackAuthor,
}: PlaceAcceptedParagraphBreakOptions): boolean => {
  const paragraph = tr.doc.nodeAt(paragraphPosition);
  const displaced = paragraph?.attrs["pPrMark"];
  if (
    !paragraph ||
    paragraph.type.name !== "paragraph" ||
    !isPPrMarkAttr(displaced)
  ) {
    return false;
  }
  return shiftAddedParagraphBreakBefore({
    tr,
    paragraphPosition,
    mark,
    fallbackAuthor,
  });
};

type AddedBreakRotationCandidate = {
  position: number;
  revisionId: number;
};

/**
 * Remember non-final added breaks before suggested nodes are removed. If one
 * becomes final because of those removals, it must rotate with its own
 * revision rather than remain as an unresolvable final mark.
 */
const collectAddedBreakRotationCandidates = (
  doc: PMNode,
): AddedBreakRotationCandidate[] => {
  const paragraphTypeName =
    doc.type.schema.nodes["paragraph"]?.name ?? "paragraph";
  const finalPositions = new Set(
    finalParagraphsOf(doc, paragraphTypeName).map(({ position }) => position),
  );
  const candidates: AddedBreakRotationCandidate[] = [];
  doc.descendants((node, position) => {
    if (node.type.name !== paragraphTypeName || finalPositions.has(position)) {
      return undefined;
    }
    const mark = node.attrs["pPrMark"];
    if (isPPrMarkAttr(mark) && paragraphMarkWasAdded(mark.kind)) {
      candidates.push({ position, revisionId: mark.info.id });
    }
    return undefined;
  });
  return candidates;
};

type MappedRotationCandidatePredicateOptions = {
  candidates: readonly AddedBreakRotationCandidate[];
  tr: Transaction;
};

const mappedRotationCandidatePredicate = ({
  candidates,
  tr,
}: MappedRotationCandidatePredicateOptions): ((
  node: PMNode,
  position: number,
) => boolean) => {
  const revisionIdsByPosition = new Map<number, Set<number>>();
  for (const candidate of candidates) {
    const mapped = tr.mapping.mapResult(candidate.position, 1);
    if (mapped.deleted) {
      continue;
    }
    const mappedPosition = mapped.pos;
    const revisionIds =
      revisionIdsByPosition.get(mappedPosition) ?? new Set<number>();
    revisionIds.add(candidate.revisionId);
    revisionIdsByPosition.set(mappedPosition, revisionIds);
  }
  return (node, position) => {
    const mark = node.attrs["pPrMark"];
    return (
      isPPrMarkAttr(mark) &&
      paragraphMarkWasAdded(mark.kind) &&
      revisionIdsByPosition.get(position)?.has(mark.info.id) === true
    );
  };
};

type RemoveSuggestedInsertNodesOptions = {
  tr: Transaction;
  positions: readonly number[];
  resolvedRevisionIds: ReadonlySet<number>;
};

/**
 * Remove suggested nodes without removing another revision's paragraph break
 * that stands on one of them.
 *
 * A suggested paragraph whose own mark another revision holds has its break
 * on the free mark before it (see {@link placeSuggestedBreaksOfJoins}), so
 * rejecting it takes that break out and the other revision's stays: it moves
 * onto that mark. An added break (a following insertion's, rotated or placed
 * there) shifts back as any added break does, its property snapshot rebased
 * when a rotation recorded one; a removed break simply moves.
 */
const removeSuggestedInsertNodes = ({
  tr,
  positions,
  resolvedRevisionIds,
}: RemoveSuggestedInsertNodesOptions): boolean => {
  for (const position of positions.toSorted((left, right) => right - left)) {
    const node = tr.doc.nodeAt(position);
    if (!node) {
      continue;
    }
    const mark: unknown = node.attrs["pPrMark"];
    const standingBreak =
      node.type.name === "paragraph" &&
      isPPrMarkAttr(mark) &&
      !resolvedRevisionIds.has(mark.info.id)
        ? mark
        : null;
    tr.delete(position, position + node.nodeSize);
    if (!standingBreak) continue;
    if (!isAddedPPrMarkAttr(standingBreak)) {
      moveRemovedBreakBefore(tr, position, standingBreak);
      continue;
    }
    const following = tr.doc.nodeAt(position);
    const snapshots =
      following?.type.name === "paragraph"
        ? (expectParagraphAttrs(following)._propertyChanges ?? [])
        : [];
    const recordedSnapshot = snapshots.some(
      ({ info }) => info.id === standingBreak.info.id,
    );
    if (
      !shiftAddedParagraphBreakBefore({
        tr,
        paragraphPosition: position,
        mark: standingBreak,
        ...(recordedSnapshot && {
          firstPropertyChange: {
            type: "rebase",
            revisionId: standingBreak.info.id,
          },
        }),
      })
    ) {
      return false;
    }
  }
  return true;
};

/**
 * Put a removed break (`del`, `moveFrom`) that stood on a removed suggested
 * paragraph on the free mark of the paragraph right before it. With none
 * there, the break goes with the paragraph, as before.
 */
const moveRemovedBreakBefore = (
  tr: Transaction,
  position: number,
  mark: unknown,
): void => {
  const resolved = tr.doc.resolve(position);
  const before = resolved.nodeBefore;
  if (before?.type.name !== "paragraph" || before.attrs["pPrMark"] != null) {
    return;
  }
  tr.setNodeAttribute(position - before.nodeSize, "pPrMark", mark);
};

/**
 * Convert suggested marks AND block/table node revisions to normal (`"user"`)
 * tracked changes. `matchesSuggestion(id)` selects which suggestion to convert
 * (one id, or all).
 */
const acceptSuggestions = (
  matchesSuggestion: (suggestionId: string) => boolean,
  options: AcceptSuggestionOptions,
): Command => {
  return (state, dispatch) => {
    if (
      !canReauthorParagraphPropertySuggestions(state.doc, matchesSuggestion)
    ) {
      return false;
    }
    const insertionType = state.schema.marks["insertion"];
    const deletionType = state.schema.marks["deletion"];
    const runPropertyChangeType = state.schema.marks["runPropertyChange"];
    const date = normalizeAcceptDate(options.date);
    const tr = state.tr;
    const acceptedParagraphBreaks: {
      position: number;
      mark: AddedParagraphMark;
      hadBreak: boolean;
    }[] = [];
    let changed = false;

    // Mark steps AND setNodeAttribute do not shift positions, so the positions
    // read from `state.doc` stay valid across the accumulated steps.
    state.doc.descendants((node, pos) => {
      const structural = readStructuralSuggestion(node);
      let nextAttrs: Record<string, unknown> | null = null;
      if (structural && matchesSuggestion(structural.suggestionId)) {
        nextAttrs = convertStructuralSuggestionAttrs(
          node,
          options.author,
          date,
        );
        if (structural.kind === "insertBlock") {
          const mark = acceptedParagraphMark(node, options.author, date);
          if (!mark) {
            return undefined;
          }
          acceptedParagraphBreaks.push({
            position: pos,
            mark,
            hadBreak: node.attrs["pPrMark"] != null,
          });
        }
      }
      const propertyChanges = reauthorParagraphPropertySuggestions({
        node,
        matchesSuggestion,
        author: options.author,
        date,
      });
      if (propertyChanges) {
        nextAttrs = {
          ...(nextAttrs ?? node.attrs),
          _propertyChanges: propertyChanges,
        };
      }
      const paragraphMark: unknown = node.attrs["pPrMark"];
      if (isPPrMarkAttr(paragraphMark)) {
        // Each record of the mark the suggestion owns becomes the user's.
        const accepted = (info: ParagraphMarkChange["info"]) => {
          const suggestionId = suggestionIdOfMarkRecord(info);
          return suggestionId !== null && matchesSuggestion(suggestionId)
            ? {
                ...info,
                author: options.author,
                date,
                provenance: "user",
                suggestionId: null,
              }
            : null;
        };
        const own = accepted(paragraphMark.info);
        const { deletion: pendingDeletion } = paragraphMark;
        const deletion =
          pendingDeletion === undefined ? null : accepted(pendingDeletion.info);
        if (own !== null || deletion !== null) {
          nextAttrs = {
            ...(nextAttrs ?? node.attrs),
            pPrMark: {
              ...paragraphMark,
              ...(own === null ? {} : { info: own }),
              ...(deletion === null || pendingDeletion === undefined
                ? {}
                : { deletion: { ...pendingDeletion, info: deletion } }),
            },
          };
        }
      }
      if (nextAttrs) {
        tr.setNodeMarkup(pos, undefined, nextAttrs);
        changed = true;
      }
      if (!node.isInline) {
        return undefined;
      }
      const from = pos;
      const to = pos + node.nodeSize;
      for (const mark of node.marks) {
        const suggestionId = mark.attrs["suggestionId"];
        const acceptOwn =
          isSuggestionMark(mark) &&
          typeof suggestionId === "string" &&
          matchesSuggestion(suggestionId);
        if (mark.type === runPropertyChangeType && acceptOwn) {
          changed = true;
          tr.removeMark(from, to, mark);
          // Re-author each recorded change under the accepting user; keep the
          // rest of the change (previousFormatting, revision id) intact.
          const nextChanges: RunPropertyChange[] = [];
          for (const change of expectRunPropertyChangeMarkAttrs(mark).changes) {
            nextChanges.push({
              ...change,
              info: { ...change.info, author: options.author, date },
            });
          }
          tr.addMark(
            from,
            to,
            mark.type.create({
              changes: nextChanges,
              provenance: "user",
              suggestionId: null,
            }),
          );
          continue;
        }
        if (mark.type !== insertionType && mark.type !== deletionType) continue;
        let acceptedAncestor = false;
        const ancestors: TrackedRevisionAncestor[] = [];
        for (const ancestor of expectTrackedChangeMarkAttrs(mark)
          ._docxRevisionAncestors ?? []) {
          const { suggestionId: owner, ...record } = ancestor;
          if (owner === undefined || !matchesSuggestion(owner)) {
            ancestors.push(ancestor);
            continue;
          }
          acceptedAncestor = true;
          ancestors.push({ ...record, author: options.author, date });
        }
        if (!acceptOwn && !acceptedAncestor) continue;
        changed = true;
        tr.removeMark(from, to, mark);
        tr.addMark(
          from,
          to,
          mark.type.create({
            ...mark.attrs,
            ...(acceptOwn
              ? {
                  author: options.author,
                  date,
                  provenance: "user",
                  suggestionId: null,
                }
              : {}),
            ...(acceptedAncestor ? { _docxRevisionAncestors: ancestors } : {}),
          }),
        );
      }
      return undefined;
    });

    for (const accepted of acceptedParagraphBreaks) {
      if (
        accepted.hadBreak &&
        !placeAcceptedParagraphBreakBefore({
          tr,
          paragraphPosition: accepted.position,
          mark: accepted.mark,
          fallbackAuthor: options.author,
        })
      ) {
        return false;
      }
    }
    const directlyMarkedParagraphPositions = new Set(
      acceptedParagraphBreaks
        .filter(({ hadBreak }) => !hadBreak)
        .map(({ position }) => position),
    );
    if (
      directlyMarkedParagraphPositions.size > 0 &&
      !rotateAddedFinalParagraphBreaks({
        tr,
        shouldRotate: (_node, position) =>
          directlyMarkedParagraphPositions.has(position),
        fallbackAuthor: options.author,
      })
    ) {
      return false;
    }
    // A proposed comment stays as an ordinary comment.
    if (resolveSuggestedCommentMarks(tr, matchesSuggestion, true)) {
      changed = true;
    }

    if (!changed) {
      return false;
    }
    if (dispatch) {
      finishResolution(state, tr);
      dispatch(tr);
    }
    return true;
  };
};

/**
 * Accept one suggestion: convert its inline marks and block/table node
 * revisions into normal tracked changes authored by `author`, keeping revision
 * ids. A whole inserted table (which OOXML cannot track) is applied directly;
 * everything else becomes a tracked change. See {@link getSuggestions} for the
 * per-suggestion `appliedAs`.
 */
export function acceptSuggestion(
  suggestionId: string,
  options: AcceptSuggestionOptions,
): Command {
  return acceptSuggestions((id) => id === suggestionId, options);
}

/** Accept every suggestion in the document. */
export function acceptAllSuggestions(
  options: AcceptSuggestionOptions,
): Command {
  return acceptSuggestions(() => true, options);
}

/**
 * Reject one suggestion: inverse-apply it in a single transaction. Runs BOTH
 * phases so a heterogeneous group (one `suggestionId` spanning whole-node
 * inserts AND other edits) is fully resolved:
 *  1. inline marks and suggested `trIns`/`trDel`/`cellMarker` revisions are
 *     inverse-applied through {@link resolveChange} (delete suggested-inserted
 *     text/rows/cells, drop suggested deletions/formatting);
 *  2. any remaining whole-node inserts (paragraph/table) are then deleted,
 *     mapped through the accumulated transaction so positions stay valid.
 *
 * A paragraph whole-node insert carries inline marks with the same revision id,
 * so phase 1 empties its text and phase 2 removes the now-empty node; a whole
 * inserted table has no matching revision, so only phase 2 removes it.
 */
export function rejectSuggestion(suggestionId: string): Command {
  return (state, dispatch) => {
    const suggestions = collectSuggestions(state);
    const entry = suggestions.get(suggestionId);
    if (!entry) {
      return false;
    }
    // The reference resolver selects structural revisions by physical ID.
    // A repeated ID within one suggestion is fine, but a shared ID cannot
    // identify just one proposal, even when their inline spans are disjoint.
    for (const [otherId, other] of suggestions) {
      if (otherId === suggestionId) continue;
      for (const id of other.revisionIds) {
        if (entry.revisionIds.has(id)) return false;
      }
    }

    const removesWholeNode =
      entry.kinds.has("insertBlock") || entry.kinds.has("insertTable");
    const rotationCandidates = removesWholeNode
      ? collectAddedBreakRotationCandidates(state.doc)
      : [];

    // Phase 1: inverse-apply inline + structural revisions, capturing the
    // transaction so phase 2 can append node deletions to the same one.
    let tr: Transaction | null = null;
    if (entry.revisionIds.size > 0) {
      const ranges = mergeSegments(entry.segments);
      const from = ranges[0]?.from ?? 0;
      const to = ranges.at(-1)?.to ?? from;
      const resolved = resolveChangeUnguarded(from, to, "reject", [
        ...entry.revisionIds,
      ])(state, (transaction) => {
        tr = transaction;
      });
      if (!resolved) return false;
    }
    const workingTr: Transaction = tr ?? state.tr;

    // Phase 2: drop this suggestion's remaining whole-node inserts.
    const positions: number[] = [];
    const rejected = {
      mode: "reject",
      resolves: (revisionId: number) => entry.revisionIds.has(revisionId),
    } as const;
    // A whole node goes only when it is this suggestion's own insertion.
    workingTr.doc.descendants((node, pos) => {
      if (
        rejectionRemovesWholeNode(node, rejected) &&
        readStructuralSuggestion(node)?.suggestionId === suggestionId
      ) {
        positions.push(pos);
      }
      return undefined;
    });
    if (
      !removeSuggestedInsertNodes({
        tr: workingTr,
        positions,
        resolvedRevisionIds: entry.revisionIds,
      })
    ) {
      return false;
    }
    if (
      positions.length > 0 &&
      rotationCandidates.length > 0 &&
      !rotateAddedFinalParagraphBreaks({
        tr: workingTr,
        shouldRotate: mappedRotationCandidatePredicate({
          candidates: rotationCandidates,
          tr: workingTr,
        }),
      })
    ) {
      return false;
    }
    // A proposed comment goes with its suggestion.
    resolveSuggestedCommentMarks(workingTr, (id) => id === suggestionId, false);

    if (workingTr.steps.length === 0) {
      return false;
    }
    if (dispatch) {
      finishResolution(state, workingTr);
      dispatch(workingTr);
    }
    return true;
  };
}

/** Reject every suggestion in the document. */
export function rejectAllSuggestions(): Command {
  return (state, dispatch) => {
    const bySuggestion = collectSuggestions(state);
    if (bySuggestion.size === 0) {
      return false;
    }
    const revisionIds = new Set<number>();
    let removesWholeNode = false;
    for (const entry of bySuggestion.values()) {
      removesWholeNode ||=
        entry.kinds.has("insertBlock") || entry.kinds.has("insertTable");
      for (const id of entry.revisionIds) {
        revisionIds.add(id);
      }
    }
    const rotationCandidates = removesWholeNode
      ? collectAddedBreakRotationCandidates(state.doc)
      : [];

    // Resolve inline/structural revisions first (deletes suggested-inserted
    // text/rows/cells, clears suggested deletions), capturing the transaction
    // so whole-node inserts can be dropped from the same transaction below.
    let tr: Transaction | null = null;
    if (revisionIds.size > 0) {
      const resolved = resolveChangeUnguarded(
        0,
        state.doc.content.size,
        "reject",
        [...revisionIds],
      )(state, (transaction) => {
        tr = transaction;
      });
      if (!resolved) return false;
    }
    const workingTr: Transaction = tr ?? state.tr;

    // Drop any remaining whole-node inserts (empty inserted paragraphs whose
    // inline text the resolve just removed, and whole inserted tables).
    const positions: number[] = [];
    const rejected = {
      mode: "reject",
      resolves: (revisionId: number) => revisionIds.has(revisionId),
    } as const;
    workingTr.doc.descendants((node, pos) => {
      if (rejectionRemovesWholeNode(node, rejected)) {
        positions.push(pos);
      }
      return undefined;
    });
    if (
      !removeSuggestedInsertNodes({
        tr: workingTr,
        positions,
        resolvedRevisionIds: revisionIds,
      })
    ) {
      return false;
    }
    if (
      positions.length > 0 &&
      rotationCandidates.length > 0 &&
      !rotateAddedFinalParagraphBreaks({
        tr: workingTr,
        shouldRotate: mappedRotationCandidatePredicate({
          candidates: rotationCandidates,
          tr: workingTr,
        }),
      })
    ) {
      return false;
    }
    resolveSuggestedCommentMarks(workingTr, () => true, false);

    if (workingTr.steps.length === 0) {
      return false;
    }
    if (dispatch) {
      finishResolution(state, workingTr);
      dispatch(workingTr);
    }
    return true;
  };
}

/**
 * The story with every pending suggestion rejected: what a save writes, what a
 * read of the saved package shows and what rejecting every change starts from.
 *
 * One rejection resolves each suggestion's revisions against a story the
 * others still stand in. Where a paragraph break it takes out joins into, or
 * is displaced onto, another rejected suggestion's paragraph, that suggestion
 * is left holding a break of its own; the rejection runs again until no
 * suggestion is left, so the story never keeps a rejected break.
 */
export const rejectEverySuggestion = (state: EditorState): EditorState =>
  rejectEverySuggestionMapped(state).state;

/** {@link rejectEverySuggestion}, with where each position of `state` ends. */
export const rejectEverySuggestionMapped = (
  state: EditorState,
): { state: EditorState; mapping: Mapping } => {
  const mapping = new Mapping();
  const count = collectSuggestions(state).size;
  let settled = state;
  if (count === 0) return { state: settled, mapping };
  // Every round that changes the story resolves at least one suggestion.
  for (let rounds = count; rounds >= 0; rounds -= 1) {
    const next = rejectSuggestionsOnce(settled);
    if (next === null) break;
    mapping.appendMapping(next.mapping);
    settled = next.state;
  }
  return { state: settled, mapping };
};

/** One `rejectAllSuggestions` over `state`; `null` when nothing was left to reject. */
const rejectSuggestionsOnce = (
  state: EditorState,
): { state: EditorState; mapping: Mapping } | null => {
  const settled = { state, mapping: new Mapping() };
  const rejected = rejectAllSuggestions()(state, (transaction) => {
    settled.state = state.apply(transaction);
    settled.mapping = transaction.mapping;
  });
  return rejected ? settled : null;
};
