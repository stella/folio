/**
 * Comment and Track Changes Commands
 *
 * PM commands for adding/removing comments and accepting/rejecting tracked changes.
 */

import type { Mark, MarkType, Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import { removeRow, TableMap } from "prosemirror-tables";

import {
  appendHeadlessInlineResolution,
  type HeadlessInlineChangeTracking,
} from "../../internal/headlessRevisionResolution";
import { stateAllowsHeadlessRevisionResolution } from "../../internal/headlessRevisionResolutionGuard";
import type { RemovedSectionReference } from "../../internal/sectionEndpointResolution";
import type {
  ParagraphFormatting,
  RunPropertyChange,
  SectionProperties,
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
} from "../../types/document";
import { PARAGRAPH_MARK_CHANGE_KINDS, type ParagraphMarkChangeKind } from "@stll/docx-core/model";

import { expectParagraphAttrs, expectRunPropertyChangeMarkAttrs } from "../attrs";
import {
  addedBreakCarrierBefore,
  finalParagraphsOf,
  paragraphEndsItsContainer,
} from "../containerFinalParagraph";
import { textFormattingToMarks } from "../conversion/toProseDoc";
import {
  markChangedParagraphRanges,
  markStructuralChange,
  markTrackedSectionEndpointRemoval,
} from "../extensions/features/ParagraphChangeTrackerExtension";
import { getDocumentStyleResolver } from "../plugins/documentStyles";
import { holdsNoContent } from "../zeroWidthAnchors";
import { getFolioNodeRevisionCarriers } from "../revisionCarriers";
import { RUN_FORMATTING_MARK_NAMES } from "../runFormattingMarkNames";
import type { ParagraphPropertyChangeAttrs } from "../schema/nodes";
import { getTableCellMergeChange } from "../tableCellMergeRevision";
import {
  hasMatchingCollapsedTableCellMerge,
  resolveCollapsedTableCellMerge,
  resolveVisibleTableCellMerge,
} from "./tableCellMergeResolution";
import {
  hasSerializableParagraphPropertyChange,
  paragraphRejectAttrPatch,
  paragraphRejectOriginalFormatting,
  paragraphPropertiesSnapshot,
  removeParagraphPropertyChanges,
  sectionRejectProperties,
  tableCellRejectAttrPatch,
  tableRejectAttrPatch,
  tableRowRejectAttrPatch,
} from "./propertyChangeScope";

/**
 * Add a comment mark to the current selection.
 */
export function addCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) {
      return false;
    }

    const commentType = state.schema.marks["comment"];
    if (!commentType) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.addMark(from, to, commentType.create({ commentId }));
      dispatch(tr);
    }
    return true;
  };
}

/**
 * Remove a comment mark by ID from the entire document.
 */
export function removeCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const commentType = state.schema.marks["comment"];
    if (!commentType) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr;
      state.doc.descendants((node, pos) => {
        if (node.isText) {
          for (const mark of node.marks) {
            if (mark.type === commentType && mark.attrs["commentId"] === commentId) {
              tr.removeMark(pos, pos + node.nodeSize, mark);
            }
          }
        }
      });
      if (tr.steps.length > 0) {
        dispatch(tr);
      }
    }
    return true;
  };
}

type ResolveMode = "accept" | "reject";

const BULK_INLINE_RESOLUTION_THRESHOLD = 256;

type ResolveExecution = "legacy" | "headless-bulk-inline";

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
function resolveChange(
  from: number,
  to: number,
  mode: "accept" | "reject",
  revisionIds?: readonly number[],
  execution: ResolveExecution = "legacy",
): Command {
  return (state, dispatch) => {
    const insertionType = state.schema.marks["insertion"];
    const deletionType = state.schema.marks["deletion"];
    const styleResolver = getDocumentStyleResolver(state);

    const keepType = mode === "accept" ? insertionType : deletionType;
    const removeType = mode === "accept" ? deletionType : insertionType;
    const revisionSet = revisionIds === undefined ? null : new Set<number>(revisionIds);
    // A range-wide removal lets ProseMirror coalesce one revision split by
    // inline formatting. The id-scoped path must keep matching each mark.
    const removeKeptMarksInBulk = revisionSet === null && keepType !== undefined;
    const canBulkInlineResolution =
      execution === "headless-bulk-inline" &&
      revisionSet === null &&
      from === 0 &&
      to === state.doc.content.size &&
      stateAllowsHeadlessRevisionResolution(state);
    const matchesRevision = (mark: { attrs: Record<string, unknown> }) =>
      revisionSet === null ||
      (typeof mark.attrs["revisionId"] === "number" && revisionSet.has(mark.attrs["revisionId"]));

    if (dispatch) {
      const tr = state.tr;
      const deleteRanges: { from: number; to: number }[] = [];
      const pPrMarkOps: PPrMarkOp[] = [];
      const tableRowStructuralOps: TableRowStructuralOp[] = [];
      const tableCellStructuralOps: TableCellStructuralOp[] = [];
      const deferredRunPropertyChanges: ResolveRunPropertyChangeOptions[] = [];
      let bulkInlineCarrierCount = 0;
      let removedSectionEndpointCount = 0;
      const removedSectionReferences: RemovedSectionReference[] = [];

      state.doc.nodesBetween(from, to, (node, pos): boolean => {
        if (node.type.name === "paragraph") {
          const op = collectPPrMarkOp(node, pos, from, to, mode, revisionSet);
          if (op) {
            pPrMarkOps.push(op);
          }

          const boundaryCovered = rangeCoversParagraphBoundary(from, to, pos, node);
          let nextAttrs: Record<string, unknown> | null = null;

          // Process paragraph property changes (w:pPrChange)
          const propertyChanges = expectParagraphAttrs(node)._propertyChanges;

          if (Array.isArray(propertyChanges) && propertyChanges.length > 0 && boundaryCovered) {
            const matchesPropertyChange = (change: ParagraphPropertyChangeAttrs) =>
              revisionSet === null || revisionSet.has(change.info.id);
            if (propertyChanges.some(matchesPropertyChange)) {
              const rejection =
                mode === "reject"
                  ? removeParagraphPropertyChanges(propertyChanges, matchesPropertyChange)
                  : null;
              let remaining: ParagraphPropertyChangeAttrs[];
              if (rejection === null) {
                remaining = propertyChanges.filter((change) => !matchesPropertyChange(change));
              } else if (rejection.type === "unchanged") {
                remaining = propertyChanges;
              } else {
                remaining = rejection.remaining;
              }
              nextAttrs = {
                ...node.attrs,
                _propertyChanges: remaining.length > 0 ? remaining : null,
              };
              if (rejection?.type === "restore-previous") {
                // Word stores the complete old pPr in the pPrChange, so a
                // reject restores it WHOLESALE within CT_PPrBase scope: a
                // property the change ADDED resets too. Out-of-scope attrs
                // (inline sectPr, paragraph-mark rPr, identity) survive; see
                // propertyChangeScope.ts. Earlier removed runs were folded
                // into the next retained entry, so only a removed trailing
                // run changes the live properties now.
                const inheritedAlignment = expectParagraphAttrs(node).alignmentFromStyle;
                let previousFormattingFromStyle: ParagraphFormatting | undefined;
                if (styleResolver) {
                  previousFormattingFromStyle = styleResolver.resolveParagraphStyle(
                    rejection.previousFormatting?.styleId,
                  ).paragraphFormatting;
                } else if (inheritedAlignment !== undefined) {
                  previousFormattingFromStyle = { alignment: inheritedAlignment };
                }
                Object.assign(
                  nextAttrs,
                  paragraphRejectAttrPatch(
                    rejection.previousFormatting,
                    previousFormattingFromStyle,
                  ),
                );
                nextAttrs["_originalFormatting"] = paragraphRejectOriginalFormatting(
                  rejection.previousFormatting,
                  node.attrs["_originalFormatting"],
                );
              }
            }
          }

          // Process inline section-property changes (w:sectPrChange) carried
          // on the paragraph's `_sectionProperties` attr.
          const sectionProperties = node.attrs["_sectionProperties"] as
            | SectionProperties
            | null
            | undefined;
          const sectionChanges = sectionProperties?.propertyChanges;
          if (
            sectionProperties &&
            Array.isArray(sectionChanges) &&
            sectionChanges.length > 0 &&
            boundaryCovered
          ) {
            const matches = sectionChanges.filter(
              (c) => revisionSet === null || (c.info && revisionSet.has(c.info.id)),
            );
            if (matches.length > 0) {
              const remaining = sectionChanges.filter(
                (c) => revisionSet !== null && (!c.info || !revisionSet.has(c.info.id)),
              );
              let restored: SectionProperties = { ...sectionProperties };
              if (mode === "reject") {
                for (const change of matches.toReversed()) {
                  restored = sectionRejectProperties(restored, change.previousProperties);
                }
              }
              delete restored.propertyChanges;
              if (remaining.length > 0) {
                restored.propertyChanges = remaining;
              }
              nextAttrs = nextAttrs ?? { ...node.attrs };
              nextAttrs["_sectionProperties"] = restored;
              nextAttrs["sectionBreakType"] = sectionBreakTypeFromSectionStart(
                restored.sectionStart,
              );
            }
          }

          if (nextAttrs) {
            tr.setNodeMarkup(pos, undefined, nextAttrs);
          }

          return true;
        }

        if (node.type.name === "tableRow" && rangeCoversNode(from, to, pos, node)) {
          const op = collectTableRowStructuralOp(node, pos, mode, revisionSet);
          if (op) {
            tableRowStructuralOps.push(op);
          }
        }
        if (
          (node.type.name === "tableCell" || node.type.name === "tableHeader") &&
          rangeCoversNode(from, to, pos, node)
        ) {
          tableCellStructuralOps.push(
            ...collectTableCellStructuralOps(node, pos, mode, revisionSet),
          );
        }

        // Table property changes (w:tblPrChange / w:trPrChange / w:tcPrChange)
        // carried on the table / row / cell node attrs.
        const tableChangeAttrName = TABLE_PROPERTY_CHANGE_ATTR_BY_NODE[node.type.name];
        if (tableChangeAttrName !== undefined) {
          if (rangeCoversNode(from, to, pos, node)) {
            const nextAttrs = resolveTablePropertyChangeAttrs(
              node,
              tableChangeAttrName,
              mode,
              revisionSet,
            );
            if (nextAttrs) {
              tr.setNodeMarkup(pos, undefined, nextAttrs);
            }
          }
          return true;
        }
        // Text AND inline atoms (image, shape, hardBreak, tab) can carry
        // tracked-change marks; widen the visitor so rejecting an inserted
        // picture removes it like inserted text. eigenpal #641.
        if (!node.isInline) {
          return true;
        }
        const nodeEnd = pos + node.nodeSize;
        const rangeFrom = Math.max(from, pos);
        const rangeTo = Math.min(to, nodeEnd);

        const runPropertyChangeMark = node.marks.find(
          (mark) => mark.type.name === "runPropertyChange",
        );
        const resolvesRunPropertyChange =
          runPropertyChangeMark !== undefined &&
          expectRunPropertyChangeMarkAttrs(runPropertyChangeMark).changes.length > 0;
        const removesNode =
          removeType !== undefined &&
          node.marks.some((mark) => mark.type === removeType && matchesRevision(mark));
        const removesKeptMark =
          keepType !== undefined &&
          node.marks.some((mark) => mark.type === keepType && matchesRevision(mark));
        if (canBulkInlineResolution) {
          if (resolvesRunPropertyChange) {
            deferredRunPropertyChanges.push({
              tr,
              node,
              from: rangeFrom,
              to: rangeTo,
              mark: runPropertyChangeMark,
              mode,
              revisionSet,
            });
          }
          if (removesNode) {
            deleteRanges.push({ from: rangeFrom, to: rangeTo });
          }
          if (resolvesRunPropertyChange || removesNode || removesKeptMark) {
            bulkInlineCarrierCount++;
          }
          return true;
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
          });
        }

        if (removesNode) {
          deleteRanges.push({ from: rangeFrom, to: rangeTo });
        }

        if (!removeKeptMarksInBulk) {
          for (const mark of node.marks) {
            if (keepType && mark.type === keepType && matchesRevision(mark)) {
              tr.removeMark(rangeFrom, rangeTo, mark);
            }
          }
        }
        return true;
      });

      let bulkInlineChangeTracking: HeadlessInlineChangeTracking | null = null;
      // The linear rewrite pays a fixed whole-document cost. The existing
      // steps are faster for small batches and retain their compact slices.
      const useBulkInlineResolution =
        canBulkInlineResolution && bulkInlineCarrierCount >= BULK_INLINE_RESOLUTION_THRESHOLD;
      if (useBulkInlineResolution) {
        bulkInlineChangeTracking = appendHeadlessInlineResolution(tr, mode, keepType, removeType);
      } else {
        for (const deferred of deferredRunPropertyChanges) {
          resolveRunPropertyChange(deferred);
        }
        if (removeKeptMarksInBulk) {
          tr.removeMark(from, to, keepType);
        }

        let rangesToDelete = deleteRanges;
        if (revisionSet === null) {
          // Adjacent inline ranges have no paragraph boundary between them, so
          // one replacement has the same mapping outside the deleted content.
          const coalescedDeleteRanges: { from: number; to: number }[] = [];
          for (const range of deleteRanges) {
            const previous = coalescedDeleteRanges.at(-1);
            if (previous && range.from <= previous.to) {
              previous.to = Math.max(previous.to, range.to);
              continue;
            }
            coalescedDeleteRanges.push({ from: range.from, to: range.to });
          }
          rangesToDelete = coalescedDeleteRanges;
        }
        for (const range of rangesToDelete.toReversed()) {
          tr.delete(range.from, range.to);
        }
      }

      // Process paragraph-mark ops from end → start so earlier positions stay
      // valid as later paragraphs collapse. Map every position through the
      // accumulated transaction so the inline deletes above don't desync the
      // attr writes or joins below.
      pPrMarkOps.sort((a, b) => b.paragraphPos - a.paragraphPos);
      for (const op of pPrMarkOps) {
        const mappedPos = tr.mapping.map(op.paragraphPos);
        const paragraph = tr.doc.nodeAt(mappedPos);
        if (!paragraph || paragraph.type.name !== "paragraph") {
          continue;
        }
        if (op.action === "clear") {
          tr.setNodeAttribute(mappedPos, "pPrMark", null);
          continue;
        }
        const joinPos = mappedPos + paragraph.nodeSize;
        const nextNode = joinPos < tr.doc.content.size ? tr.doc.nodeAt(joinPos) : null;
        const joinable = nextNode?.type.name === paragraph.type.name;
        if (!joinable) {
          // Nothing to join with: the next sibling is a table, or the paragraph
          // ends its container — a body, a cell, a header, a note or a text box
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
          // Section properties belong to the paragraph mark being resolved.
          // Removing that mark removes its section endpoint; transferring the
          // properties backward would retain the section the revision deleted.
          const resolved = tr.doc.resolve(mappedPos);
          const endsItsContainer = paragraphEndsItsContainer(resolved, paragraph.type.name);
          const canGo = op.markWasAdded || !endsItsContainer;
          if (holdsNoContent(paragraph) && canGo && resolved.parent.childCount > 1) {
            if (ownsSectionEndpoint(paragraph)) {
              removedSectionEndpointCount++;
              removedSectionReferences.push(...sectionReferencesOf(paragraph));
            }
            tr.delete(mappedPos, mappedPos + paragraph.nodeSize);
            continue;
          }
          tr.setNodeAttribute(mappedPos, "pPrMark", null);
          continue;
        }
        // The inline sweep above has already run, so a paragraph that is empty
        // here is one whose whole content was resolved away: a deleted
        // paragraph being accepted, or an inserted one being rejected. Nothing
        // of it survives but the join, and the paragraph the reader is left
        // with is the NEXT one — which keeps its own mark, and in OOXML a
        // paragraph's properties live on its mark. PM's `join` keeps the
        // first node's attrs, so they are restored explicitly; otherwise a
        // deleted heading would hand its style to the paragraph below it.
        const emptyFirstParagraph = holdsNoContent(paragraph);
        // The next paragraph's own `pPrMark` travels with its attrs: it is a
        // different revision, and resolving this one must not resolve it.
        //
        // Section properties live on the paragraph mark. Resolving that mark
        // away removes its section endpoint, so the joined paragraph keeps
        // only a section endpoint already owned by the following paragraph.
        const formattingOwner = emptyFirstParagraph ? nextNode : paragraph;
        const joinedAttrs = {
          ...formattingOwner.attrs,
          pPrMark: nextNode.attrs["pPrMark"],
          sectionBreakType: nextNode.attrs["sectionBreakType"],
          _sectionProperties: nextNode.attrs["_sectionProperties"],
        };
        try {
          tr.join(joinPos);
          tr.setNodeMarkup(mappedPos, undefined, joinedAttrs);
          if (ownsSectionEndpoint(paragraph)) {
            removedSectionEndpointCount++;
            removedSectionReferences.push(...sectionReferencesOf(paragraph));
          }
        } catch {
          // PM rejects the join if the two blocks aren't structurally
          // compatible (e.g. paragraph followed by a table). Leaving the
          // marker is the safe fallback.
        }
      }

      tableRowStructuralOps.sort((left, right) => right.rowPos - left.rowPos);
      let resolvedTableRowStructure = false;
      for (const op of tableRowStructuralOps) {
        const mappedPos = tr.mapping.map(op.rowPos);
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

      tableCellStructuralOps.sort((left, right) => right.cellPos - left.cellPos);
      let resolvedTableCellStructure = false;
      let failedTableCellMergeResolution = false;
      for (const op of tableCellStructuralOps) {
        const mappedPos = tr.mapping.map(op.cellPos);
        const cell = tr.doc.nodeAt(mappedPos);
        if (!cell || (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader")) {
          continue;
        }
        if (op.type === "merge") {
          const resolved =
            op.source === "collapsed"
              ? resolveCollapsedTableCellMerge(tr, mappedPos, op.mode, op.revisionSet)
              : resolveVisibleTableCellMerge(tr, mappedPos, op.mode);
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
        deleteTableCellAt(tr, mappedPos);
        resolvedTableCellStructure = true;
      }
      if (failedTableCellMergeResolution) {
        return false;
      }
      if (resolvedTableCellStructure) {
        markStructuralChange(tr);
      }

      if (bulkInlineChangeTracking) {
        markChangedParagraphRanges(tr, bulkInlineChangeTracking);
      }

      if (removedSectionEndpointCount > 0) {
        markTrackedSectionEndpointRemoval(tr, {
          sourceDoc: state.doc,
          removedEndpointCount: removedSectionEndpointCount,
          removedReferences: removedSectionReferences,
        });
      }

      if (tr.steps.length > 0) {
        dispatch(tr);
      }
    }
    return true;
  };
}

type ResolveRunPropertyChangeOptions = {
  tr: Transaction;
  node: PMNode;
  from: number;
  to: number;
  mark: Mark;
  mode: "accept" | "reject";
  revisionSet: Set<number> | null;
};

const resolveRunPropertyChange = ({
  tr,
  node,
  from,
  to,
  mark,
  mode,
  revisionSet,
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
  if (mode === "accept") {
    return;
  }

  const previousFormatting: RunPropertyChange["previousFormatting"] =
    matches.at(0)?.previousFormatting;
  for (const currentMark of node.marks) {
    if (RUN_FORMATTING_MARK_NAMES.has(currentMark.type.name)) {
      tr.removeMark(from, to, currentMark.type);
    }
  }
  for (const previousMark of textFormattingToMarks(previousFormatting)) {
    tr.addMark(from, to, previousMark);
  }
  if (previousFormatting?.styleId) {
    const characterStyle = node.type.schema.marks["characterStyle"];
    if (characterStyle) {
      tr.addMark(
        from,
        to,
        characterStyle.create({ styleId: previousFormatting.styleId, _styleRPr: null }),
      );
    }
  }
};

type PPrMarkOp = {
  paragraphPos: number;
  action: "clear" | "join";
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

type TableRowRevisionAttr = {
  revisionId: number;
  author?: string;
  date?: string | null;
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

type TableCellRevisionAttr =
  | {
      kind: "ins" | "del";
      info: {
        revisionId: number;
      };
    }
  | {
      kind: "merge";
      info: {
        revisionId: number;
      };
      verticalMerge?: "continue" | "rest";
      verticalMergeOriginal?: "continue" | "rest";
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

function isTableRowRevisionAttr(value: unknown): value is TableRowRevisionAttr {
  return (
    typeof value === "object" &&
    value !== null &&
    "revisionId" in value &&
    typeof value.revisionId === "number"
  );
}

/**
 * Whether an inline revision mark inside a row belongs to that row's own
 * structural revision. Folio writes both halves under one id; Word mints a
 * fresh `w:id` per element, so an author + timestamp match counts too. A third
 * party's edit inside the same row matches neither and survives untouched.
 */
function markBelongsToRowRevision(mark: Mark, revision: TableRowRevisionAttr): boolean {
  if (mark.attrs["revisionId"] === revision.revisionId) {
    return true;
  }
  return (
    revision.author !== undefined &&
    mark.attrs["author"] === revision.author &&
    (mark.attrs["date"] ?? null) === (revision.date ?? null)
  );
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
  const removals: { from: number; to: number; mark: Mark }[] = [];
  row.descendants((node, offset) => {
    if (!node.isInline) {
      return true;
    }
    for (const mark of node.marks) {
      if (mark.type === markType && markBelongsToRowRevision(mark, revision)) {
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
      const keepsCell = (marker.kind === "ins") === (mode === "accept");
      operations.push({
        type: "membership",
        cellPos,
        action: keepsCell ? "clear" : "remove",
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

function isTableCellRevisionAttr(value: unknown): value is TableCellRevisionAttr {
  if (typeof value !== "object" || value === null || !("kind" in value) || !("info" in value)) {
    return false;
  }
  if (value.kind !== "ins" && value.kind !== "del" && value.kind !== "merge") {
    return false;
  }
  const info = value.info;
  return (
    typeof info === "object" &&
    info !== null &&
    "revisionId" in info &&
    typeof info.revisionId === "number"
  );
}

function deleteTableCellAt(tr: Transaction, cellPos: number): void {
  const cell = tr.doc.nodeAt(cellPos);
  if (!cell || (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader")) {
    return;
  }
  const resolved = tr.doc.resolve(cellPos);
  const row = resolved.parent;
  if (row.type.spec["tableRole"] !== "row") {
    return;
  }
  if (row.childCount > 1) {
    tr.delete(cellPos, cellPos + cell.nodeSize);
    return;
  }
  deleteTableRowAt(tr, resolved.start() - 1);
}

function deleteTableRowAt(tr: Transaction, rowPos: number): void {
  const resolved = tr.doc.resolve(rowPos);
  const table = resolved.parent;
  if (table.type.spec["tableRole"] !== "table") {
    return;
  }
  const rowIndex = resolved.index();
  if (table.childCount > 1) {
    const map = TableMap.get(table);
    removeRow(
      tr,
      {
        map,
        table,
        tableStart: resolved.start(),
        left: 0,
        top: rowIndex,
        right: map.width,
        bottom: rowIndex + 1,
      },
      rowIndex,
    );
    return;
  }

  const tablePosition = resolved.start() - 1;
  const tableEnd = tablePosition + table.nodeSize;
  const outerResolved = tr.doc.resolve(tablePosition);
  const parent = outerResolved.parent;
  const tableIndex = outerResolved.index();
  if (parent.canReplace(tableIndex, tableIndex + 1)) {
    tr.delete(tablePosition, tableEnd);
    return;
  }
  const emptyParagraph = tr.doc.type.schema.nodes["paragraph"]?.createAndFill();
  if (
    emptyParagraph &&
    parent.canReplaceWith(tableIndex, tableIndex + 1, emptyParagraph.type, emptyParagraph.marks)
  ) {
    tr.replaceWith(tablePosition, tableEnd, emptyParagraph);
  }
}

export type ParagraphBoundaryChange = {
  from: number;
  to: number;
  type: "insertion" | "deletion";
  author?: string;
  date?: string;
  revisionId?: number;
};

type RevisionInfoAttrs = {
  id?: unknown;
  author?: unknown;
  date?: unknown;
};

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
  if (revisionSet !== null && !revisionSet.has(pPrMark.info.id)) {
    return null;
  }
  // Accepting an added break, or rejecting a removed one, keeps the break
  // (clear the attr). The other two remove it (join with the next paragraph).
  const markWasAdded = paragraphMarkWasAdded(pPrMark.kind);
  const action: PPrMarkOp["action"] = markWasAdded === (mode === "accept") ? "clear" : "join";
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

const SECTION_BREAK_TYPE_VALUES = ["nextPage", "continuous", "oddPage", "evenPage"] as const;

function sectionBreakTypeFromSectionStart(
  sectionStart: SectionProperties["sectionStart"],
): (typeof SECTION_BREAK_TYPE_VALUES)[number] | null {
  const match = SECTION_BREAK_TYPE_VALUES.find((value) => value === sectionStart);
  return match ?? null;
}

/** PM node type name → the attr its tracked property-change records live on. */
const TABLE_PROPERTY_CHANGE_ATTR_BY_NODE: Record<
  string,
  "tblPrChange" | "trPrChange" | "tcPrChange" | undefined
> = {
  table: "tblPrChange",
  tableRow: "trPrChange",
  tableCell: "tcPrChange",
  tableHeader: "tcPrChange",
};

type TablePropertyChangeEntry = {
  info?: { id: number; author: string; date?: string };
  previousFormatting?: TableFormatting | TableRowFormatting | TableCellFormatting;
};

/**
 * Resolve the tracked property-change records on one table / row / cell node.
 * Accept keeps the live formatting and clears the matched records; reject
 * additionally restores the stored previous formatting wholesale (the change
 * element stores the complete old property set — see propertyChangeScope.ts).
 * Returns the next attrs, or `null` when no record matches.
 */
function resolveTablePropertyChangeAttrs(
  node: PMNode,
  attrName: "tblPrChange" | "trPrChange" | "tcPrChange",
  mode: "accept" | "reject",
  revisionSet: Set<number> | null,
): Record<string, unknown> | null {
  const changes = node.attrs[attrName] as TablePropertyChangeEntry[] | null | undefined;
  if (!Array.isArray(changes) || changes.length === 0) {
    return null;
  }
  const matches = changes.filter(
    (c) => revisionSet === null || (c.info && revisionSet.has(c.info.id)),
  );
  if (matches.length === 0) {
    return null;
  }
  const remaining = changes.filter(
    (c) => revisionSet !== null && (!c.info || !revisionSet.has(c.info.id)),
  );
  const nextAttrs: Record<string, unknown> = {
    ...node.attrs,
    [attrName]: remaining.length > 0 ? remaining : null,
  };
  if (mode === "reject") {
    for (const change of matches.toReversed()) {
      if (attrName === "tblPrChange") {
        Object.assign(
          nextAttrs,
          tableRejectAttrPatch(change.previousFormatting as TableFormatting | undefined),
        );
      } else if (attrName === "trPrChange") {
        Object.assign(
          nextAttrs,
          tableRowRejectAttrPatch(change.previousFormatting as TableRowFormatting | undefined),
        );
      } else {
        Object.assign(
          nextAttrs,
          tableCellRejectAttrPatch(
            change.previousFormatting as TableCellFormatting | undefined,
            node.attrs["_originalFormatting"] as TableCellFormatting | null | undefined,
          ),
        );
      }
    }
  }
  return nextAttrs;
}

function isPPrMarkAttr(value: unknown): value is {
  kind: ParagraphMarkChangeKind;
  info: { id: number; author?: unknown; date?: unknown; initials?: unknown };
} {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const kind = (value as { kind?: unknown }).kind;
  const info = (value as { info?: unknown }).info;
  if (!PARAGRAPH_MARK_CHANGE_KINDS.some((allowed) => allowed === kind)) {
    return false;
  }
  if (typeof info !== "object" || info === null) {
    return false;
  }
  return typeof (info as { id?: unknown }).id === "number";
}

/**
 * Whether the mark says the paragraph break was ADDED. A relocation's
 * destination break was added exactly as an insertion's was, and its source
 * break went exactly as a deletion's did; the kinds differ so a reader is told
 * the two ends belong together, not because they resolve differently.
 */
const paragraphMarkWasAdded = (kind: ParagraphMarkChangeKind): boolean =>
  kind === "ins" || kind === "moveTo";

const ownsSectionEndpoint = (paragraph: PMNode): boolean => {
  const attrs = expectParagraphAttrs(paragraph);
  return attrs._sectionProperties !== undefined || attrs.sectionBreakType !== undefined;
};

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

function readRevisionInfo(info: RevisionInfoAttrs | undefined): {
  author?: string;
  date?: string;
  revisionId?: number;
} {
  const revision: { author?: string; date?: string; revisionId?: number } = {};
  if (typeof info?.author === "string") {
    revision.author = info.author;
  }
  if (typeof info?.date === "string") {
    revision.date = info.date;
  }
  if (typeof info?.id === "number") {
    revision.revisionId = info.id;
  }
  return revision;
}

function getListPropertyChangeType(
  attrs: Record<string, unknown>,
  change: ParagraphPropertyChangeAttrs,
): ParagraphBoundaryChange["type"] | null {
  const previousFormatting = change.previousFormatting;
  if (previousFormatting == null || !Object.hasOwn(previousFormatting, "numPr")) {
    return null;
  }

  const currentNumPr = attrs["numPr"];
  const previousNumPr = previousFormatting["numPr"];
  if (previousNumPr == null && currentNumPr != null) {
    return "insertion";
  }
  if (previousNumPr != null && currentNumPr == null) {
    return "deletion";
  }
  if (!areNumPrValuesEqual(previousNumPr, currentNumPr)) {
    return currentNumPr == null ? "deletion" : "insertion";
  }
  return null;
}

function areNumPrValuesEqual(left: unknown, right: unknown): boolean {
  if (left == null || right == null) {
    return left == right;
  }
  if (!isObjectRecord(left) || !isObjectRecord(right)) {
    return Object.is(left, right);
  }
  return left["numId"] === right["numId"] && left["ilvl"] === right["ilvl"];
}

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function toParagraphBoundaryChange(
  node: PMNode,
  pos: number,
  type: ParagraphBoundaryChange["type"],
  info?: RevisionInfoAttrs,
): ParagraphBoundaryChange {
  return {
    from: pos + node.nodeSize - 1,
    to: pos + node.nodeSize,
    type,
    ...readRevisionInfo(info),
  };
}

export function findParagraphBoundaryChangeAtPosition(
  state: EditorState,
  pos: number,
): ParagraphBoundaryChange | null {
  const $pos = state.doc.resolve(pos);
  const node = $pos.parent;
  if (node.type.name !== "paragraph") {
    return null;
  }

  const paragraphPos = $pos.before($pos.depth);
  const pPrMark = node.attrs["pPrMark"];
  if (isPPrMarkAttr(pPrMark)) {
    return toParagraphBoundaryChange(
      node,
      paragraphPos,
      // A relocation's destination break was added exactly as an insertion's
      // was; naming the kinds one by one here read `moveTo` as a deletion and
      // showed the reader a paragraph arriving as one going away.
      paragraphMarkWasAdded(pPrMark.kind) ? "insertion" : "deletion",
      pPrMark.info,
    );
  }

  const propertyChanges = expectParagraphAttrs(node)._propertyChanges;
  if (!Array.isArray(propertyChanges)) {
    return null;
  }

  for (const change of propertyChanges) {
    const type = getListPropertyChangeType(node.attrs, change);
    if (type) {
      return toParagraphBoundaryChange(node, paragraphPos, type, change.info);
    }
  }

  return null;
}

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

/**
 * Accept all tracked changes in the document.
 */
export function acceptAllChanges(): Command {
  return (state, dispatch) => acceptChange(0, state.doc.content.size)(state, dispatch);
}

/**
 * Reject all tracked changes in the document.
 */
export function rejectAllChanges(): Command {
  return (state, dispatch) => rejectChange(0, state.doc.content.size)(state, dispatch);
}

/**
 * Resolve a complete state synchronously for a headless reader or writer.
 *
 * @internal The replacement transaction is consumed here and never exposed:
 * it is deliberately not an editor command and must not be transported or
 * mapped through concurrent edits.
 */
export function resolveAllChangesInHeadlessState(
  state: EditorState,
  mode: ResolveMode,
): EditorState {
  let resolvedState = state;
  resolveChange(
    0,
    state.doc.content.size,
    mode,
    undefined,
    "headless-bulk-inline",
  )(state, (transaction) => {
    resolvedState = state.apply(transaction);
  });
  return resolvedState;
}

/**
 * Find the document range covered by all revision carriers with any of the
 * given ids. Returns null when none are present (already accepted/rejected, or
 * never existed). A replace operation typically passes two ids; standalone
 * text, formatting, paragraph, section, and table revisions pass a single id.
 */
export function findAIEditRevisionRange(
  state: EditorState,
  revisionIds: number | readonly number[],
): { from: number; to: number } | null {
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  const runPropertyChangeType = state.schema.marks["runPropertyChange"];
  const idSet = new Set<number>(typeof revisionIds === "number" ? [revisionIds] : revisionIds);

  const range = { from: null as number | null, to: null as number | null };
  const includeRange = (from: number, to: number): void => {
    if (range.from === null || from < range.from) {
      range.from = from;
    }
    if (range.to === null || to > range.to) {
      range.to = to;
    }
  };

  state.doc.descendants((node, pos) => {
    for (const carrier of getFolioNodeRevisionCarriers(node, pos)) {
      if (idSet.has(carrier.id)) {
        includeRange(carrier.from, carrier.to);
      }
    }
    if (node.type.name === "tableRow") {
      for (const attrName of ["trIns", "trDel"] as const) {
        const marker = node.attrs[attrName];
        if (isTableRowRevisionAttr(marker) && idSet.has(marker.revisionId)) {
          includeRange(pos, pos + node.nodeSize);
          return false;
        }
      }
    }
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const marker = node.attrs["cellMarker"];
      const hasDirectRevision =
        isTableCellRevisionAttr(marker) && idSet.has(marker.info.revisionId);
      const continuationCells = node.attrs["_docxVMergeContinuationCells"];
      const hasCollapsedRevision =
        Array.isArray(continuationCells) &&
        continuationCells.some((cell) => {
          const change = getTableCellMergeChange(cell);
          return change !== null && idSet.has(change.info.id);
        });
      if (hasDirectRevision || hasCollapsedRevision) {
        includeRange(pos, pos + node.nodeSize);
        return false;
      }
    }
    // Widen from `isText` to `isInline` so an AI-edit revision on an inline
    // atom (image, shape) shows up in the matched range. eigenpal #641.
    if (!node.isInline) {
      return;
    }
    for (const mark of node.marks) {
      if (
        mark.type === runPropertyChangeType &&
        expectRunPropertyChangeMarkAttrs(mark).changes.some((change) => idSet.has(change.info.id))
      ) {
        includeRange(pos, pos + node.nodeSize);
        break;
      }
      if (
        (mark.type === insertionType || mark.type === deletionType) &&
        typeof mark.attrs["revisionId"] === "number" &&
        idSet.has(mark.attrs["revisionId"])
      ) {
        includeRange(pos, pos + node.nodeSize);
        break;
      }
    }
    return undefined;
  });

  if (range.from === null || range.to === null) {
    return null;
  }
  return { from: range.from, to: range.to };
}

/**
 * Accept the tracked-change marks belonging to an AI-edit operation.
 * Pass a single revisionId for inserts, standalone deletions, or formatting changes; or the
 * full id list for a replace (one id per side). Returns false when
 * none of the ids match anything in the doc.
 */
export function acceptAIEditRevision(revisionIds: number | readonly number[]): Command {
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
export function rejectAIEditRevision(revisionIds: number | readonly number[]): Command {
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
 * Suggestion (AI-proposed tracked change) commands.
 *
 * A suggestion is a set of `insertion` / `deletion` / `runPropertyChange` marks
 * sharing one `suggestionId` and carrying `provenance: "suggested"`. Accepting
 * rewrites those marks to normal (`"user"`) tracked changes authored by the
 * accepting user; rejecting inverse-applies them like `rejectChange`, reusing
 * the same {@link resolveChange} machinery scoped to the suggestion's revision
 * ids. All three operate as ordinary PM transactions, so they sync via collab
 * and are undoable.
 */

const SUGGESTION_MARK_NAMES = new Set(["insertion", "deletion", "runPropertyChange"]);

const isSuggestionMark = (mark: Mark): boolean =>
  mark.attrs["provenance"] === "suggested" && SUGGESTION_MARK_NAMES.has(mark.type.name);

export type SuggestionKind =
  | "insertion"
  | "deletion"
  | "formatting"
  | "insertBlock"
  | "insertTable"
  | "insertRow"
  | "deleteRow"
  | "insertColumn"
  | "deleteColumn";

/**
 * How a suggestion is applied when accepted:
 * - `"tracked"` — converts to a normal OOXML tracked change (`w:ins`/`w:del`,
 *   paragraph-mark `w:ins`, row/cell `w:ins`/`w:del`);
 * - `"direct"` — no OOXML tracked representation exists (a whole inserted
 *   table), so accepting applies it directly;
 * - `"mixed"` — a heterogeneous group (one `suggestionId` spanning a whole
 *   inserted table AND other edits) whose parts apply both ways. Accept still
 *   resolves the whole group in one transaction; this flag lets the host
 *   message that some parts landed directly.
 */
export type SuggestionAppliedAs = "tracked" | "direct" | "mixed";

export type FolioSuggestion = {
  suggestionId: string;
  /** Contiguous document ranges the suggestion covers, in document order. */
  ranges: readonly { from: number; to: number }[];
  /** Which kinds of change the suggestion contains (deduped, in a stable order). */
  kinds: readonly SuggestionKind[];
  /** How accepting this suggestion applies it (see {@link SuggestionAppliedAs}). */
  appliedAs: SuggestionAppliedAs;
};

const suggestionKindOf = (mark: Mark): SuggestionKind => {
  if (mark.type.name === "insertion") {
    return "insertion";
  }
  if (mark.type.name === "deletion") {
    return "deletion";
  }
  return "formatting";
};

/**
 * A node-attr (block/table) suggestion read off a single node: a whole-node
 * `_suggestedInsert`, or a suggested `trIns`/`trDel`/`cellMarker`.
 */
type StructuralSuggestion = {
  suggestionId: string;
  revisionId: number;
  kind: SuggestionKind;
  /** True for a whole-node insert (paragraph/table): rejected by node deletion. */
  isNodeInsert: boolean;
};

const readSuggestedMarker = (
  marker: unknown,
): { suggestionId: string; revisionId: number } | null => {
  if (
    typeof marker !== "object" ||
    marker === null ||
    (marker as { provenance?: unknown }).provenance !== "suggested"
  ) {
    return null;
  }
  const suggestionId = (marker as { suggestionId?: unknown }).suggestionId;
  const revisionId = (marker as { revisionId?: unknown }).revisionId;
  if (typeof suggestionId !== "string" || typeof revisionId !== "number") {
    return null;
  }
  return { suggestionId, revisionId };
};

const readSuggestedParagraphPropertyChange = (
  change: ParagraphPropertyChangeAttrs,
): { suggestionId: string; revisionId: number } | null => {
  if (change.info.provenance !== "suggested" || typeof change.info.suggestionId !== "string") {
    return null;
  }
  return { suggestionId: change.info.suggestionId, revisionId: change.info.id };
};

const readStructuralSuggestion = (node: PMNode): StructuralSuggestion | null => {
  const attrs = node.attrs;
  // `_suggestedInsert` only carries whole-node semantics for paragraphs
  // (accept → paragraph-mark `w:ins`) and tables (accept → direct). Rows and
  // cells use suggested `trIns`/`trDel`/`cellMarker` instead, so a stray marker
  // on any other node type is ignored rather than mis-classified as a block.
  const name = node.type.name;
  const insertMarker = attrs["_suggestedInsert"];
  if (
    (name === "paragraph" || name === "table") &&
    typeof insertMarker === "object" &&
    insertMarker !== null
  ) {
    const suggestionId = (insertMarker as { suggestionId?: unknown }).suggestionId;
    const revisionId = (insertMarker as { revisionId?: unknown }).revisionId;
    if (typeof suggestionId === "string" && typeof revisionId === "number") {
      return {
        suggestionId,
        revisionId,
        kind: name === "table" ? "insertTable" : "insertBlock",
        isNodeInsert: true,
      };
    }
  }
  if (node.type.name === "tableRow") {
    const ins = readSuggestedMarker(attrs["trIns"]);
    if (ins) {
      return { ...ins, kind: "insertRow", isNodeInsert: false };
    }
    const del = readSuggestedMarker(attrs["trDel"]);
    if (del) {
      return { ...del, kind: "deleteRow", isNodeInsert: false };
    }
  }
  if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
    const cellMarker = attrs["cellMarker"];
    if (typeof cellMarker === "object" && cellMarker !== null) {
      // Only insertion/deletion cell markers participate in suggestions.
      // Merge markers never carry suggestion provenance (cell merge/split is
      // `unsupportedMode`), so any other kind is ignored rather than
      // mis-classified as a column insertion.
      const markerKind = (cellMarker as { kind?: unknown }).kind;
      if (markerKind === "ins" || markerKind === "del") {
        const marker = readSuggestedMarker((cellMarker as { info?: unknown }).info);
        if (marker) {
          const kind: SuggestionKind = markerKind === "del" ? "deleteColumn" : "insertColumn";
          return { ...marker, kind, isNodeInsert: false };
        }
      }
    }
  }
  return null;
};

type SuggestionAccumulator = {
  segments: { from: number; to: number }[];
  kinds: Set<SuggestionKind>;
  revisionIds: Set<number>;
};

const collectSuggestions = (state: EditorState): Map<string, SuggestionAccumulator> => {
  const runPropertyChangeType = state.schema.marks["runPropertyChange"];
  const bySuggestion = new Map<string, SuggestionAccumulator>();
  const entryFor = (suggestionId: string): SuggestionAccumulator => {
    const existing = bySuggestion.get(suggestionId);
    if (existing) {
      return existing;
    }
    const created: SuggestionAccumulator = {
      segments: [],
      kinds: new Set<SuggestionKind>(),
      revisionIds: new Set<number>(),
    };
    bySuggestion.set(suggestionId, created);
    return created;
  };

  state.doc.descendants((node, pos) => {
    const structural = readStructuralSuggestion(node);
    if (structural) {
      const entry = entryFor(structural.suggestionId);
      entry.segments.push({ from: pos, to: pos + node.nodeSize });
      entry.kinds.add(structural.kind);
      entry.revisionIds.add(structural.revisionId);
    }
    if (node.type.name === "paragraph") {
      const propertyChanges = expectParagraphAttrs(node)._propertyChanges;
      if (Array.isArray(propertyChanges)) {
        for (const change of propertyChanges) {
          const suggested = readSuggestedParagraphPropertyChange(change);
          if (!suggested) {
            continue;
          }
          const entry = entryFor(suggested.suggestionId);
          entry.segments.push({ from: pos, to: pos + node.nodeSize });
          entry.kinds.add("formatting");
          entry.revisionIds.add(suggested.revisionId);
        }
      }
    }
    if (!node.isInline) {
      return undefined;
    }
    for (const mark of node.marks) {
      if (!isSuggestionMark(mark)) {
        continue;
      }
      const suggestionId = mark.attrs["suggestionId"];
      if (typeof suggestionId !== "string" || suggestionId.length === 0) {
        continue;
      }
      const entry = entryFor(suggestionId);
      entry.segments.push({ from: pos, to: pos + node.nodeSize });
      entry.kinds.add(suggestionKindOf(mark));
      if (mark.type === runPropertyChangeType) {
        for (const change of expectRunPropertyChangeMarkAttrs(mark).changes) {
          entry.revisionIds.add(change.info.id);
        }
      } else if (typeof mark.attrs["revisionId"] === "number") {
        entry.revisionIds.add(mark.attrs["revisionId"]);
      }
    }
    return undefined;
  });

  return bySuggestion;
};

/**
 * A whole inserted table (`insertTable`) has no OOXML tracked representation and
 * accepts directly; every other kind accepts as a tracked change. A group that
 * mixes a table insert with any other kind (a caller stamping heterogeneous
 * operations with one `suggestionId`) is `"mixed"`.
 */
const suggestionAppliedAs = (kinds: ReadonlySet<SuggestionKind>): SuggestionAppliedAs => {
  if (!kinds.has("insertTable")) {
    return "tracked";
  }
  const hasOtherKind = [...kinds].some((kind) => kind !== "insertTable");
  return hasOtherKind ? "mixed" : "direct";
};

/** Merge sorted, possibly adjacent/overlapping segments into contiguous ranges. */
const mergeSegments = (
  segments: readonly { from: number; to: number }[],
): { from: number; to: number }[] => {
  const sorted = [...segments].toSorted((a, b) => a.from - b.from || a.to - b.to);
  const merged: { from: number; to: number }[] = [];
  for (const segment of sorted) {
    const last = merged.at(-1);
    if (last && segment.from <= last.to) {
      last.to = Math.max(last.to, segment.to);
      continue;
    }
    merged.push({ ...segment });
  }
  return merged;
};

const SUGGESTION_KIND_ORDER: readonly SuggestionKind[] = [
  "insertion",
  "deletion",
  "formatting",
  "insertBlock",
  "insertTable",
  "insertRow",
  "deleteRow",
  "insertColumn",
  "deleteColumn",
];

/**
 * List every suggestion in the document for host consumption: its id, the
 * contiguous ranges it covers, which kinds of change it contains, and how
 * accepting it applies (tracked vs direct).
 */
export function getSuggestions(state: EditorState): FolioSuggestion[] {
  const bySuggestion = collectSuggestions(state);
  const suggestions: FolioSuggestion[] = [];
  for (const [suggestionId, entry] of bySuggestion) {
    suggestions.push({
      suggestionId,
      ranges: mergeSegments(entry.segments),
      kinds: SUGGESTION_KIND_ORDER.filter((kind) => entry.kinds.has(kind)),
      appliedAs: suggestionAppliedAs(entry.kinds),
    });
  }
  // Document order by first range start keeps output stable for the host.
  return suggestions.toSorted((a, b) => (a.ranges[0]?.from ?? 0) - (b.ranges[0]?.from ?? 0));
}

/**
 * The document range covering every mark belonging to `suggestionId`, plus the
 * revision ids those marks carry. Returns null when the suggestion is absent
 * (already accepted/rejected, or never existed).
 */
export function findSuggestionRange(
  state: EditorState,
  suggestionId: string,
): { from: number; to: number; revisionIds: number[] } | null {
  const entry = collectSuggestions(state).get(suggestionId);
  if (!entry || entry.segments.length === 0) {
    return null;
  }
  const ranges = mergeSegments(entry.segments);
  const from = ranges[0]?.from ?? 0;
  const to = ranges.at(-1)?.to ?? from;
  return { from, to, revisionIds: [...entry.revisionIds] };
}

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
  return Number.isNaN(parsed.getTime()) ? new Date().toISOString() : parsed.toISOString();
};

type AddedParagraphMark = {
  kind: "ins" | "moveTo";
  info: {
    id: number;
    author?: unknown;
    date?: unknown;
    initials?: unknown;
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
  if (typeof marker !== "object" || marker === null || !("revisionId" in marker)) {
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
      ...(typeof source.initials === "string" ? { initials: source.initials } : {}),
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
      return { ...attrs, trIns: userInfo(attrs["trIns"] as { revisionId?: unknown }) };
    }
    if (readSuggestedMarker(attrs["trDel"])) {
      return { ...attrs, trDel: userInfo(attrs["trDel"] as { revisionId?: unknown }) };
    }
  }
  if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
    const cellMarker = attrs["cellMarker"] as { kind?: unknown; info?: unknown } | null | undefined;
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
}: ReauthorParagraphPropertySuggestionsOptions): ParagraphPropertyChangeAttrs[] | null => {
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
    info: { id: number; author?: unknown; date?: unknown; initials?: unknown };
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
  const author = typeof mark.info.author === "string" ? mark.info.author : fallbackAuthor;
  if (author === undefined) {
    return false;
  }
  const existing = expectParagraphAttrs(current)._propertyChanges;
  if (hasSerializableParagraphPropertyChange(existing)) {
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
          ...(typeof mark.info.date === "string" ? { date: mark.info.date } : {}),
          ...(typeof mark.info.initials === "string" ? { initials: mark.info.initials } : {}),
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
  const paragraphTypeName = tr.doc.type.schema.nodes["paragraph"]?.name ?? "paragraph";
  const finals = finalParagraphsOf(tr.doc, paragraphTypeName).filter(({ node, position }) =>
    shouldRotate(node, position),
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

type FirstRotatedPropertyChange = { type: "append" } | { type: "rebase"; revisionId: number };

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
        ? Object.assign({}, change, { previousFormatting: paragraphPropertiesSnapshot(previous) })
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

  for (let index = resolved.index() + 1; index < resolved.parent.childCount; index++) {
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
  const carrier = addedBreakCarrierBefore(resolved, paragraph.type.name);
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

/**
 * Place an accepted paragraph's break when a following accepted insertion has
 * already rotated its own break onto that paragraph.
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
    !isPPrMarkAttr(displaced) ||
    !paragraphMarkWasAdded(displaced.kind)
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
const collectAddedBreakRotationCandidates = (doc: PMNode): AddedBreakRotationCandidate[] => {
  const paragraphTypeName = doc.type.schema.nodes["paragraph"]?.name ?? "paragraph";
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
}: MappedRotationCandidatePredicateOptions): ((node: PMNode, position: number) => boolean) => {
  const revisionIdsByPosition = new Map<number, Set<number>>();
  for (const candidate of candidates) {
    const mapped = tr.mapping.mapResult(candidate.position, 1);
    if (mapped.deleted) {
      continue;
    }
    const mappedPosition = mapped.pos;
    const revisionIds = revisionIdsByPosition.get(mappedPosition) ?? new Set<number>();
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
 * Remove suggested nodes without deleting a following accepted insertion's
 * paragraph break that has rotated onto one of them.
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
    const mark = node.attrs["pPrMark"];
    const followingBreak =
      node.type.name === "paragraph" &&
      isAddedPPrMarkAttr(mark) &&
      !resolvedRevisionIds.has(mark.info.id)
        ? mark
        : null;
    tr.delete(position, position + node.nodeSize);
    if (
      followingBreak &&
      !shiftAddedParagraphBreakBefore({
        tr,
        paragraphPosition: position,
        mark: followingBreak,
        firstPropertyChange: { type: "rebase", revisionId: followingBreak.info.id },
      })
    ) {
      return false;
    }
  }
  return true;
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
    if (!canReauthorParagraphPropertySuggestions(state.doc, matchesSuggestion)) {
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
        nextAttrs = convertStructuralSuggestionAttrs(node, options.author, date);
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
        nextAttrs = { ...(nextAttrs ?? node.attrs), _propertyChanges: propertyChanges };
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
        if (
          !isSuggestionMark(mark) ||
          typeof mark.attrs["suggestionId"] !== "string" ||
          !matchesSuggestion(mark.attrs["suggestionId"])
        ) {
          continue;
        }
        changed = true;
        tr.removeMark(from, to, mark);
        if (mark.type === runPropertyChangeType) {
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
            mark.type.create({ changes: nextChanges, provenance: "user", suggestionId: null }),
          );
          continue;
        }
        if (mark.type === insertionType || mark.type === deletionType) {
          tr.addMark(
            from,
            to,
            mark.type.create({
              ...mark.attrs,
              author: options.author,
              date,
              provenance: "user",
              suggestionId: null,
            }),
          );
        }
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
      acceptedParagraphBreaks.filter(({ hadBreak }) => !hadBreak).map(({ position }) => position),
    );
    if (
      directlyMarkedParagraphPositions.size > 0 &&
      !rotateAddedFinalParagraphBreaks({
        tr,
        shouldRotate: (_node, position) => directlyMarkedParagraphPositions.has(position),
        fallbackAuthor: options.author,
      })
    ) {
      return false;
    }

    if (!changed) {
      return false;
    }
    if (dispatch) {
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
export function acceptSuggestion(suggestionId: string, options: AcceptSuggestionOptions): Command {
  return acceptSuggestions((id) => id === suggestionId, options);
}

/** Accept every suggestion in the document. */
export function acceptAllSuggestions(options: AcceptSuggestionOptions): Command {
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
    const entry = collectSuggestions(state).get(suggestionId);
    if (!entry) {
      return false;
    }

    const removesWholeNode = entry.kinds.has("insertBlock") || entry.kinds.has("insertTable");
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
      resolveChange(from, to, "reject", [...entry.revisionIds])(state, (resolved) => {
        tr = resolved;
      });
    }
    const workingTr: Transaction = tr ?? state.tr;

    // Phase 2: drop this suggestion's remaining whole-node inserts.
    const positions: number[] = [];
    workingTr.doc.descendants((node, pos) => {
      const structural = readStructuralSuggestion(node);
      if (structural?.isNodeInsert && structural.suggestionId === suggestionId) {
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

    if (workingTr.steps.length === 0) {
      return false;
    }
    if (dispatch) {
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
      removesWholeNode ||= entry.kinds.has("insertBlock") || entry.kinds.has("insertTable");
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
      resolveChange(0, state.doc.content.size, "reject", [...revisionIds])(state, (resolved) => {
        tr = resolved;
      });
    }
    const workingTr: Transaction = tr ?? state.tr;

    // Drop any remaining whole-node inserts (empty inserted paragraphs whose
    // inline text the resolve just removed, and whole inserted tables).
    const positions: number[] = [];
    workingTr.doc.descendants((node, pos) => {
      if (readStructuralSuggestion(node)?.isNodeInsert) {
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

    if (workingTr.steps.length === 0) {
      return false;
    }
    if (dispatch) {
      dispatch(workingTr);
    }
    return true;
  };
}

type ChangeRange = {
  from: number;
  to: number;
  type: "insertion" | "deletion";
};

/**
 * Find the tracked change mark range at a given cursor position.
 * If the cursor is inside a tracked change, returns the full extent
 * of that mark (expanding to cover all adjacent nodes with the same
 * revision ID). If from !== to (range selection), returns {from, to}.
 */
export function findChangeAtPosition(
  state: EditorState,
  from: number,
  to: number,
): { from: number; to: number } {
  // If there's a range selection, use it directly
  if (from !== to) {
    return { from, to };
  }

  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType && !deletionType) {
    return { from, to };
  }

  // Resolve the position and check marks at cursor
  const $pos = state.doc.resolve(from);
  const node = $pos.parent;
  if (!node.isTextblock) {
    return { from, to };
  }

  // Find the text node at this position and its mark instance. We capture
  // the specific instance (not just the type) so the adjacency expansion
  // below stays inside a single revision — two back-to-back insertions
  // belonging to different `revisionId`s must not be treated as one range.
  let markStart = from;
  let markEnd = from;
  let foundMark: Mark | undefined;

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, offset) => {
    const childStart = $pos.start() + offset;
    const childEnd = childStart + child.nodeSize;
    if (from >= childStart && from <= childEnd && child.isText) {
      for (const mark of child.marks) {
        if (mark.type === insertionType || mark.type === deletionType) {
          foundMark = mark;
          markStart = childStart;
          markEnd = childEnd;
        }
      }
    }
  });

  if (foundMark === undefined) {
    const paragraphChange = findParagraphBoundaryChangeAtPosition(state, from);
    return paragraphChange ? { from: paragraphChange.from, to: paragraphChange.to } : { from, to };
  }

  // Expand to adjacent nodes carrying the *same* mark instance (matching
  // attrs, including revisionId). Two passes — one left-to-right and one
  // right-to-left — so the expansion can cross more than one neighbouring
  // text node on either side (forEach doesn't revisit earlier siblings,
  // which a single-pass walk would need to do to extend leftward by more
  // than one step).
  const sameMark = foundMark;
  const children: {
    childStart: number;
    childEnd: number;
    marks: readonly Mark[];
  }[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, offset) => {
    if (!child.isText) {
      return;
    }
    const childStart = $pos.start() + offset;
    children.push({
      childStart,
      childEnd: childStart + child.nodeSize,
      marks: child.marks,
    });
  });
  let extended = true;
  while (extended) {
    extended = false;
    for (const child of children) {
      if (!child.marks.some((m) => m.eq(sameMark))) {
        continue;
      }
      if (child.childEnd === markStart) {
        markStart = child.childStart;
        extended = true;
      }
      if (child.childStart === markEnd) {
        markEnd = child.childEnd;
        extended = true;
      }
    }
  }

  return { from: markStart, to: markEnd };
}

/**
 * Walk outward from `[fromHint, toHint]` and return the full extent of the
 * tracked-change span carrying `mark` (matched via `Mark.eq`, so attrs like
 * `revisionId` distinguish adjacent changes). Used to keep the
 * navigation/scroll helpers honest when a tracked-change span is split
 * across multiple text nodes due to inline formatting (e.g., a bold word
 * inside an insertion).
 */
function expandTrackedChangeRange(
  state: EditorState,
  mark: Mark,
  fromHint: number,
  toHint: number,
): { from: number; to: number } {
  const carriesSameInlineMark = (node: PMNode | null): node is PMNode =>
    node?.isInline === true && node.marks.some((m) => m.eq(mark));

  // Resolve the boundary positions and hop outward through `nodeBefore`
  // / `nodeAfter` while the neighbouring inline node still carries the
  // same mark instance. O(K) in the number of text nodes that make up
  // the span — `nodesBetween`-based fixed-point expansion is O(K²) and
  // re-walks the same subtree on every iteration.
  let from = fromHint;
  let to = toHint;
  let $from = state.doc.resolve(from);
  let nodeBefore = $from.nodeBefore;
  while (carriesSameInlineMark(nodeBefore)) {
    from -= nodeBefore.nodeSize;
    $from = state.doc.resolve(from);
    nodeBefore = $from.nodeBefore;
  }
  let $to = state.doc.resolve(to);
  let nodeAfter = $to.nodeAfter;
  while (carriesSameInlineMark(nodeAfter)) {
    to += nodeAfter.nodeSize;
    $to = state.doc.resolve(to);
    nodeAfter = $to.nodeAfter;
  }
  return { from, to };
}

/**
 * Find the next tracked change after the given position. Returns the full
 * range of the change (including adjacent text nodes that share the same
 * insertion/deletion mark instance), not just the first text node.
 */
export function findNextChange(state: EditorState, startPos: number): ChangeRange | null {
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType && !deletionType) {
    return null;
  }

  const result = { value: null as ChangeRange | null };

  state.doc.descendants((node, pos) => {
    if (result.value) {
      return false;
    }
    // Widen from `isText` to `isInline` so an image-only insertion / deletion
    // appears in the find-next walk (an atomic image carries the mark itself,
    // not as a text-node sibling). eigenpal #641.
    if (!node.isInline) {
      return;
    }
    if (pos + node.nodeSize <= startPos) {
      return;
    }

    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        // Return the FULL expanded range, even when `startPos` lands
        // inside the matched span. A toolbar that does
        // `findNextChange(state, selectionEnd)` and then accepts the
        // returned range must see the whole revision — clamping `from`
        // up to `startPos` truncates the earlier portion of the same
        // change and leaves orphaned marks behind after accept.
        const expanded = expandTrackedChangeRange(state, mark, pos, pos + node.nodeSize);
        result.value = {
          from: expanded.from,
          to: expanded.to,
          type: mark.type === insertionType ? "insertion" : "deletion",
        };
        return false;
      }
    }
    return undefined;
  });

  // Wrap around (only once)
  if (result.value === null && startPos > 0) {
    return findNextChange(state, 0);
  }

  return result.value;
}

/**
 * Find the previous tracked change before the given position. Returns the
 * full range of the change (including adjacent text nodes that share the
 * same insertion/deletion mark instance), not just the last text node.
 */
export function findPreviousChange(state: EditorState, startPos: number): ChangeRange | null {
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType && !deletionType) {
    return null;
  }

  const result = { value: null as ChangeRange | null };
  // Remember the specific mark instance that produced the kept result, so
  // the walk can skip later text nodes covered by the same expansion
  // without skipping a sibling that carries a *different* tracked-change
  // mark (e.g., an `insertion + deletion` overlay where the same text
  // node belongs to two distinct revisions). Skipping by position alone
  // would miss the nearer overlapping change.
  let resultMark: Mark | null = null;

  state.doc.descendants((node, pos) => {
    // Widen from `isText` to `isInline` so an image-only change appears in
    // the find-previous walk. eigenpal #641.
    if (!node.isInline) {
      return;
    }
    if (pos >= startPos) {
      return false;
    }
    if (
      result.value &&
      resultMark &&
      pos < result.value.to &&
      node.marks.every(
        (m) => (m.type !== insertionType && m.type !== deletionType) || m.eq(resultMark!),
      )
    ) {
      // Already covered by the previous expansion AND no additional
      // tracked-change mark sits on this node — safe to skip.
      return;
    }

    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        const expanded = expandTrackedChangeRange(state, mark, pos, pos + node.nodeSize);
        result.value = {
          from: expanded.from,
          to: expanded.to,
          type: mark.type === insertionType ? "insertion" : "deletion",
        };
        resultMark = mark;
      }
    }
    return undefined;
  });

  // Wrap around (only once — guard prevents infinite recursion)
  if (result.value === null && startPos < state.doc.content.size) {
    return findPreviousChange(state, state.doc.content.size);
  }

  return result.value;
}
