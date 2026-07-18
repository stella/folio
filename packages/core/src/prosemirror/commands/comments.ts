/**
 * Comment and Track Changes Commands
 *
 * PM commands for adding/removing comments and accepting/rejecting tracked changes.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import { removeRow, TableMap } from "prosemirror-tables";

import type {
  RunPropertyChange,
  SectionProperties,
  TableCellFormatting,
  TableFormatting,
  TableRowFormatting,
} from "../../types/document";
import { expectRunPropertyChangeMarkAttrs } from "../attrs";
import { textFormattingToMarks } from "../conversion/toProseDoc";
import { markStructuralChange } from "../extensions/features/ParagraphChangeTrackerExtension";
import { RUN_FORMATTING_MARK_NAMES } from "../runFormattingMarkNames";
import { getTableCellMergeChange } from "../tableCellMergeRevision";
import {
  hasMatchingCollapsedTableCellMerge,
  resolveCollapsedTableCellMerge,
  resolveVisibleTableCellMerge,
} from "./tableCellMergeResolution";
import {
  paragraphRejectAttrPatch,
  paragraphRejectOriginalFormatting,
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
): Command {
  return (state, dispatch) => {
    const insertionType = state.schema.marks["insertion"];
    const deletionType = state.schema.marks["deletion"];

    const keepType = mode === "accept" ? insertionType : deletionType;
    const removeType = mode === "accept" ? deletionType : insertionType;
    const revisionSet = revisionIds === undefined ? null : new Set<number>(revisionIds);
    const matchesRevision = (mark: { attrs: Record<string, unknown> }) =>
      revisionSet === null ||
      (typeof mark.attrs["revisionId"] === "number" && revisionSet.has(mark.attrs["revisionId"]));

    if (dispatch) {
      const tr = state.tr;
      const deleteRanges: { from: number; to: number }[] = [];
      const pPrMarkOps: PPrMarkOp[] = [];
      const tableRowStructuralOps: TableRowStructuralOp[] = [];
      const tableCellStructuralOps: TableCellStructuralOp[] = [];

      state.doc.nodesBetween(from, to, (node, pos): boolean => {
        if (node.type.name === "paragraph") {
          const op = collectPPrMarkOp(node, pos, from, to, mode, revisionSet);
          if (op) {
            pPrMarkOps.push(op);
          }

          const boundaryCovered = rangeCoversParagraphBoundary(from, to, pos, node);
          let nextAttrs: Record<string, unknown> | null = null;

          // Process paragraph property changes (w:pPrChange)
          const propertyChanges = node.attrs["_propertyChanges"] as
            | {
                info?: { id: number; author: string; date: string };
                previousFormatting?: Record<string, unknown>;
              }[]
            | undefined;

          if (Array.isArray(propertyChanges) && propertyChanges.length > 0 && boundaryCovered) {
            const matches = propertyChanges.filter(
              (c) => revisionSet === null || (c.info && revisionSet.has(c.info.id)),
            );
            if (matches.length > 0) {
              const remaining = propertyChanges.filter(
                (c) => revisionSet !== null && (!c.info || !revisionSet.has(c.info.id)),
              );
              nextAttrs = {
                ...node.attrs,
                _propertyChanges: remaining.length > 0 ? remaining : null,
              };
              if (mode === "reject") {
                // Word stores the complete old pPr in the pPrChange, so a
                // reject restores it WHOLESALE within CT_PPrBase scope: a
                // property the change ADDED resets too. Out-of-scope attrs
                // (inline sectPr, paragraph-mark rPr, identity) survive —
                // see propertyChangeScope.ts. Iterating end → start makes
                // the earliest change's stored pPr win for scoped keys.
                for (const change of matches.toReversed()) {
                  Object.assign(nextAttrs, paragraphRejectAttrPatch(change.previousFormatting));
                  nextAttrs["_originalFormatting"] = paragraphRejectOriginalFormatting(
                    change.previousFormatting,
                    node.attrs["_originalFormatting"],
                  );
                }
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

        if (removeType && node.marks.some((m) => m.type === removeType && matchesRevision(m))) {
          deleteRanges.push({ from: rangeFrom, to: rangeTo });
        }

        for (const mark of node.marks) {
          if (keepType && mark.type === keepType && matchesRevision(mark)) {
            tr.removeMark(rangeFrom, rangeTo, mark);
          }
        }
        return true;
      });

      for (const range of deleteRanges.toReversed()) {
        tr.delete(range.from, range.to);
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
        if (joinPos >= tr.doc.content.size) {
          // No next sibling to join with (paragraph terminates the doc).
          // Leave the marker in place — Word treats this the same way.
          continue;
        }
        try {
          tr.join(joinPos);
          // PM's `join` keeps the first paragraph's attrs, so the marker
          // would survive an otherwise-resolved revision. Drop it now.
          tr.setNodeAttribute(mappedPos, "pPrMark", null);
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
};

type TableRowStructuralOp = {
  rowPos: number;
  attrName: "trIns" | "trDel";
  action: "clear" | "remove";
};

type TableRowRevisionAttr = {
  revisionId: number;
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

type ParagraphPropertyChangeAttrs = {
  info?: RevisionInfoAttrs;
  previousFormatting?: Record<string, unknown> | null;
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
  // accept-ins / reject-del keep the paragraph break (clear attr).
  // reject-ins / accept-del remove the paragraph break (join with next).
  const action: PPrMarkOp["action"] =
    (pPrMark.kind === "ins") === (mode === "accept") ? "clear" : "join";
  return { paragraphPos: pos, action };
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

function isPPrMarkAttr(value: unknown): value is { kind: "ins" | "del"; info: { id: number } } {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const kind = (value as { kind?: unknown }).kind;
  const info = (value as { info?: unknown }).info;
  if (kind !== "ins" && kind !== "del") {
    return false;
  }
  if (typeof info !== "object" || info === null) {
    return false;
  }
  return typeof (info as { id?: unknown }).id === "number";
}

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
      pPrMark.kind === "ins" ? "insertion" : "deletion",
      pPrMark.info,
    );
  }

  const propertyChanges = node.attrs["_propertyChanges"] as
    | ParagraphPropertyChangeAttrs[]
    | undefined;
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
 * Find the document range covered by all inline revision marks
 * carrying any of the given AI-edit `revisionIds`. Returns null when
 * none of those marks are present (already accepted/rejected, or
 * never existed). A replace operation typically passes two ids (one
 * for its deletion side, one for its insertion side); inserts and
 * standalone deletions and formatting changes pass a single id.
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

  state.doc.descendants((node, pos) => {
    if (node.type.name === "tableRow") {
      for (const attrName of ["trIns", "trDel"] as const) {
        const marker = node.attrs[attrName];
        if (isTableRowRevisionAttr(marker) && idSet.has(marker.revisionId)) {
          const start = pos;
          const end = pos + node.nodeSize;
          if (range.from === null || start < range.from) {
            range.from = start;
          }
          if (range.to === null || end > range.to) {
            range.to = end;
          }
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
        const start = pos;
        const end = pos + node.nodeSize;
        if (range.from === null || start < range.from) {
          range.from = start;
        }
        if (range.to === null || end > range.to) {
          range.to = end;
        }
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
        const start = pos;
        const end = pos + node.nodeSize;
        if (range.from === null || start < range.from) {
          range.from = start;
        }
        if (range.to === null || end > range.to) {
          range.to = end;
        }
        break;
      }
      if (
        (mark.type === insertionType || mark.type === deletionType) &&
        typeof mark.attrs["revisionId"] === "number" &&
        idSet.has(mark.attrs["revisionId"])
      ) {
        const start = pos;
        const end = pos + node.nodeSize;
        if (range.from === null || start < range.from) {
          range.from = start;
        }
        if (range.to === null || end > range.to) {
          range.to = end;
        }
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
    const marker = insertMarker as { revisionId?: unknown; initials?: unknown };
    if (name === "table") {
      // Whole inserted table has no OOXML tracked form → accept applies it
      // directly by clearing the suggestion marker (the table stays as content).
      return { ...attrs, _suggestedInsert: null };
    }
    // Inserted paragraph → real inserted-paragraph tracked change: mark the
    // paragraph break as `w:ins` (the inline runs are re-authored by the mark
    // pass). Keeps the revision id so the halves stay paired.
    if (typeof marker.revisionId === "number") {
      return {
        ...attrs,
        _suggestedInsert: null,
        pPrMark: {
          kind: "ins",
          info: {
            id: marker.revisionId,
            author,
            date,
            ...(typeof marker.initials === "string" ? { initials: marker.initials } : {}),
          },
        },
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
    const insertionType = state.schema.marks["insertion"];
    const deletionType = state.schema.marks["deletion"];
    const runPropertyChangeType = state.schema.marks["runPropertyChange"];
    const date = normalizeAcceptDate(options.date);
    const tr = state.tr;
    let changed = false;

    // Mark steps AND setNodeAttribute do not shift positions, so the positions
    // read from `state.doc` stay valid across the accumulated steps.
    state.doc.descendants((node, pos) => {
      const structural = readStructuralSuggestion(node);
      if (structural && matchesSuggestion(structural.suggestionId)) {
        const nextAttrs = convertStructuralSuggestionAttrs(node, options.author, date);
        if (nextAttrs) {
          tr.setNodeMarkup(pos, undefined, nextAttrs);
          changed = true;
        }
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
    for (const pos of positions.toSorted((a, b) => b - a)) {
      const node = workingTr.doc.nodeAt(pos);
      if (node) {
        workingTr.delete(pos, pos + node.nodeSize);
      }
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
    for (const entry of bySuggestion.values()) {
      for (const id of entry.revisionIds) {
        revisionIds.add(id);
      }
    }

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
    for (const pos of positions.toSorted((a, b) => b - a)) {
      const node = workingTr.doc.nodeAt(pos);
      if (node) {
        workingTr.delete(pos, pos + node.nodeSize);
      }
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
