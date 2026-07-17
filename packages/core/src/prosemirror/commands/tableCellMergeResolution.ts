import { Result } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import type { TableCell, TableCellFormatting } from "../../types/document";
import { standaloneTableCellToProseMirror } from "../conversion/toProseDoc";
import { getTableCellMergeChange } from "../tableCellMergeRevision";

export const hasMatchingCollapsedTableCellMerge = (
  node: PMNode,
  revisionSet: Set<number> | null,
): boolean => {
  const continuationCells = node.attrs["_docxVMergeContinuationCells"];
  if (!Array.isArray(continuationCells)) {
    return false;
  }
  return continuationCells.some((cell) => {
    const change = getTableCellMergeChange(cell);
    return change !== null && (revisionSet === null || revisionSet.has(change.info.id));
  });
};

type TableCellRevisionAttr = {
  kind: "merge";
  info: {
    revisionId: number;
  };
  verticalMerge?: "continue" | "rest";
  verticalMergeOriginal?: "continue" | "rest";
};

const isTableCellMergeRevisionAttr = (value: unknown): value is TableCellRevisionAttr => {
  if (
    typeof value !== "object" ||
    value === null ||
    !("kind" in value) ||
    value.kind !== "merge" ||
    !("info" in value)
  ) {
    return false;
  }
  const info = value.info;
  return (
    typeof info === "object" &&
    info !== null &&
    "revisionId" in info &&
    typeof info.revisionId === "number"
  );
};

type TableCellContext = {
  table: PMNode;
  tableStart: number;
  relativeCellPos: number;
};

const tableCellContext = (doc: PMNode, cellPos: number): TableCellContext | null => {
  const resolved = doc.resolve(cellPos);
  for (let depth = resolved.depth; depth > 0; depth -= 1) {
    const node = resolved.node(depth);
    if (node.type.spec["tableRole"] !== "table") {
      continue;
    }
    const tablePos = resolved.before(depth);
    return {
      table: node,
      tableStart: tablePos + 1,
      relativeCellPos: cellPos - tablePos - 1,
    };
  }
  return null;
};

export const resolveVisibleTableCellMerge = (
  tr: Transaction,
  cellPos: number,
  mode: "accept" | "reject",
): boolean => {
  const cell = tr.doc.nodeAt(cellPos);
  const marker = cell?.attrs["cellMarker"];
  if (!cell || !isTableCellMergeRevisionAttr(marker)) {
    return false;
  }
  if (mode === "accept") {
    tr.setNodeAttribute(cellPos, "cellMarker", null);
    return true;
  }

  const originalState = marker.verticalMergeOriginal ?? "rest";
  if (originalState === "continue") {
    return mergeTableCellWithCellAbove(tr, cellPos);
  }

  const originalFormatting = cell.attrs["_originalFormatting"];
  let nextOriginalFormatting = originalFormatting;
  if (typeof originalFormatting === "object" && originalFormatting !== null) {
    nextOriginalFormatting = { ...originalFormatting };
    delete nextOriginalFormatting.vMerge;
  }
  tr.setNodeMarkup(cellPos, undefined, {
    ...cell.attrs,
    cellMarker: null,
    _originalFormatting: nextOriginalFormatting,
  });
  return true;
};

const mergeTableCellWithCellAbove = (tr: Transaction, cellPos: number): boolean => {
  const context = tableCellContext(tr.doc, cellPos);
  const cell = tr.doc.nodeAt(cellPos);
  if (!context || !cell) {
    return false;
  }
  const map = TableMap.get(context.table);
  const rectangle = map.findCell(context.relativeCellPos);
  if (rectangle.top === 0) {
    return false;
  }
  const aboveRelativePos = map.map[(rectangle.top - 1) * map.width + rectangle.left];
  if (aboveRelativePos === undefined || aboveRelativePos === context.relativeCellPos) {
    return false;
  }
  const aboveRectangle = map.findCell(aboveRelativePos);
  if (
    aboveRectangle.left !== rectangle.left ||
    aboveRectangle.right !== rectangle.right ||
    aboveRectangle.bottom !== rectangle.top
  ) {
    return false;
  }
  const abovePos = context.tableStart + aboveRelativePos;
  const aboveCell = tr.doc.nodeAt(abovePos);
  const aboveRowspan = aboveCell?.attrs["rowspan"];
  const cellRowspan = cell.attrs["rowspan"];
  if (
    !aboveCell ||
    typeof aboveRowspan !== "number" ||
    typeof cellRowspan !== "number" ||
    aboveRowspan < 1 ||
    cellRowspan < 1
  ) {
    return false;
  }

  const continuationCells = tableCellContinuationCells(aboveCell, aboveRowspan);
  continuationCells.push(tableCellContinuationFromNode(cell));
  const nestedContinuations = cell.attrs["_docxVMergeContinuationCells"];
  if (Array.isArray(nestedContinuations)) {
    continuationCells.push(...nestedContinuations);
  }

  const cellWasEmpty =
    cell.childCount === 1 &&
    cell.firstChild?.isTextblock === true &&
    cell.firstChild.childCount === 0;
  const aboveCellWasEmpty =
    aboveCell.childCount === 1 &&
    aboveCell.firstChild?.isTextblock === true &&
    aboveCell.firstChild.childCount === 0;
  tr.delete(cellPos, cellPos + cell.nodeSize);
  tr.setNodeMarkup(abovePos, undefined, {
    ...aboveCell.attrs,
    rowspan: aboveRowspan + cellRowspan,
    _docxVMergeContinuationCells: continuationCells,
  });
  if (!cellWasEmpty) {
    if (aboveCellWasEmpty) {
      tr.replaceWith(abovePos + 1, abovePos + 1 + aboveCell.content.size, cell.content);
      return true;
    }
    const contentEnd = abovePos + 1 + aboveCell.content.size;
    tr.insert(contentEnd, cell.content);
  }
  return true;
};

const tableCellContinuationCells = (cell: PMNode, rowspan: number): TableCell[] => {
  const stored = cell.attrs["_docxVMergeContinuationCells"];
  const cells: TableCell[] = Array.isArray(stored) ? [...stored] : [];
  while (cells.length < rowspan - 1) {
    cells.push(emptyVerticalMergeContinuation());
  }
  return cells;
};

const tableCellContinuationFromNode = (cell: PMNode): TableCell => {
  const originalFormatting = cell.attrs["_originalFormatting"];
  const formatting: TableCellFormatting =
    typeof originalFormatting === "object" && originalFormatting !== null
      ? { ...originalFormatting, vMerge: "continue" }
      : { vMerge: "continue" };
  const colspan = cell.attrs["colspan"];
  if (typeof colspan === "number" && colspan > 1) {
    formatting.gridSpan = colspan;
  }
  const continuation: TableCell = {
    type: "tableCell",
    formatting,
    content: [{ type: "paragraph", content: [] }],
  };
  const propertyChanges = cell.attrs["tcPrChange"];
  if (Array.isArray(propertyChanges) && propertyChanges.length > 0) {
    continuation.propertyChanges = [...propertyChanges];
  }
  return continuation;
};

const emptyVerticalMergeContinuation = (): TableCell => ({
  type: "tableCell",
  formatting: { vMerge: "continue" },
  content: [{ type: "paragraph", content: [] }],
});

export const resolveCollapsedTableCellMerge = (
  tr: Transaction,
  cellPos: number,
  mode: "accept" | "reject",
  revisionSet: Set<number> | null,
): boolean => {
  const cell = tr.doc.nodeAt(cellPos);
  const stored = cell?.attrs["_docxVMergeContinuationCells"];
  if (!cell || !Array.isArray(stored)) {
    return false;
  }

  const matchingIndices: number[] = [];
  const nextCells = stored.map((continuationCell, index) => {
    const change = getTableCellMergeChange(continuationCell);
    if (!change || (revisionSet !== null && !revisionSet.has(change.info.id))) {
      return continuationCell;
    }
    matchingIndices.push(index);
    const nextCell = { ...continuationCell };
    delete nextCell.structuralChange;
    return nextCell;
  });
  if (matchingIndices.length === 0) {
    return false;
  }
  if (mode === "accept") {
    tr.setNodeAttribute(cellPos, "_docxVMergeContinuationCells", nextCells);
    return true;
  }

  const splitIndices = matchingIndices.filter((index) => {
    const change = getTableCellMergeChange(stored[index]);
    return (
      change?.verticalMerge === "continue" && (change.verticalMergeOriginal ?? "rest") === "rest"
    );
  });
  if (splitIndices.length === 0) {
    tr.setNodeAttribute(cellPos, "_docxVMergeContinuationCells", nextCells);
    return true;
  }
  const firstSplitIndex = splitIndices.at(0);
  if (
    firstSplitIndex === undefined ||
    splitIndices.length !== stored.length - firstSplitIndex ||
    splitIndices.some((index, offset) => index !== firstSplitIndex + offset)
  ) {
    return false;
  }

  const context = tableCellContext(tr.doc, cellPos);
  if (!context) {
    return false;
  }
  const map = TableMap.get(context.table);
  const rectangle = map.findCell(context.relativeCellPos);
  const rowspan = cell.attrs["rowspan"];
  if (
    typeof rowspan !== "number" ||
    rowspan !== stored.length + 1 ||
    rectangle.bottom - rectangle.top !== rowspan
  ) {
    return false;
  }

  const restorations: { index: number; cell: PMNode }[] = [];
  for (const index of splitIndices) {
    const source = nextCells[index];
    if (!source) {
      return false;
    }
    const restoredCell = createRestoredTableCell(cell, source);
    if (!restoredCell) {
      return false;
    }
    restorations.push({ index, cell: restoredCell });
  }

  const mapFrom = tr.mapping.maps.length;
  for (const restoration of restorations) {
    const row = rectangle.top + restoration.index + 1;
    const insertionPos = context.tableStart + map.positionAt(row, rectangle.left, context.table);
    tr.insert(tr.mapping.slice(mapFrom).map(insertionPos, 1), restoration.cell);
  }

  const remainingCells = nextCells.slice(0, firstSplitIndex);
  tr.setNodeMarkup(cellPos, undefined, {
    ...cell.attrs,
    rowspan: firstSplitIndex + 1,
    _docxVMergeContinuationCells: remainingCells.length > 0 ? remainingCells : null,
  });
  return true;
};

const createRestoredTableCell = (origin: PMNode, source: TableCell): PMNode | null => {
  const formatting = source.formatting ? { ...source.formatting } : undefined;
  if (formatting) {
    delete formatting.vMerge;
  }
  const restoredSource: TableCell = {
    type: "tableCell",
    ...(formatting ? { formatting } : {}),
    ...(source.propertyChanges ? { propertyChanges: source.propertyChanges } : {}),
    content: source.content,
  };
  const restored = standaloneTableCellToProseMirror(
    restoredSource,
    origin.type.name === "tableHeader" ? "tableHeader" : "tableCell",
  );
  return Result.try(() => origin.type.schema.nodeFromJSON(restored.toJSON())).unwrapOr(null);
};
