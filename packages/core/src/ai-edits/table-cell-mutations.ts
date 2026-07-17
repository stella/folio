import { Result } from "better-result";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { expectTableCellAttrs } from "../prosemirror/attrs";
import { standaloneTableCellFromProseMirror } from "../prosemirror/conversion/fromProseDoc";
import { standaloneTableCellToProseMirror } from "../prosemirror/conversion/toProseDoc";
import { markStructuralChange } from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import type { TableCell, TableCellFormatting } from "../types/document";
import { tableRectanglesEqual, type TableRectangle } from "./table-mutation-plan";
import { tableRectangleCutsMergedCell } from "./table-targets";

type TableCellRectangleMutationOptions = {
  tr: Transaction;
  tablePosition: number;
  table: PMNode;
  rectangle: TableRectangle;
};

type TrackedTableCellRectangleMutationOptions = TableCellRectangleMutationOptions & {
  revisionId: number;
  author: string;
  date: string;
};

export const mergeTableRectangle = ({
  tr,
  tablePosition,
  table,
  rectangle,
}: TableCellRectangleMutationOptions): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height ||
    rectangle.left >= rectangle.right ||
    rectangle.top >= rectangle.bottom ||
    (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) ||
    tableRectangleCutsMergedCell(map, rectangle)
  ) {
    return null;
  }

  const tableStart = tablePosition + 1;
  const seen = new Set<number>();
  const cells: { position: number; cell: PMNode }[] = [];
  let appendedContent = Fragment.empty;
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    for (let column = rectangle.left; column < rectangle.right; column++) {
      const cellPosition = map.map[row * map.width + column];
      if (cellPosition === undefined || seen.has(cellPosition)) {
        continue;
      }
      const cell = table.nodeAt(cellPosition);
      if (!cell) {
        return null;
      }
      seen.add(cellPosition);
      cells.push({ position: cellPosition, cell });
      if (cells.length > 1 && !isEmptyTableCell(cell)) {
        appendedContent = appendedContent.append(cell.content);
      }
    }
  }
  const merged = cells.at(0);
  if (!merged || cells.length < 2) {
    return null;
  }

  const colspan: unknown = merged.cell.attrs["colspan"];
  const rowspan: unknown = merged.cell.attrs["rowspan"];
  const colwidth: unknown = merged.cell.attrs["colwidth"];
  if (
    typeof colspan !== "number" ||
    !Number.isInteger(colspan) ||
    colspan < 1 ||
    typeof rowspan !== "number" ||
    !Number.isInteger(rowspan) ||
    rowspan < 1 ||
    (colwidth !== null &&
      colwidth !== undefined &&
      (!Array.isArray(colwidth) ||
        colwidth.length !== colspan ||
        !colwidth.every((width) => typeof width === "number")))
  ) {
    return null;
  }
  const mergedColspan = rectangle.right - rectangle.left;
  const nextColwidth = Array.isArray(colwidth) ? [...colwidth] : null;
  while (nextColwidth && nextColwidth.length < mergedColspan) {
    nextColwidth.push(0);
  }
  const mapFrom = tr.mapping.maps.length;
  for (const { position: cellPosition, cell } of cells.slice(1)) {
    const position = tr.mapping.slice(mapFrom).map(tableStart + cellPosition);
    tr = tr.delete(position, position + cell.nodeSize);
  }
  const absoluteMergedPosition = tableStart + merged.position;
  tr = tr.setNodeMarkup(absoluteMergedPosition, undefined, {
    ...merged.cell.attrs,
    colspan: mergedColspan,
    rowspan: rectangle.bottom - rectangle.top,
    colwidth: nextColwidth,
  });
  if (appendedContent.size > 0) {
    const contentEnd = absoluteMergedPosition + 1 + merged.cell.content.size;
    const contentStart = isEmptyTableCell(merged.cell) ? absoluteMergedPosition + 1 : contentEnd;
    tr = tr.replaceWith(contentStart, contentEnd, appendedContent);
  }
  return tr;
};

export const mergeTrackedVerticalTableCells = ({
  tr,
  tablePosition,
  table,
  rectangle,
  revisionId,
  author,
  date,
}: TrackedTableCellRectangleMutationOptions): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.right - rectangle.left !== 1 ||
    rectangle.bottom - rectangle.top < 2 ||
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height ||
    tableRectangleCutsMergedCell(map, rectangle)
  ) {
    return null;
  }

  const cells: { position: number; cell: PMNode }[] = [];
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    const position = map.map[row * map.width + rectangle.left];
    if (position === undefined) {
      return null;
    }
    const cell = table.nodeAt(position);
    if (
      !cell ||
      !tableRectanglesEqual(map.findCell(position), {
        left: rectangle.left,
        top: row,
        right: rectangle.right,
        bottom: row + 1,
      })
    ) {
      return null;
    }
    const attrs = expectTableCellAttrs(cell);
    if (
      attrs.colspan !== 1 ||
      attrs.rowspan !== 1 ||
      attrs.cellMarker !== undefined ||
      attrs._docxVMergeContinuationCells !== undefined ||
      attrs._preserveVMergeRestart === true ||
      attrs._originalFormatting?.vMerge !== undefined
    ) {
      return null;
    }
    cells.push({ position, cell });
  }

  const origin = cells.at(0);
  if (
    !origin ||
    cells.length !== rectangle.bottom - rectangle.top ||
    cells.some(({ cell }) => cell.type !== origin.cell.type) ||
    cells.slice(1).some(({ cell }) => !isEmptyTableCell(cell))
  ) {
    return null;
  }

  const continuationCells: TableCell[] = [];
  for (const { cell } of cells.slice(1)) {
    const continuation = standaloneTableCellFromProseMirror(cell);
    continuation.formatting = {
      ...continuation.formatting,
      vMerge: "continue",
    };
    continuation.structuralChange = {
      type: "tableCellMerge",
      info: { id: revisionId, author, date },
      verticalMerge: "continue",
      verticalMergeOriginal: "rest",
    };
    continuationCells.push(continuation);
  }

  const nextTr = mergeTableRectangle({ tr, tablePosition, table, rectangle });
  if (!nextTr) {
    return null;
  }
  nextTr.setNodeAttribute(
    tablePosition + 1 + origin.position,
    "_docxVMergeContinuationCells",
    continuationCells,
  );
  markStructuralChange(nextTr);
  return nextTr;
};

export const splitTableRectangle = ({
  tr,
  tablePosition,
  table,
  rectangle,
}: TableCellRectangleMutationOptions): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height ||
    rectangle.left >= rectangle.right ||
    rectangle.top >= rectangle.bottom ||
    (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1)
  ) {
    return null;
  }
  const cellPosition = map.map[rectangle.top * map.width + rectangle.left];
  if (cellPosition === undefined) {
    return null;
  }
  const cellRectangle = map.findCell(cellPosition);
  if (!tableRectanglesEqual(cellRectangle, rectangle)) {
    return null;
  }
  const cell = table.nodeAt(cellPosition);
  if (!cell) {
    return null;
  }

  const colspan: unknown = cell.attrs["colspan"];
  const rowspan: unknown = cell.attrs["rowspan"];
  const colwidth: unknown = cell.attrs["colwidth"];
  if (
    typeof colspan !== "number" ||
    !Number.isInteger(colspan) ||
    colspan !== rectangle.right - rectangle.left ||
    typeof rowspan !== "number" ||
    !Number.isInteger(rowspan) ||
    rowspan !== rectangle.bottom - rectangle.top ||
    (colwidth !== null &&
      colwidth !== undefined &&
      (!Array.isArray(colwidth) ||
        colwidth.length !== colspan ||
        !colwidth.every((width) => typeof width === "number")))
  ) {
    return null;
  }

  const baseAttrs = {
    ...cell.attrs,
    colspan: 1,
    rowspan: 1,
    _preserveVMergeRestart: null,
    _docxVMergeContinuationCells: null,
  };
  const attrsByColumn = Array.from({ length: colspan }, (_, index) => ({
    ...baseAttrs,
    colwidth:
      Array.isArray(colwidth) && typeof colwidth[index] === "number" && colwidth[index] > 0
        ? [colwidth[index]]
        : null,
  }));
  const tableStart = tablePosition + 1;
  const insertions: { position: number; cell: PMNode }[] = [];
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    let position = map.positionAt(row, rectangle.left, table);
    if (row === rectangle.top) {
      position += cell.nodeSize;
    }
    for (let column = rectangle.left; column < rectangle.right; column++) {
      if (row === rectangle.top && column === rectangle.left) {
        continue;
      }
      const attrs = attrsByColumn[column - rectangle.left];
      const nextCell = attrs ? cell.type.createAndFill(attrs) : null;
      if (!nextCell) {
        return null;
      }
      insertions.push({ position: tableStart + position, cell: nextCell });
    }
  }

  const mapFrom = tr.mapping.maps.length;
  for (const insertion of insertions) {
    tr = tr.insert(tr.mapping.slice(mapFrom).map(insertion.position, 1), insertion.cell);
  }
  return tr.setNodeMarkup(tableStart + cellPosition, undefined, attrsByColumn[0]);
};

export const splitTrackedVerticalTableCell = ({
  tr,
  tablePosition,
  table,
  rectangle,
  revisionId,
  author,
  date,
}: TrackedTableCellRectangleMutationOptions): Transaction | null => {
  const map = TableMap.get(table);
  if (
    rectangle.right - rectangle.left !== 1 ||
    rectangle.bottom - rectangle.top < 2 ||
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height
  ) {
    return null;
  }
  const relativeCellPosition = map.map[rectangle.top * map.width + rectangle.left];
  if (relativeCellPosition === undefined) {
    return null;
  }
  const cellRectangle = map.findCell(relativeCellPosition);
  if (!tableRectanglesEqual(cellRectangle, rectangle)) {
    return null;
  }
  const cell = table.nodeAt(relativeCellPosition);
  if (!cell) {
    return null;
  }
  const attrs = expectTableCellAttrs(cell);
  if (
    attrs.colspan !== 1 ||
    attrs.rowspan !== rectangle.bottom - rectangle.top ||
    attrs.cellMarker !== undefined
  ) {
    return null;
  }
  const continuationCells = attrs._docxVMergeContinuationCells;
  if (
    continuationCells !== undefined &&
    continuationCells.length !== rectangle.bottom - rectangle.top - 1
  ) {
    return null;
  }

  const marker = {
    kind: "merge" as const,
    info: { revisionId, author, date },
    verticalMergeOriginal: "continue" as const,
  };
  const insertedCells: PMNode[] = [];
  for (let index = 0; index < rectangle.bottom - rectangle.top - 1; index++) {
    const source = continuationCells?.[index];
    if (source?.structuralChange !== undefined) {
      return null;
    }
    const insertedCell = source
      ? trackedSplitCellFromStoredSource({ origin: cell, source, marker })
      : cell.type.createAndFill({
          ...cell.attrs,
          rowspan: 1,
          cellMarker: marker,
          _originalFormatting: formattingWithoutVerticalMerge(attrs._originalFormatting),
          _preserveVMergeRestart: null,
          _docxVMergeContinuationCells: null,
        });
    if (!insertedCell) {
      return null;
    }
    insertedCells.push(insertedCell);
  }

  const tableStart = tablePosition + 1;
  const mapFrom = tr.mapping.maps.length;
  for (const [index, insertedCell] of insertedCells.entries()) {
    const row = rectangle.top + index + 1;
    const position = tableStart + map.positionAt(row, rectangle.left, table);
    tr = tr.insert(tr.mapping.slice(mapFrom).map(position, 1), insertedCell);
  }
  tr = tr.setNodeMarkup(tableStart + relativeCellPosition, undefined, {
    ...cell.attrs,
    rowspan: 1,
    _originalFormatting: formattingWithoutVerticalMerge(attrs._originalFormatting),
    _preserveVMergeRestart: null,
    _docxVMergeContinuationCells: null,
  });
  markStructuralChange(tr);
  return tr;
};

const isEmptyTableCell = (cell: PMNode): boolean =>
  cell.childCount === 1 &&
  cell.firstChild?.isTextblock === true &&
  cell.firstChild.childCount === 0;

type TrackedSplitCellMarker = {
  kind: "merge";
  info: {
    revisionId: number;
    author: string;
    date: string;
  };
  verticalMergeOriginal: "continue";
};

type TrackedSplitCellFromStoredSourceOptions = {
  origin: PMNode;
  source: TableCell;
  marker: TrackedSplitCellMarker;
};

const trackedSplitCellFromStoredSource = ({
  origin,
  source,
  marker,
}: TrackedSplitCellFromStoredSourceOptions): PMNode | null => {
  const formatting = formattingWithoutVerticalMerge(source.formatting);
  const restoredSource: TableCell = {
    type: "tableCell",
    ...(formatting ? { formatting } : {}),
    ...(source.propertyChanges ? { propertyChanges: source.propertyChanges } : {}),
    content: source.content,
  };
  const converted = standaloneTableCellToProseMirror(
    restoredSource,
    origin.type.name === "tableHeader" ? "tableHeader" : "tableCell",
  );
  const compatible = Result.try(() => origin.type.schema.nodeFromJSON(converted.toJSON())).unwrapOr(
    null,
  );
  if (!compatible) {
    return null;
  }
  return compatible.type.create(
    {
      ...compatible.attrs,
      rowspan: 1,
      cellMarker: marker,
      _preserveVMergeRestart: null,
      _docxVMergeContinuationCells: null,
    },
    compatible.content,
  );
};

const formattingWithoutVerticalMerge = (
  formatting: TableCellFormatting | undefined,
): TableCellFormatting | undefined => {
  if (!formatting) {
    return undefined;
  }
  const next = { ...formatting };
  delete next.vMerge;
  return next;
};
