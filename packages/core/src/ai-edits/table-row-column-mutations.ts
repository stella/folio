import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import {
  columnIsHeader,
  removeColumn,
  removeRow,
  rowIsHeader,
  TableMap,
  tableNodeTypes,
} from "prosemirror-tables";

import { markStructuralChange } from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import type { TrackedChangeProvenance } from "../prosemirror/schema/marks";
import { stripBlockIdentityAttrs } from "./block-identity";
import {
  findEnclosingTableCell,
  findEnclosingTableRow,
  type TableRowTarget,
} from "./table-targets";

export type TableRowInsertion = {
  tableStart: number;
  rowPosition: number;
  rowType: PMNode["type"];
  cells: readonly PMNode[];
  rowspanUpdates: readonly number[];
};

export type TableRowDeletion = Omit<TableRowTarget, "table">;

export type TableColumnInsertion = {
  tablePosition: number;
  columnIndex: number;
};

export type TableColumnDeletion = TableColumnInsertion;

export type TableStructureRevision = {
  revisionId: number;
  author: string;
  date: string;
  /** Optional author initials (w:initials), carried for round-trip. */
  initials?: string;
  /**
   * `"suggested"` marks the produced `trIns`/`trDel`/`cellMarker` as an AI
   * proposal that is stripped from serialized DOCX until accepted.
   */
  provenance?: TrackedChangeProvenance;
  suggestionId?: string;
};

type TableRowColumnMutationResult =
  | { type: "applied"; transaction: Transaction; revisionId: number | null }
  | { type: "unsupported" };

export const getTableColumnCoordinateKey = ({
  tablePosition,
  columnIndex,
}: TableColumnInsertion): string => `${tablePosition}:${columnIndex}`;

type FindTableColumnInsertionOptions = {
  doc: PMNode;
  blockFrom: number;
  position: "after" | "before";
};

export const findTableColumnInsertion = ({
  doc,
  blockFrom,
  position,
}: FindTableColumnInsertionOptions):
  | (TableColumnInsertion & {
      boundaryPosition: number;
    })
  | null => {
  const target = findEnclosingTableCell(doc, blockFrom);
  if (!target) {
    return null;
  }
  return {
    tablePosition: target.tablePosition,
    columnIndex: position === "before" ? target.leftColumnIndex : target.rightColumnIndex,
    boundaryPosition: position === "before" ? target.cellPosition : target.cellEndPosition,
  };
};

type FindTableRowInsertionOptions = FindTableColumnInsertionOptions;

export const findTableRowInsertion = ({
  doc,
  blockFrom,
  position,
}: FindTableRowInsertionOptions): TableRowInsertion | null => {
  const target = findEnclosingTableRow(doc, blockFrom);
  if (!target) {
    return null;
  }
  const map = TableMap.get(target.table);
  const rowIndex = position === "before" ? target.rowIndex : target.rowIndex + 1;
  return buildTableRowInsertion({
    map,
    table: target.table,
    tableStart: target.tableStart,
    rowIndex,
  });
};

type ApplyTableRowInsertionOptions = {
  tr: Transaction;
  insertion: TableRowInsertion;
  cellTexts: readonly string[] | undefined;
  revision: TableStructureRevision | null;
};

export const applyTableRowInsertion = ({
  tr,
  insertion,
  cellTexts,
  revision,
}: ApplyTableRowInsertionOptions): TableRowColumnMutationResult => {
  if (revision && insertion.rowspanUpdates.length > 0) {
    return { type: "unsupported" };
  }

  const liveRowspanUpdates: { position: number; cell: PMNode; rowspan: number }[] = [];
  for (const updatePosition of insertion.rowspanUpdates) {
    const cellPosition = insertion.tableStart + updatePosition;
    const liveCell = tr.doc.nodeAt(cellPosition);
    const rowspan: unknown = liveCell?.attrs["rowspan"];
    if (!liveCell || typeof rowspan !== "number") {
      return { type: "unsupported" };
    }
    liveRowspanUpdates.push({ position: cellPosition, cell: liveCell, rowspan });
  }

  for (const update of liveRowspanUpdates) {
    tr.setNodeMarkup(update.position, undefined, {
      ...update.cell.attrs,
      rowspan: update.rowspan + 1,
    });
  }

  const rowAttrs = revision
    ? {
        trIns: revision,
      }
    : null;
  const row = insertion.rowType.create(rowAttrs, insertion.cells);
  tr.insert(insertion.rowPosition, populateTableRow(row, cellTexts));
  return applied(tr, revision);
};

type ApplyTableColumnInsertionOptions = {
  tr: Transaction;
  insertion: TableColumnInsertion;
  cellTexts: readonly string[] | undefined;
  revision: TableStructureRevision | null;
};

export const applyTableColumnInsertion = ({
  tr,
  insertion,
  cellTexts,
  revision,
}: ApplyTableColumnInsertionOptions): TableRowColumnMutationResult => {
  const tablePosition = tr.mapping.map(insertion.tablePosition, 1);
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") {
    return { type: "unsupported" };
  }

  const actions = buildTableColumnInsertionActions({
    map: TableMap.get(table),
    table,
    columnIndex: insertion.columnIndex,
    cellTexts,
  });
  if (!actions || (revision && actions.some(({ type }) => type === "setColspan"))) {
    return { type: "unsupported" };
  }

  const mapFrom = tr.mapping.maps.length;
  for (const action of actions) {
    const position = tr.mapping.slice(mapFrom).map(tablePosition + 1 + action.position);
    if (action.type === "setColspan") {
      tr.setNodeMarkup(position, undefined, action.attrs);
      continue;
    }

    const cell = revision
      ? action.cell.type.create(
          {
            ...action.cell.attrs,
            cellMarker: {
              kind: "ins",
              info: revision,
            },
          },
          action.cell.content,
        )
      : action.cell;
    tr.insert(position, cell);
  }
  return applied(tr, revision);
};

type ApplyTableColumnDeletionOptions = {
  tr: Transaction;
  deletion: TableColumnDeletion;
  insertedColumnCount: number;
  revision: TableStructureRevision | null;
};

export const applyTableColumnDeletion = ({
  tr,
  deletion,
  insertedColumnCount,
  revision,
}: ApplyTableColumnDeletionOptions): TableRowColumnMutationResult => {
  const tablePosition = tr.mapping.map(deletion.tablePosition, 1);
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") {
    return { type: "unsupported" };
  }

  const map = TableMap.get(table);
  const columnIndex = deletion.columnIndex + insertedColumnCount;
  if (columnIndex < 0 || columnIndex >= map.width) {
    return { type: "unsupported" };
  }

  if (revision) {
    const cellPositions = getTableColumnDeletionCellPositions({ map, table, columnIndex });
    if (!cellPositions) {
      return { type: "unsupported" };
    }
    for (const cellPosition of cellPositions) {
      tr.setNodeAttribute(tablePosition + 1 + cellPosition, "cellMarker", {
        kind: "del",
        info: revision,
      });
    }
    markStructuralChange(tr);
    return applied(tr, revision);
  }

  if (map.width === 1) {
    const transaction = deleteTableNode({ tr, tablePosition, table });
    return transaction ? applied(transaction, null) : { type: "unsupported" };
  }
  removeColumn(
    tr,
    {
      map,
      table,
      tableStart: tablePosition + 1,
      left: columnIndex,
      top: 0,
      right: columnIndex + 1,
      bottom: map.height,
    },
    columnIndex,
  );
  return applied(tr, null);
};

type ApplyTableRowDeletionOptions = {
  tr: Transaction;
  deletion: TableRowDeletion;
  revision: TableStructureRevision | null;
};

export const applyTableRowDeletion = ({
  tr,
  deletion,
  revision,
}: ApplyTableRowDeletionOptions): TableRowColumnMutationResult => {
  if (revision) {
    const rowPosition = tr.mapping.map(deletion.rowPosition);
    const row = tr.doc.nodeAt(rowPosition);
    if (
      !row ||
      row.type.spec["tableRole"] !== "row" ||
      row.attrs["trIns"] != null ||
      row.attrs["trDel"] != null
    ) {
      return { type: "unsupported" };
    }
    tr.setNodeAttribute(rowPosition, "trDel", revision);
    markStructuralChange(tr);
    return applied(tr, revision);
  }

  const table = tr.doc.nodeAt(deletion.tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table" || deletion.rowIndex >= table.childCount) {
    return { type: "unsupported" };
  }
  if (table.childCount === 1) {
    const transaction = deleteTableNode({ tr, tablePosition: deletion.tablePosition, table });
    return transaction ? applied(transaction, null) : { type: "unsupported" };
  }

  const map = TableMap.get(table);
  removeRow(
    tr,
    {
      map,
      table,
      tableStart: deletion.tableStart,
      left: 0,
      top: deletion.rowIndex,
      right: map.width,
      bottom: deletion.rowIndex + 1,
    },
    deletion.rowIndex,
  );
  return applied(tr, null);
};

const applied = (
  transaction: Transaction,
  revision: TableStructureRevision | null,
): TableRowColumnMutationResult => ({
  type: "applied",
  transaction,
  revisionId: revision?.revisionId ?? null,
});

type DeleteTableNodeOptions = {
  tr: Transaction;
  tablePosition: number;
  table: PMNode;
};

const deleteTableNode = ({
  tr,
  tablePosition,
  table,
}: DeleteTableNodeOptions): Transaction | null => {
  const tableEnd = tablePosition + table.nodeSize;
  const resolved = tr.doc.resolve(tablePosition);
  const parent = resolved.parent;
  const tableIndex = resolved.index();
  if (parent.canReplace(tableIndex, tableIndex + 1)) {
    return tr.delete(tablePosition, tableEnd);
  }
  const emptyParagraph = tr.doc.type.schema.nodes["paragraph"]?.createAndFill();
  if (
    !emptyParagraph ||
    !parent.canReplaceWith(tableIndex, tableIndex + 1, emptyParagraph.type, emptyParagraph.marks)
  ) {
    return null;
  }
  return tr.replaceWith(tablePosition, tableEnd, emptyParagraph);
};

type BuildTableRowInsertionOptions = {
  map: TableMap;
  table: PMNode;
  tableStart: number;
  rowIndex: number;
};

const buildTableRowInsertion = ({
  map,
  table,
  tableStart,
  rowIndex,
}: BuildTableRowInsertionOptions): TableRowInsertion | null => {
  const cells: PMNode[] = [];
  const rowspanUpdates: number[] = [];
  let referenceRow: number | null = rowIndex > 0 ? -1 : 0;
  if (rowIsHeader(map, table, rowIndex + referenceRow)) {
    referenceRow = rowIndex === 0 || rowIndex === map.height ? null : 0;
  }
  for (let col = 0, index = map.width * rowIndex; col < map.width; col++, index++) {
    if (rowIndex > 0 && rowIndex < map.height && map.map[index] === map.map[index - map.width]) {
      const cellPosition = map.map[index];
      if (cellPosition === undefined) {
        return null;
      }
      const cell = table.nodeAt(cellPosition);
      if (!cell) {
        return null;
      }
      const colspan: unknown = cell.attrs["colspan"];
      if (typeof colspan !== "number" || colspan < 1) {
        return null;
      }
      rowspanUpdates.push(cellPosition);
      col += colspan - 1;
      index += colspan - 1;
      continue;
    }
    const referencePosition =
      referenceRow === null ? undefined : map.map[index + referenceRow * map.width];
    const cellType =
      referencePosition === undefined
        ? tableNodeTypes(table.type.schema).cell
        : table.nodeAt(referencePosition)?.type;
    const cell = cellType?.createAndFill();
    if (!cell) {
      return null;
    }
    cells.push(cell);
  }
  if (cells.length === 0) {
    return null;
  }
  let rowPosition = tableStart;
  for (let index = 0; index < rowIndex; index++) {
    rowPosition += table.child(index).nodeSize;
  }
  return {
    tableStart,
    rowPosition,
    rowType: tableNodeTypes(table.type.schema).row,
    cells,
    rowspanUpdates,
  };
};

const populateTableRow = (row: PMNode, cellTexts: readonly string[] | undefined): PMNode => {
  if (cellTexts === undefined) {
    return row;
  }
  const cells: PMNode[] = [];
  row.forEach((cell, _offset, index) => {
    const paragraph = cell.firstChild;
    if (!paragraph?.isTextblock) {
      cells.push(cell);
      return;
    }
    const text = cellTexts[index] ?? "";
    const content = text.length > 0 ? row.type.schema.text(text) : null;
    const nextParagraph = paragraph.type.create(stripBlockIdentityAttrs(paragraph.attrs), content);
    cells.push(cell.type.create(cell.attrs, nextParagraph));
  });
  return row.type.create(row.attrs, cells);
};

const getInsertedPhysicalColumnRows = (map: TableMap, columnIndex: number): number[] => {
  const rows: number[] = [];
  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + columnIndex;
    const crossesColspan =
      columnIndex > 0 && columnIndex < map.width && map.map[index - 1] === map.map[index];
    if (!crossesColspan) {
      rows.push(row);
    }
  }
  return rows;
};

type TableColumnInsertionAction =
  | { type: "insertCell"; position: number; cell: PMNode }
  | { type: "setColspan"; position: number; attrs: Record<string, unknown> };

type BuildTableColumnInsertionActionsOptions = {
  map: TableMap;
  table: PMNode;
  columnIndex: number;
  cellTexts: readonly string[] | undefined;
};

const expandTableCellColspan = (
  cell: PMNode,
  colspanOffset: number,
): Record<string, unknown> | null => {
  const colspan: unknown = cell.attrs["colspan"];
  const rowspan: unknown = cell.attrs["rowspan"];
  const colwidth: unknown = cell.attrs["colwidth"];
  if (
    typeof colspan !== "number" ||
    !Number.isInteger(colspan) ||
    colspan < 1 ||
    typeof rowspan !== "number" ||
    !Number.isInteger(rowspan) ||
    rowspan < 1 ||
    !Number.isInteger(colspanOffset) ||
    colspanOffset < 0 ||
    colspanOffset > colspan
  ) {
    return null;
  }
  if (colwidth !== null && !Array.isArray(colwidth)) {
    return null;
  }
  if (Array.isArray(colwidth) && !colwidth.every((width) => typeof width === "number")) {
    return null;
  }
  const nextColwidth = Array.isArray(colwidth) ? [...colwidth] : null;
  nextColwidth?.splice(colspanOffset, 0, 0);
  return {
    ...cell.attrs,
    colspan: colspan + 1,
    colwidth: nextColwidth,
  };
};

const populateTableCell = (cell: PMNode, text: string): PMNode | null => {
  const paragraph = cell.firstChild;
  if (!paragraph?.isTextblock) {
    return null;
  }
  const content = text.length > 0 ? cell.type.schema.text(text) : null;
  const nextParagraph = paragraph.type.create(stripBlockIdentityAttrs(paragraph.attrs), content);
  return cell.type.create(cell.attrs, nextParagraph);
};

const buildTableColumnInsertionActions = ({
  map,
  table,
  columnIndex,
  cellTexts,
}: BuildTableColumnInsertionActionsOptions): TableColumnInsertionAction[] | null => {
  if (columnIndex < 0 || columnIndex > map.width) {
    return null;
  }
  const physicalRows = getInsertedPhysicalColumnRows(map, columnIndex);
  if ((cellTexts?.length ?? 0) > physicalRows.length) {
    return null;
  }
  let referenceColumn: number | null = columnIndex > 0 ? -1 : 0;
  if (columnIsHeader(map, table, columnIndex + referenceColumn)) {
    referenceColumn = columnIndex === 0 || columnIndex === map.width ? null : 0;
  }
  const actions: TableColumnInsertionAction[] = [];
  let physicalCellIndex = 0;
  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + columnIndex;
    if (columnIndex > 0 && columnIndex < map.width && map.map[index - 1] === map.map[index]) {
      const position = map.map[index];
      if (position === undefined) {
        return null;
      }
      const cell = table.nodeAt(position);
      if (!cell) {
        return null;
      }
      const colspanOffset = columnIndex - map.colCount(position);
      const attrs = expandTableCellColspan(cell, colspanOffset);
      if (!attrs) {
        return null;
      }
      actions.push({ type: "setColspan", position, attrs });
      row += Number(cell.attrs["rowspan"]) - 1;
      continue;
    }
    const referencePosition =
      referenceColumn === null ? undefined : map.map[index + referenceColumn];
    const cellType =
      referencePosition === undefined
        ? tableNodeTypes(table.type.schema).cell
        : table.nodeAt(referencePosition)?.type;
    const cell = cellType?.createAndFill();
    if (!cell) {
      return null;
    }
    const populatedCell = populateTableCell(cell, cellTexts?.[physicalCellIndex] ?? "");
    if (!populatedCell) {
      return null;
    }
    actions.push({
      type: "insertCell",
      position: map.positionAt(row, columnIndex, table),
      cell: populatedCell,
    });
    physicalCellIndex++;
  }
  return actions;
};

type GetTableColumnDeletionCellPositionsOptions = {
  map: TableMap;
  table: PMNode;
  columnIndex: number;
};

const getTableColumnDeletionCellPositions = ({
  map,
  table,
  columnIndex,
}: GetTableColumnDeletionCellPositionsOptions): number[] | null => {
  if (columnIndex < 0 || columnIndex >= map.width) {
    return null;
  }
  const positions = new Set<number>();
  for (let row = 0; row < map.height; row++) {
    const position = map.map[row * map.width + columnIndex];
    if (position === undefined) {
      return null;
    }
    const cell = table.nodeAt(position);
    const tableRole = cell?.type.spec["tableRole"];
    if (
      !cell ||
      (tableRole !== "cell" && tableRole !== "header_cell") ||
      cell.attrs["colspan"] !== 1 ||
      cell.attrs["cellMarker"] != null
    ) {
      return null;
    }
    positions.add(position);
  }
  return [...positions];
};
