import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import type { TableRectangle } from "./table-mutation-plan";

export type TableBoundary = {
  before: number;
  after: number;
};

export type TableRowTarget = {
  table: PMNode;
  tableStart: number;
  tablePosition: number;
  rowIndex: number;
  rowPosition: number;
};

export type TableCellTarget = {
  tablePosition: number;
  cellPosition: number;
  cellEndPosition: number;
  leftColumnIndex: number;
  rightColumnIndex: number;
  topRowIndex: number;
  bottomRowIndex: number;
};

/** Resolve the nearest enclosing table boundary for block-adjacent insertions. */
export const findEnclosingTableBoundary = (
  doc: PMNode,
  blockFrom: number,
): TableBoundary | null => {
  const resolved = doc.resolve(blockFrom);
  for (let depth = resolved.depth; depth > 0; depth--) {
    if (resolved.node(depth).type.name === "table") {
      return { before: resolved.before(depth), after: resolved.after(depth) };
    }
  }
  return null;
};

export const findEnclosingTableRow = (doc: PMNode, blockFrom: number): TableRowTarget | null => {
  const resolved = doc.resolve(blockFrom);
  for (let rowDepth = resolved.depth; rowDepth > 0; rowDepth--) {
    if (resolved.node(rowDepth).type.spec["tableRole"] !== "row") {
      continue;
    }
    const tableDepth = rowDepth - 1;
    const table = resolved.node(tableDepth);
    if (table.type.spec["tableRole"] !== "table") {
      return null;
    }
    return {
      table,
      tableStart: resolved.start(tableDepth),
      tablePosition: resolved.before(tableDepth),
      rowIndex: resolved.index(tableDepth),
      rowPosition: resolved.before(rowDepth),
    };
  }
  return null;
};

export const findEnclosingTableCell = (doc: PMNode, blockFrom: number): TableCellTarget | null => {
  const resolved = doc.resolve(blockFrom);
  for (let cellDepth = resolved.depth; cellDepth > 0; cellDepth--) {
    const cell = resolved.node(cellDepth);
    const tableRole = cell.type.spec["tableRole"];
    if (tableRole !== "cell" && tableRole !== "header_cell") {
      continue;
    }
    const rowDepth = cellDepth - 1;
    const tableDepth = rowDepth - 1;
    if (tableDepth < 0) {
      return null;
    }
    if (
      resolved.node(rowDepth).type.spec["tableRole"] !== "row" ||
      resolved.node(tableDepth).type.spec["tableRole"] !== "table"
    ) {
      return null;
    }
    const table = resolved.node(tableDepth);
    const tableStart = resolved.start(tableDepth);
    const cellPosition = resolved.before(cellDepth);
    const cellRectangle = TableMap.get(table).findCell(cellPosition - tableStart);
    return {
      tablePosition: resolved.before(tableDepth),
      cellPosition,
      cellEndPosition: cellPosition + cell.nodeSize,
      leftColumnIndex: cellRectangle.left,
      rightColumnIndex: cellRectangle.right,
      topRowIndex: cellRectangle.top,
      bottomRowIndex: cellRectangle.bottom,
    };
  }
  return null;
};

export const tableRectangleCutsMergedCell = (map: TableMap, rectangle: TableRectangle): boolean => {
  if (
    rectangle.left < 0 ||
    rectangle.top < 0 ||
    rectangle.right > map.width ||
    rectangle.bottom > map.height
  ) {
    return false;
  }
  for (let row = rectangle.top; row < rectangle.bottom; row++) {
    const leftIndex = row * map.width + rectangle.left;
    const rightIndex = row * map.width + rectangle.right - 1;
    if (
      (rectangle.left > 0 && map.map[leftIndex] === map.map[leftIndex - 1]) ||
      (rectangle.right < map.width && map.map[rightIndex] === map.map[rightIndex + 1])
    ) {
      return true;
    }
  }
  for (let column = rectangle.left; column < rectangle.right; column++) {
    const topIndex = rectangle.top * map.width + column;
    const bottomIndex = (rectangle.bottom - 1) * map.width + column;
    if (
      (rectangle.top > 0 && map.map[topIndex] === map.map[topIndex - map.width]) ||
      (rectangle.bottom < map.height && map.map[bottomIndex] === map.map[bottomIndex + map.width])
    ) {
      return true;
    }
  }
  return false;
};
