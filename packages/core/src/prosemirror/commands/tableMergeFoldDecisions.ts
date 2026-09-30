import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import { calculateRowSpans } from "../../docx/verticalMergeProjection";
import {
  isCellMergeContinuation,
  isCellMergeStart,
  isTableCellMergeRevisionContinuation,
} from "../../docx/tableParser";
import { expectTableCellAttrs } from "../attrs";
import { fromProseDoc } from "../conversion/fromProseDoc";
import type { TableCell } from "../../types/document";

type TableMergeFoldDecisionsOptions = {
  table: PMNode;
  mode: "accept" | "reject";
  revisionSet: ReadonlySet<number> | null;
};

/** Project the resolved merge states with the reader's content and row rules. */
export const tableMergeFoldDecisions = ({
  table,
  mode,
  revisionSet,
}: TableMergeFoldDecisionsOptions) => {
  const projected = fromProseDoc(table.type.schema.topNodeType.create(null, table));
  const model = projected.package.document.content.at(0);
  if (model?.type !== "table") return panic("Merge resolution did not project a table");
  for (const row of model.rows) {
    for (const cell of row.cells) {
      const marker = cell.structuralChange;
      if (
        marker?.type !== "tableCellMerge" ||
        (revisionSet !== null && !revisionSet.has(marker.info.id))
      )
        continue;
      const continuation = mode === "accept" ? marker.verticalMerge : marker.verticalMergeOriginal;
      cell.formatting = { ...cell.formatting };
      if (isTableCellMergeRevisionContinuation(continuation)) cell.formatting.vMerge = "continue";
      else delete cell.formatting.vMerge;
      delete cell.structuralChange;
    }
  }
  const map = TableMap.get(table);
  const modelCells = new Map<string, TableCell>();
  for (const [rowIndex, row] of model.rows.entries()) {
    let column = row.formatting?.gridBefore ?? 0;
    for (const cell of row.cells) {
      modelCells.set(`${rowIndex}-${column}`, cell);
      column += cell.formatting?.gridSpan ?? 1;
    }
  }
  const origins = new Set<number>();
  const invalid = new Set<number>();
  // A tracked merge records its continuations, while its origin can still be
  // an ordinary cell. Establish the resolved chain before applying reader rules.
  const inspected = new Set<number>();
  for (const [index, position] of map.map.entries()) {
    if (inspected.has(position)) continue;
    inspected.add(position);
    const node = table.nodeAt(position);
    const marker =
      node?.type.spec["tableRole"] === "cell" || node?.type.spec["tableRole"] === "header_cell"
        ? expectTableCellAttrs(node).cellMarker
        : null;
    if (
      marker?.kind !== "merge" ||
      (revisionSet !== null && !revisionSet.has(marker.info.revisionId)) ||
      !isTableCellMergeRevisionContinuation(
        mode === "accept" ? marker.verticalMerge : marker.verticalMergeOriginal,
      )
    )
      continue;
    const row = Math.floor(index / map.width);
    const column = index % map.width;
    if (row === 0) {
      invalid.add(position);
      continue;
    }
    const cell = modelCells.get(`${row}-${column}`);
    const above = modelCells.get(`${row - 1}-${column}`);
    if (!cell || !above || (cell.formatting?.gridSpan ?? 1) !== (above.formatting?.gridSpan ?? 1)) {
      invalid.add(position);
      continue;
    }
    if (isCellMergeContinuation(above) || isCellMergeStart(above)) continue;
    const abovePosition = map.map[index - map.width];
    if (abovePosition === undefined) return panic("Merge resolution lost its origin");
    above.formatting = { ...above.formatting, vMerge: "restart" };
    origins.add(abovePosition);
  }
  const projection = calculateRowSpans(model);
  const decisions = new Map<number, boolean>();
  const visited = new Set<number>();
  for (const [index, position] of map.map.entries()) {
    if (visited.has(position)) continue;
    visited.add(position);
    const cell = table.nodeAt(position);
    if (cell?.attrs["cellMarker"]?.kind !== "merge") continue;
    const row = Math.floor(index / map.width);
    const column = index % map.width;
    const info = projection.get(`${row}-${column}`);
    if (!info) return panic("Merge resolution lost a table cell", { row, column });
    decisions.set(position, info.skip);
  }
  return { folds: decisions, origins, invalid };
};

/** Persist an implicit merge origin even when reader rules retain its continuation. */
export const resolvedTableMergeOriginAttrs = (cell: PMNode) => ({
  ...cell.attrs,
  _originalFormatting: { ...expectTableCellAttrs(cell)._originalFormatting, vMerge: "restart" },
  _preserveVMergeRestart: true,
});

type ResolvedVisibleTableCellMergeOptions = {
  cell: PMNode;
  mode: "accept" | "reject";
  origin: boolean;
};

/** A retained continuation must save the merge state the reviewer resolved. */
export const resolvedVisibleTableCellMergeAttrs = ({
  cell,
  mode,
  origin,
}: ResolvedVisibleTableCellMergeOptions) => {
  const attrs = expectTableCellAttrs(cell);
  const marker = attrs.cellMarker;
  if (marker?.kind !== "merge") return panic("Cell has no merge revision");
  const continuation = mode === "accept" ? marker.verticalMerge : marker.verticalMergeOriginal;
  const formatting = { ...attrs._originalFormatting };
  if (origin) formatting.vMerge = "restart";
  else if (isTableCellMergeRevisionContinuation(continuation)) formatting.vMerge = "continue";
  else delete formatting.vMerge;
  return {
    ...cell.attrs,
    cellMarker: null,
    _originalFormatting: formatting,
    _preserveVMergeRestart: origin,
  };
};
