import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import { calculateRowSpans } from "../../docx/verticalMergeProjection";
import { expectTableCellAttrs } from "../attrs";
import { fromProseDoc } from "../conversion/fromProseDoc";

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
}: TableMergeFoldDecisionsOptions): ReadonlyMap<number, boolean> => {
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
      if (continuation === "continue") cell.formatting.vMerge = "continue";
      else delete cell.formatting.vMerge;
      delete cell.structuralChange;
    }
  }
  const projection = calculateRowSpans(model);
  const map = TableMap.get(table);
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
  return decisions;
};

/** A retained continuation must save the merge state the reviewer resolved. */
export const resolvedVisibleTableCellMergeAttrs = (cell: PMNode, mode: "accept" | "reject") => {
  const attrs = expectTableCellAttrs(cell);
  const marker = attrs.cellMarker;
  if (marker?.kind !== "merge") return panic("Cell has no merge revision");
  const continuation = mode === "accept" ? marker.verticalMerge : marker.verticalMergeOriginal;
  const formatting = { ...attrs._originalFormatting };
  if (continuation === "continue") formatting.vMerge = "continue";
  else delete formatting.vMerge;
  return { ...cell.attrs, cellMarker: null, _originalFormatting: formatting };
};
