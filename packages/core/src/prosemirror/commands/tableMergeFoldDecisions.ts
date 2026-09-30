import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import { tableCellHasMeaningfulContent } from "../../docx/verticalMergeProjection";
import {
  isCellMergeContinuation,
  isCellMergeStart,
  isTableCellMergeRevisionContinuation,
} from "../../docx/tableParser";
import { expectTableCellAttrs } from "../attrs";
import { standaloneTableCellFromProseMirror } from "../conversion/fromProseDoc";

type TableMergeFoldDecisionsOptions = {
  table: PMNode;
  mode: "accept" | "reject";
  revisionSet: ReadonlySet<number> | null;
};

/** Resolve raw merge states without converting unrelated table revisions. */
export const tableMergeFoldDecisions = ({
  table,
  mode,
  revisionSet,
}: TableMergeFoldDecisionsOptions) => {
  const decisions = new Map<number, boolean>();
  const origins = new Set<number>();
  const invalid = new Set<number>();
  const map = TableMap.get(table);
  const visited = new Set<number>();
  for (const [index, position] of map.map.entries()) {
    if (visited.has(position)) continue;
    visited.add(position);
    const cell = table.nodeAt(position);
    if (cell?.attrs["cellMarker"]?.kind !== "merge") continue;
    const marker = expectTableCellAttrs(cell).cellMarker;
    if (
      marker?.kind !== "merge" ||
      (revisionSet !== null && !revisionSet.has(marker.info.revisionId))
    )
      continue;
    const continues = isTableCellMergeRevisionContinuation(
      mode === "accept" ? marker.verticalMerge : marker.verticalMergeOriginal,
    );
    // Rejection restores the previous span, including its captured content.
    // Acceptance keeps authored continuation content visible, as on import.
    decisions.set(
      position,
      continues &&
        (mode === "reject" ||
          !tableCellHasMeaningfulContent(standaloneTableCellFromProseMirror(cell))),
    );
    if (!continues) continue;
    const row = Math.floor(index / map.width);
    if (row === 0) {
      invalid.add(position);
      continue;
    }
    const abovePosition = map.map[index - map.width];
    if (abovePosition === undefined) return panic("Merge resolution lost its origin");
    const above = table.nodeAt(abovePosition);
    if (!above) return panic("Merge resolution lost its origin");
    const rectangle = map.findCell(position);
    const aboveRectangle = map.findCell(abovePosition);
    if (
      rectangle.left !== aboveRectangle.left ||
      rectangle.right !== aboveRectangle.right ||
      rectangle.top !== aboveRectangle.bottom
    ) {
      invalid.add(position);
      continue;
    }
    const aboveAttrs = expectTableCellAttrs(above);
    const aboveMarker = aboveAttrs.cellMarker;
    const aboveContinues =
      aboveMarker?.kind === "merge" &&
      (revisionSet === null || revisionSet.has(aboveMarker.info.revisionId)) &&
      isTableCellMergeRevisionContinuation(
        mode === "accept" ? aboveMarker.verticalMerge : aboveMarker.verticalMergeOriginal,
      );
    const aboveModel = standaloneTableCellFromProseMirror(above);
    // A visible continuation closes the reader's merge chain. Cells below it
    // cannot fold into it when accepting that chain.
    if (
      mode === "accept" &&
      ((aboveContinues && decisions.get(abovePosition) === false) ||
        (isCellMergeContinuation(aboveModel) && tableCellHasMeaningfulContent(aboveModel)))
    )
      decisions.set(position, false);
    if (aboveContinues || isCellMergeContinuation(aboveModel) || isCellMergeStart(aboveModel))
      continue;
    origins.add(abovePosition);
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
