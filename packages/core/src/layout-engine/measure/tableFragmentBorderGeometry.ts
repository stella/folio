import type { CellBorderSpec, TableBlock, TableFragment, TableMeasure } from "../types";
import {
  buildTableCellGrid,
  buildTableCellPlacements,
  type TableCellPlacements,
} from "./tableCellGrid";

type TableFragmentBottomBorder = {
  readonly left: number;
  readonly width: number;
  readonly border: CellBorderSpec;
};

type TableFragmentBottomBordersOptions = {
  readonly fragment: TableFragment;
  readonly block: TableBlock;
  readonly measure: TableMeasure;
  readonly placements?: TableCellPlacements;
};

const isVisible = (border: CellBorderSpec | undefined): border is CellBorderSpec =>
  border !== undefined && border.style !== "none" && border.style !== "nil";

export const tableFragmentBottomBorders = ({
  fragment,
  block,
  measure,
  placements: suppliedPlacements,
}: TableFragmentBottomBordersOptions): readonly TableFragmentBottomBorder[] => {
  if (fragment.continuesOnNext !== true || fragment.bottomClip === undefined) return [];

  const splitRow = block.rows.at(fragment.toRow - 1);
  if (!splitRow) return [];
  const placements =
    suppliedPlacements ??
    buildTableCellPlacements({
      grid: buildTableCellGrid(block.rows, measure.columnWidths.length),
      columnWidths: measure.columnWidths,
      bidi: block.bidi === true,
    });
  return splitRow.cells.flatMap((cell) => {
    const placement = placements.get(cell);
    const border = cell.borders?.bottom;
    return placement && isVisible(border)
      ? [{ left: placement.left, width: placement.width, border }]
      : [];
  });
};
