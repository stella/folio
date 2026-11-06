import type { Node as PMNode } from "prosemirror-model";

/**
 * What a row omits, as its placeholders hold it now: column edits move them, and
 * the formatting the row was read with goes stale. Before its cells, this is
 * what places them. After them it stays a statement of no more than the row
 * omits: the row may stop short of the table by more than it states.
 */
export const rowGridOmissions = (
  row: PMNode,
  statedAfter: number,
): { gridBefore: number; gridAfter: number } => {
  let before = 0;
  let after = 0;
  row.forEach((cell) => {
    const colspan = Number(cell.attrs["colspan"]) || 1;
    if (cell.attrs["_omittedGridSlot"] === "before") before += colspan;
    if (cell.attrs["_omittedGridSlot"] === "after") after += colspan;
  });
  return { gridBefore: before, gridAfter: Math.min(statedAfter, after) };
};
