import { resolveTableCellPadding, type TableBlock } from "../types";

type TableInlinePlacement =
  | { alignment: "center" }
  | { alignment: "left" | "right"; offset: number };

/** Resolve an inline table's horizontal anchor without losing RTL leading-edge semantics. */
export const resolveTableInlinePlacement = (
  table: Pick<TableBlock, "bidi" | "indent" | "justification" | "rows">,
): TableInlinePlacement => {
  if (table.justification === "center") {
    return { alignment: "center" };
  }
  if (table.justification === "right") {
    return { alignment: "right", offset: 0 };
  }

  const firstCell = table.rows.at(0)?.cells.at(0);
  const firstCellPadding = firstCell ? resolveTableCellPadding(firstCell) : undefined;
  if (table.justification === "left" || table.bidi !== true) {
    return { alignment: "left", offset: table.indent ?? -(firstCellPadding?.left ?? 0) };
  }
  return { alignment: "right", offset: table.indent ?? -(firstCellPadding?.right ?? 0) };
};
