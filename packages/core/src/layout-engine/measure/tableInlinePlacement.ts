import type { TableBlock } from "../types";

type TableInlinePlacement =
  | { alignment: "center" }
  | { alignment: "left" | "right"; offset: number };

/** Resolve an inline table's horizontal anchor without losing RTL leading-edge semantics. */
export const resolveTableInlinePlacement = (
  table: Pick<TableBlock, "bidi" | "indent" | "justification">,
  rowJustification?: TableBlock["justification"],
): TableInlinePlacement => {
  const logicalJustification = rowJustification ?? table.justification ?? "left";
  if (logicalJustification === "center") {
    return { alignment: "center" };
  }

  // w:tblInd adds space before the table's leading edge. Cell margins affect
  // content inside that edge and therefore must not move the table itself.
  // OOXML justification is logical: bidiVisual mirrors left/right only after
  // deciding whether the leading-edge indent applies.
  const offset = logicalJustification === "left" ? (table.indent ?? 0) : 0;
  if (table.bidi !== true) {
    return { alignment: logicalJustification, offset };
  }
  return {
    alignment: logicalJustification === "left" ? "right" : "left",
    offset,
  };
};

type ResolveTableInlineOffsetOptions = {
  table: Pick<TableBlock, "bidi" | "indent" | "justification">;
  rowJustification?: TableBlock["justification"];
  frameWidth: number;
  tableWidth: number;
};

/** Resolve a table's physical inline offset within its current content frame. */
export const resolveTableInlineOffset = ({
  table,
  rowJustification,
  frameWidth,
  tableWidth,
}: ResolveTableInlineOffsetOptions): number => {
  const placement = resolveTableInlinePlacement(table, rowJustification);
  if (placement.alignment === "center") {
    return (frameWidth - tableWidth) / 2;
  }
  if (placement.alignment === "right") {
    return frameWidth - tableWidth - placement.offset;
  }
  return placement.offset;
};
