import { resolveTableCellPadding, type TableBlock } from "../types";

type TableInlinePlacement =
  | { alignment: "center" }
  | { alignment: "left" | "right"; offset: number };

type LeadingEdgeTable = Pick<
  TableBlock,
  "bidi" | "indent" | "indentCompatibility" | "justification" | "rows"
>;

/**
 * Distance the leading border is pulled back so the leading cell's text lands
 * on `w:tblInd`. Zero unless the document uses the pre-Word-2013 indent
 * semantics, where the indent measures to that text edge rather than to the
 * border. Only the leading, indent-bearing edge compensates; centered and
 * trailing-edge tables are placed by their border box.
 */
const leadingEdgeCompensation = (table: LeadingEdgeTable): number => {
  if (table.indentCompatibility?.type !== "legacy") {
    return 0;
  }
  const padding = resolveTableCellPadding(table.rows.at(0)?.cells.at(0));
  return table.bidi === true ? padding.right : padding.left;
};

/** Resolve an inline table's horizontal anchor without losing RTL leading-edge semantics. */
export const resolveTableInlinePlacement = (
  table: LeadingEdgeTable,
  rowJustification?: TableBlock["justification"],
): TableInlinePlacement => {
  const logicalJustification = rowJustification ?? table.justification ?? "left";
  if (logicalJustification === "center") {
    return { alignment: "center" };
  }

  // OOXML justification is logical: bidiVisual mirrors left/right only after
  // deciding whether the leading-edge indent applies.
  const offset =
    logicalJustification === "left" ? (table.indent ?? 0) - leadingEdgeCompensation(table) : 0;
  if (table.bidi !== true) {
    return { alignment: logicalJustification, offset };
  }
  return {
    alignment: logicalJustification === "left" ? "right" : "left",
    offset,
  };
};

type ResolveTableInlineOffsetOptions = {
  table: LeadingEdgeTable;
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
