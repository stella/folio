import type { TableBlock } from "../types";

type TableInlinePlacement =
  | { alignment: "center" }
  | { alignment: "left" | "right"; offset: number };

/** Resolve an inline table's horizontal anchor without losing RTL leading-edge semantics. */
export const resolveTableInlinePlacement = (
  table: Pick<TableBlock, "bidi" | "indent" | "justification">,
): TableInlinePlacement => {
  if (table.justification === "center") {
    return { alignment: "center" };
  }
  if (table.justification === "right") {
    return { alignment: "right", offset: 0 };
  }

  // w:tblInd adds space before the table's leading edge. Cell margins affect
  // content inside that edge and therefore must not move the table itself.
  if (table.justification === "left" || table.bidi !== true) {
    return { alignment: "left", offset: table.indent ?? 0 };
  }
  return { alignment: "right", offset: table.indent ?? 0 };
};
