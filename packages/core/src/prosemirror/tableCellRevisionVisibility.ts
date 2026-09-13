import { panic } from "better-result";

import type { TableCellAttrs } from "./schema/nodes";

/** Whether a tracked cell remains in the requested reviewed view. */
export const isTableCellRetainedInReviewView = (
  cellMarker: TableCellAttrs["cellMarker"],
  view: "original" | "final",
): boolean => {
  if (!cellMarker) {
    return true;
  }
  switch (cellMarker.kind) {
    case "merge":
      return true;
    case "ins":
      return view === "final";
    case "del":
      return view === "original";
    default: {
      const unreachable: never = cellMarker;
      return panic("Unhandled table cell revision", { marker: unreachable });
    }
  }
};
