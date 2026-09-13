import { panic } from "better-result";

import type { TableCellAttrs } from "./schema/nodes";

/** Whether a tracked cell remains in the requested reviewed view. */
export const isTableCellRetainedInReviewView = (
  kind: NonNullable<TableCellAttrs["cellMarker"]>["kind"] | undefined,
  view: "original" | "final",
): boolean => {
  if (kind === undefined) {
    return true;
  }
  switch (kind) {
    case "merge":
      return true;
    case "ins":
      return view === "final";
    case "del":
      return view === "original";
    default: {
      const unreachable: never = kind;
      return panic("Unhandled table cell revision", { marker: unreachable });
    }
  }
};
