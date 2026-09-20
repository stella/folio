import type { TableAlignment } from "../types/formatting";

export type PhysicalTableAlignment = "left" | "center" | "right";

/** Resolve an OOXML table placement to a physical inline alignment. */
export const resolveTablePlacementAlignment = (
  placement: TableAlignment,
  rightToLeft: boolean,
): PhysicalTableAlignment => {
  switch (placement) {
    case "center":
      return "center";
    case "left":
    case "start":
      return rightToLeft ? "right" : "left";
    case "right":
    case "end":
      return rightToLeft ? "left" : "right";
    default:
      return placement satisfies never;
  }
};
