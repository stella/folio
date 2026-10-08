import type { RunContent } from "../../types/document";

/** These source units render as native text while keeping one logical atom each. */
export const HYPHEN_TEXT_CARRIERS = {
  softHyphen: "\u00ad",
  noBreakHyphen: "\u2011",
} as const satisfies Record<
  Extract<RunContent, { type: "softHyphen" | "noBreakHyphen" }>["type"],
  string
>;
