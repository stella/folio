import type { FieldRun, ParagraphBlock, TextRun } from "../layout-engine/types";
import { detectBaseDirection } from "./baseDirection";

const isDirectionalTextRun = (run: ParagraphBlock["runs"][number]): run is TextRun | FieldRun =>
  run.kind === "text" || run.kind === "field";

/** Resolve the paragraph direction used by both measurement and painting. */
export const isRtlParagraph = (block: ParagraphBlock): boolean => {
  if (block.attrs?.bidi !== undefined) {
    return block.attrs.bidi;
  }
  const runs = block.runs.filter(isDirectionalTextRun);
  if (!runs.some(({ rtl }) => rtl === true)) {
    return false;
  }
  const text = runs.map((run) => (run.kind === "text" ? run.text : (run.fallback ?? ""))).join("");
  return detectBaseDirection(text) !== "ltr";
};
