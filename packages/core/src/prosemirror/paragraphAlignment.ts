import type { ParagraphAlignment } from "../types/document";
import { expectParagraphPropertyState } from "./paragraphPropertyState";
import type { ParagraphAttrs } from "./schema/nodes";

/**
 * Read the authored paragraph-level `w:jc`, independently of the effective
 * alignment used for layout. Mandatory authored state is the sole source;
 * effective attrs never acquire authorship merely because they differ from a
 * style.
 */
export const directParagraphAlignment = (attrs: ParagraphAttrs): ParagraphAlignment | undefined => {
  return expectParagraphPropertyState(attrs._paragraphPropertyState).authoredPPr.alignment;
};
