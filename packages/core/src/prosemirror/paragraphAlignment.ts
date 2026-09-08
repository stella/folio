import type { ParagraphAlignment } from "../types/document";
import type { ParagraphAttrs } from "./schema/nodes";

/**
 * Read the authored paragraph-level `w:jc`, independently of the effective
 * alignment used for layout. Imported and command-authored paragraphs carry
 * explicit provenance; the final branch covers PM-created content whose
 * effective value is observably different from its style.
 */
export const directParagraphAlignment = (attrs: ParagraphAttrs): ParagraphAlignment | undefined => {
  if (attrs._originalFormatting?.alignment != null) {
    return attrs._originalFormatting.alignment;
  }
  if (attrs.alignment != null && attrs.alignment !== attrs.alignmentFromStyle) {
    return attrs.alignment;
  }
  return undefined;
};
