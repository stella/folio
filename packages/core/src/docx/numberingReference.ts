/**
 * The one reader over a `<w:numPr>` element.
 *
 * The union, the reserved id and the cascade fold are `@stll/docx-core`'s,
 * because `docx-core` owns the model and used to hand-inline its own copy of
 * the sentinel test (`validate/docx.ts`) rather than import across the package
 * boundary. This module adds the half that needs an XML element and re-exports
 * the rest, so a consumer reaches one name from `@stll/folio-core/docx`
 * instead of mirroring a constant.
 */

export {
  isNumberingReference,
  mergeParagraphNumbering,
  NO_NUMBERING_NUM_ID,
  NO_PARAGRAPH_NUMBERING,
  paragraphNumberingFromSlots,
  paragraphNumberingLevel,
  paragraphNumberingReference,
  paragraphNumberingReferenceId,
  paragraphNumberingSlots,
  resolveParagraphNumbering,
  sameEffectiveParagraphNumbering,
  sameStatedParagraphNumbering,
  type ParagraphNumberingOverride,
  type ParagraphNumberingReference,
  type ParagraphNumberingSlots,
  type ResolvedParagraphNumbering,
} from "@stll/docx-core/model";

import {
  paragraphNumberingFromSlots,
  type ParagraphNumberingOverride,
} from "@stll/docx-core/model";

import { findChild, parseNumberingLevelAttribute, parseNumericAttribute } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/**
 * What a `<w:numPr>` states: absent (the tier states nothing), the reserved
 * cancellation, a reference, or a level alone.
 *
 * `w:numberingChange` is not read here. It is a historical record of what a
 * reviewer changed, not a statement about this paragraph's numbering, and it
 * travels as captured markup.
 */
export const readParagraphNumbering = (
  numPr: XmlElement | null | undefined,
): ParagraphNumberingOverride | undefined => {
  if (!numPr) {
    return undefined;
  }
  const numIdElement = findChild(numPr, "w", "numId");
  const ilvlElement = findChild(numPr, "w", "ilvl");
  return paragraphNumberingFromSlots({
    numId: numIdElement ? parseNumericAttribute(numIdElement, "w", "val") : undefined,
    ilvl: ilvlElement ? parseNumberingLevelAttribute(ilvlElement) : undefined,
  });
};
