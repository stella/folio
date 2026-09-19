/**
 * Parse-boundary tolerance for `w:numPr` numbering references.
 *
 * A file Word opens must never reach a panic, so a reference that resolves to
 * nothing is repaired here, on both tiers that can carry one: the paragraph and
 * the paragraph style. Downstream code (and `assertStyleNumberingReferences` in
 * particular) can then treat a numbering reference as resolvable.
 *
 * The repair writes the sentinel rather than deleting the `w:numPr`. ECMA-376
 * 17.9.18 reserves `w:numId w:val="0"` for "the removal of numbering properties
 * at a particular level in the style hierarchy"; deleting the element instead
 * removes nothing, it only uncovers the tier below, handing the paragraph its
 * `w:pStyle` numbering or the style its `w:basedOn` numbering. A source that
 * shows no number would come back numbered.
 */

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type { DocumentBody, Endnote, Footnote, HeaderFooter, Style } from "../types/document";
import type { NumberingMap } from "./numberingParser";
import { isNumberingReference, NO_NUMBERING_NUM_ID } from "./numberingReference";
import { visitDocxParagraphs } from "./paragraphTraversal";

/**
 * Whether a `w:numId` still reaches a level definition: a `w:num` that names a
 * `w:abstractNum` the numbering part defines. Both hops can dangle, and a
 * `w:num` with no `w:abstractNum` numbers nothing just as a missing one does.
 */
const resolvesNumbering = (numId: number, numbering: NumberingMap | undefined): boolean => {
  if (!numbering) {
    return false;
  }
  const abstractNumId = numbering.getAbstractNumId(numId);
  return abstractNumId !== null && numbering.getAbstract(abstractNumId) !== null;
};

/** The codes this normalisation is reported under, owned here, not at the caller. */
export const UNNUMBERED_PARAGRAPH_WARNING = PARSE_WARNING_CODES.unnumberedParagraph;
export const UNNUMBERED_STYLE_WARNING = PARSE_WARNING_CODES.unnumberedStyle;

type NormalizeNumberingReferencesInput = {
  documentBody: DocumentBody;
  numbering: NumberingMap;
  headers?: Map<string, HeaderFooter>;
  footers?: Map<string, HeaderFooter>;
  footnotes?: readonly Footnote[];
  endnotes?: readonly Endnote[];
};

type NormalizeNumberingReferencesResult = {
  unnumberedDanglingReferences: number;
};

export const normalizeNumberingReferences = ({
  documentBody,
  numbering,
  headers,
  footers,
  footnotes,
  endnotes,
}: NormalizeNumberingReferencesInput): NormalizeNumberingReferencesResult => {
  let unnumberedDanglingReferences = 0;

  visitDocxParagraphs({ documentBody, headers, footers, footnotes, endnotes }, (paragraph) => {
    const formatting = paragraph.formatting;
    const numId = formatting?.numPr?.numId;
    if (!formatting || !isNumberingReference(numId) || resolvesNumbering(numId, numbering)) {
      return;
    }
    formatting.numPr = { numId: NO_NUMBERING_NUM_ID };
    // The reference is the paragraph's own now, whatever tier stated it.
    delete formatting.numPrFromStyle;
    delete paragraph.listRendering;
    unnumberedDanglingReferences += 1;
  });

  return { unnumberedDanglingReferences };
};

type NormalizeStyleNumberingReferencesInput = {
  styles: readonly Style[];
  numbering: NumberingMap | undefined;
};

type NormalizeStyleNumberingReferencesResult = {
  /** Style ids whose dangling reference became the "no numbering" sentinel. */
  unnumberedStyleIds: string[];
};

export const normalizeStyleNumberingReferences = ({
  styles,
  numbering,
}: NormalizeStyleNumberingReferencesInput): NormalizeStyleNumberingReferencesResult => {
  const unnumberedStyleIds: string[] = [];

  for (const style of styles) {
    const pPr = style.pPr;
    const numId = pPr?.numPr?.numId;
    if (!pPr || !isNumberingReference(numId) || resolvesNumbering(numId, numbering)) {
      continue;
    }
    pPr.numPr = { numId: NO_NUMBERING_NUM_ID };
    unnumberedStyleIds.push(style.styleId);
  }

  return { unnumberedStyleIds };
};
