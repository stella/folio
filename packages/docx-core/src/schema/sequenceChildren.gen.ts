/**
 * GENERATED FILE — do not edit.
 *
 * The containers whose content model declares one order for its children, with
 * the children written in it, derived from the committed schema graph by
 * `scripts/generate-container-children.ts`.
 *
 * It lives in the lower package because two packages write a property set and
 * the dependency runs one way: `@stll/docx-core` compiles a `w:rPr` from a
 * legal source and cannot import `@stll/folio-core`. folio-core's
 * declared-child table spreads this one in, so both write the same order and
 * neither restates it.
 *
 * Regenerate with:
 *
 *   bun run generate:container-children
 */

export const SEQUENCE_CHILDREN = {
  "table-properties": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"],
  "run-properties": ["ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"],
  "section-properties": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
} as const;

/** A container whose declared children have an order. */
export type SequenceContainer = keyof typeof SEQUENCE_CHILDREN;

/** Every child the schema declares for `Container`, in declaration order. */
export type SequenceChild<Container extends SequenceContainer> =
  (typeof SEQUENCE_CHILDREN)[Container][number];
