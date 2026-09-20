/**
 * GENERATED FILE — do not edit.
 *
 * The containers whose content model declares one order for its children, with
 * the children written in it, derived from the committed schema graph by
 * `scripts/generate-container-children.ts`.
 *
 * It lives in the lower package because the order belongs to the model rather
 * than to one serializer, and the dependency runs one way: `@stll/docx-core`
 * cannot import `@stll/folio-core`. folio-core's declared-child table spreads
 * this one in, so both read the same order and neither restates it.
 *
 * Regenerate with:
 *
 *   bun run generate:container-children
 */

export const SEQUENCE_CHILDREN = {
  "paragraph-properties": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
} as const;

/** A container whose declared children have an order. */
export type SequenceContainer = keyof typeof SEQUENCE_CHILDREN;

/** Every child the schema declares for `Container`, in declaration order. */
export type SequenceChild<Container extends SequenceContainer> =
  (typeof SEQUENCE_CHILDREN)[Container][number];
