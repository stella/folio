/**
 * GENERATED FILE — do not edit.
 *
 * Every child the Transitional content model declares for a container folio
 * dispatches, derived from the committed schema graph by
 * `scripts/generate-container-children.ts`. The shared child dispatcher makes
 * a handler map total over these names, so a container cannot gain a declared
 * child without somebody deciding whether it is modelled or captured.
 * Regenerate with:
 *
 *   bun run generate:container-children
 */

export const CONTAINER_CHILDREN = {
  "w:comment": ["altChunk", "bookmarkEnd", "bookmarkStart", "commentRangeEnd", "commentRangeStart", "customXml", "customXmlDelRangeEnd", "customXmlDelRangeStart", "customXmlInsRangeEnd", "customXmlInsRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveFromRangeStart", "customXmlMoveToRangeEnd", "customXmlMoveToRangeStart", "del", "ins", "moveFrom", "moveFromRangeEnd", "moveFromRangeStart", "moveTo", "moveToRangeEnd", "moveToRangeStart", "p", "permEnd", "permStart", "proofErr", "sdt", "tbl"],
} as const;

/** A container the shared child dispatcher covers. */
export type DispatchedContainer = keyof typeof CONTAINER_CHILDREN;

/** Every child name the schema declares for `Container`. */
export type DeclaredChild<Container extends DispatchedContainer> =
  (typeof CONTAINER_CHILDREN)[Container][number];
