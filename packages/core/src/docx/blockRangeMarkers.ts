/**
 * Range markers that stand between two blocks.
 *
 * `w:body`, `w:tc`, a header and an SDT's content all admit
 * `EG_RunLevelElts` and `EG_RangeMarkupElements` beside their paragraphs, and
 * every block container dropped them: a `w:permStart` between two paragraphs
 * is the whole of a document-protection range, so losing it removes the
 * protection from the saved file without a word. folio models none of these,
 * and their position is their meaning, so they are captured verbatim and
 * replayed where they stood — the same treatment `w:sdt`'s sibling markers
 * already get (MS-OE376 §2.5.2.30).
 *
 * Bookmarks are absent from the set on purpose: they are modelled, and the
 * block containers already relocate them into the neighbouring paragraph.
 */

const BLOCK_RANGE_MARKER_NAMES: ReadonlySet<string> = new Set([
  "commentRangeEnd",
  "commentRangeStart",
  "customXmlDelRangeEnd",
  "customXmlDelRangeStart",
  "customXmlInsRangeEnd",
  "customXmlInsRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveToRangeEnd",
  "customXmlMoveToRangeStart",
  "moveFromRangeEnd",
  "moveFromRangeStart",
  "moveToRangeEnd",
  "moveToRangeStart",
  "permEnd",
  "permStart",
]);

export const isBlockRangeMarker = (localName: string): boolean =>
  BLOCK_RANGE_MARKER_NAMES.has(localName);

/** Hand the markers collected so far to the block they stood before. */
export const attachPendingRangeMarkers = (
  block: { rawMarkersBefore?: string },
  pending: string[],
): void => {
  if (pending.length === 0) {
    return;
  }
  block.rawMarkersBefore = pending.join("");
  pending.length = 0;
};

/**
 * Markers after the last block ride on it, since there is no block after them.
 * With no block at all they are dropped: a container holding markers and no
 * content has nothing for them to delimit.
 */
export const attachTrailingRangeMarkers = (
  blocks: readonly { rawMarkersAfter?: string }[],
  pending: string[],
): void => {
  const last = blocks.at(-1);
  if (pending.length === 0 || last === undefined) {
    pending.length = 0;
    return;
  }
  last.rawMarkersAfter = pending.join("");
  pending.length = 0;
};

/** Wrap a serialized block in the markup that stood around it. */
export const withBlockRangeMarkers = (
  block: { rawMarkersBefore?: string; rawMarkersAfter?: string },
  xml: string,
): string => `${block.rawMarkersBefore ?? ""}${xml}${block.rawMarkersAfter ?? ""}`;
