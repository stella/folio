/**
 * The number or bullet a paragraph shows beside its text, as a reader reports
 * it: counted in document order by the counter the page paints its markers
 * with (`advanceVisibleListMarker`), so a reader's labels follow every edit.
 * The `listMarker` attr is the marker the parser resolved when the document
 * was opened; an item inserted, deleted, moved or restarted since leaves it
 * stale for that item and every item after it.
 *
 * A reader counts inside the walk it already makes over the document (the AI
 * snapshot's, Markdown's), one step per paragraph, so an edit costs nothing
 * until something reads the document.
 */

import { bulletMarkerFontName, convertBulletToUnicode } from "../docx/bulletMarkers";
import { paragraphNumberingReferenceId } from "../docx/numberingReference";
import {
  advanceVisibleListMarker,
  createListCounterState,
  type ListCounterStreams,
} from "./listMarker";
import { hasListRendering } from "./listRenderingAttrs";
import type { ParagraphAttrs } from "./schema/nodes";

/**
 * The cached text of a `LISTNUM` field the parser folded into the marker
 * (`7.1\t(a)`). The level's template does not carry it, so it is kept as read.
 */
const foldedFieldSuffix = (attrs: Readonly<ParagraphAttrs>): string => {
  const marker = attrs.listMarker;
  const template = attrs.listMarkerTemplate;
  // Without a template the marker itself is resolved, suffix and all.
  if (!marker || !template || template.includes("\t")) {
    return "";
  }
  const tab = marker.indexOf("\t");
  return tab === -1 ? "" : marker.slice(tab);
};

const staticMarker = (attrs: Readonly<ParagraphAttrs>): string | undefined => {
  const marker = attrs.listMarker;
  if (!marker) {
    return undefined;
  }
  const text = (
    attrs.listIsBullet
      ? convertBulletToUnicode(marker, bulletMarkerFontName(attrs.listMarkerFormatting))
      : marker
  ).trim();
  return text || undefined;
};

/**
 * Advances past one paragraph and answers the label it shows, or `undefined`
 * when it shows none.
 */
export type ListLabelCounter = (attrs: Readonly<ParagraphAttrs>) => string | undefined;

/**
 * A counter over paragraph attrs in document order. Call it for every
 * paragraph, numbered or not: whether a list continues depends on what stands
 * between its items.
 */
export const createListLabelCounter = (): ListLabelCounter => {
  const streams: ListCounterStreams = {
    final: createListCounterState(),
    original: createListCounterState(),
  };
  return (attrs) => {
    const visible = advanceVisibleListMarker(attrs, streams);
    if (attrs.listMarkerHidden === true || !hasListRendering(attrs)) {
      return undefined;
    }
    // An unnumbered paragraph's counted marker is a removed number (a tracked
    // change's previous state), which the page strikes through.
    const counted =
      paragraphNumberingReferenceId(attrs.numPr) === undefined ? null : visible.marker?.trim();
    if (counted) {
      return `${counted}${foldedFieldSuffix(attrs)}`;
    }
    // What the page paints when the counter resolves nothing: the marker as
    // the paragraph carries it.
    return staticMarker(attrs);
  };
};
