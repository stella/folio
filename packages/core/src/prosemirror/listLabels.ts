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
import type { ListRendering } from "../types/document";
import { paragraphNumberingAttr, type ParagraphNumberingSources } from "./numberingAttr";
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
 * The attrs the counter reads for a paragraph of the document model, as
 * `toProseDoc` projects them (`listRenderingAttrPatch`), without the layout
 * attrs no label depends on: a reader that walks the model (Markdown) counts
 * every list paragraph through here.
 */
export const listLabelAttrsFromRendering = (
  rendering: ListRendering | undefined,
  { numPr, numPrFromStyle }: ParagraphNumberingSources,
): ParagraphAttrs => {
  const attrs: ParagraphAttrs = {
    ...(numPr === undefined ? {} : { numPr: paragraphNumberingAttr(numPr) }),
    ...(numPrFromStyle === undefined
      ? {}
      : { numPrFromStyle: paragraphNumberingAttr(numPrFromStyle) }),
  };
  if (rendering === undefined) return attrs;
  if (rendering.marker) attrs.listMarker = rendering.marker;
  if (rendering.markerTemplate) attrs.listMarkerTemplate = rendering.markerTemplate;
  if (rendering.isBullet) attrs.listIsBullet = true;
  if (rendering.isLegal) attrs.listIsLegal = true;
  if (rendering.numFmt) attrs.listNumFmt = rendering.numFmt;
  if (rendering.markerHidden) attrs.listMarkerHidden = true;
  if (rendering.markerFormatting) attrs.listMarkerFormatting = rendering.markerFormatting;
  if (rendering.levelNumFmts) attrs.listLevelNumFmts = rendering.levelNumFmts;
  if (rendering.levelStarts) attrs.listLevelStarts = rendering.levelStarts;
  if (rendering.abstractNumId !== undefined) attrs.listAbstractNumId = rendering.abstractNumId;
  if (rendering.startOverride !== undefined) attrs.listStartOverride = rendering.startOverride;
  if (rendering.implicitChildLevelAdvances !== undefined) {
    attrs.listImplicitChildLevelAdvances = rendering.implicitChildLevelAdvances;
  }
  return attrs;
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
export const createListLabelCounter =
  (
    streams: ListCounterStreams = {
      final: createListCounterState(),
      original: createListCounterState(),
    },
  ): ListLabelCounter =>
  (attrs) => {
    const visible = advanceVisibleListMarker(attrs, streams);
    const rendered = hasListRendering(visible.counterAttrs);
    const marker = visible.marker;
    if (visible.counterAttrs.listMarkerHidden === true || !rendered) return undefined;
    const counted = marker?.trim();
    if (counted) {
      return `${counted}${foldedFieldSuffix(visible.counterAttrs)}`;
    }
    // What the page paints when the counter resolves nothing: the marker as
    // the paragraph carries it.
    return staticMarker(visible.counterAttrs);
  };
