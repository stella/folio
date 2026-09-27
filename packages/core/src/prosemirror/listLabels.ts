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
import {
  type ParagraphNumberingOverride,
  paragraphNumberingReference,
  paragraphNumberingReferenceId,
} from "../docx/numberingReference";
import type { ListRendering } from "../types/document";
import { paragraphNumberingAttr } from "./numberingAttr";
import {
  advanceListMarker,
  advanceVisibleListMarker,
  cloneListCounterState,
  createListCounterState,
  type ListCounterState,
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

const withoutNumbering = (attrs: Readonly<ParagraphAttrs>): ParagraphAttrs => {
  const unnumbered = { ...attrs };
  Reflect.deleteProperty(unnumbered, "numPr");
  return unnumbered;
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

/** Most paragraphs: nothing to count and no marker to show. */
const isPlainParagraph = (attrs: Readonly<ParagraphAttrs>): boolean =>
  attrs.numPr == null &&
  attrs.listMarker == null &&
  attrs.listMarkerTemplate == null &&
  (attrs._propertyChanges == null || attrs._propertyChanges.length === 0);

const NO_PREVIOUS_LIST = { abstractNumId: null, fromStyle: false, numId: null };

/**
 * The attrs the counter reads for a paragraph of the document model, as
 * `toProseDoc` projects them (`listRenderingAttrPatch`), without the layout
 * attrs no label depends on: a reader that walks the model (Markdown) counts
 * every list paragraph through here.
 */
export const listLabelAttrsFromRendering = (
  rendering: ListRendering,
  numPrFromStyle: ParagraphNumberingOverride | undefined,
): ParagraphAttrs => {
  const attrs: ParagraphAttrs = {
    numPr: paragraphNumberingAttr(
      paragraphNumberingReference({ numId: rendering.numId, ilvl: rendering.level }),
    ),
  };
  if (numPrFromStyle) attrs.numPrFromStyle = paragraphNumberingAttr(numPrFromStyle);
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
export const createListLabelCounter = (): ListLabelCounter => {
  const final = createListCounterState();
  // The page also counts a tracked change's original state, in a second
  // stream. Until the first tracked paragraph that stream advances exactly as
  // `final` does, except that unnumbered paragraphs never reach it, so it is
  // forked from `final` there rather than advanced twice all along.
  let streams: ListCounterStreams | null = null;
  let originalPreviousList: NonNullable<ListCounterState["previousList"]> = NO_PREVIOUS_LIST;
  const advance = (counted: Readonly<ParagraphAttrs>): string | null => {
    const tracked =
      counted.pPrMark != null ||
      (counted._propertyChanges != null && counted._propertyChanges.length > 0);
    if (streams === null && tracked) {
      const original = cloneListCounterState(final);
      original.previousList = { ...originalPreviousList };
      streams = { final, original };
    }
    if (streams !== null) {
      return advanceVisibleListMarker(counted, streams).marker;
    }
    const marker = advanceListMarker(counted, final);
    if (counted.numPr) {
      originalPreviousList = final.previousList ?? NO_PREVIOUS_LIST;
    }
    return marker;
  };
  return (attrs) => {
    if (isPlainParagraph(attrs)) {
      // All the counter does for a paragraph with no numbering, now or in a
      // tracked change, is end the run of list items before it.
      final.previousList = NO_PREVIOUS_LIST;
      return undefined;
    }
    // A paragraph numbered at a level its list does not define carries no
    // rendering (see `hasListRendering`): Word paints no marker for it and it
    // counts as unnumbered, as the parser counts it.
    const rendered = hasListRendering(attrs);
    const marker = advance(rendered ? attrs : withoutNumbering(attrs));
    if (attrs.listMarkerHidden === true || !rendered) {
      return undefined;
    }
    // An unnumbered paragraph's counted marker is a removed number (a tracked
    // change's previous state), which the page strikes through.
    const counted =
      paragraphNumberingReferenceId(attrs.numPr) === undefined ? null : marker?.trim();
    if (counted) {
      return `${counted}${foldedFieldSuffix(attrs)}`;
    }
    // What the page paints when the counter resolves nothing: the marker as
    // the paragraph carries it.
    return staticMarker(attrs);
  };
};
