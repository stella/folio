/**
 * The attributes shared by every WordprocessingML range marker.
 *
 * `w:bookmarkStart`, `w:bookmarkEnd`, `w:commentRangeStart`,
 * `w:commentRangeEnd` and the four move-range markers are all `CT_MarkupRange`
 * or a derivation of it, but each used to read and write its own hand-picked
 * subset of the attribute set, so each lost a different part of it on save.
 * One reader and one writer per schema type is what keeps a marker from
 * quietly forgetting an attribute again.
 */

import type {
  BookmarkRangeMarker,
  DisplacedByCustomXml,
  MarkupRangeMarker,
  MoveBookmarkMarker,
} from "../types/document";
import { getAttribute, parseNumericAttribute, type XmlElement } from "./xmlParser";

/** ST_DisplacedByCustomXml. A value outside it is not a placement folio can honour. */
const DISPLACED_BY_CUSTOM_XML: ReadonlySet<string> = new Set<DisplacedByCustomXml>([
  "next",
  "prev",
]);

const parseDisplacedByCustomXml = (node: XmlElement): DisplacedByCustomXml | undefined => {
  const value = getAttribute(node, "w", "displacedByCustomXml") ?? "";
  // SAFETY: membership in the set is membership in the union.
  return DISPLACED_BY_CUSTOM_XML.has(value) ? (value as DisplacedByCustomXml) : undefined;
};

export const parseMarkupRangeMarker = (node: XmlElement): MarkupRangeMarker => {
  const marker: MarkupRangeMarker = { id: parseNumericAttribute(node, "w", "id") ?? 0 };
  const displacedByCustomXml = parseDisplacedByCustomXml(node);
  if (displacedByCustomXml !== undefined) {
    marker.displacedByCustomXml = displacedByCustomXml;
  }
  return marker;
};

export const parseBookmarkRangeMarker = (node: XmlElement): BookmarkRangeMarker => {
  const marker: BookmarkRangeMarker = {
    ...parseMarkupRangeMarker(node),
    name: getAttribute(node, "w", "name") ?? "",
  };
  const colFirst = parseNumericAttribute(node, "w", "colFirst");
  if (colFirst !== undefined) {
    marker.colFirst = colFirst;
  }
  const colLast = parseNumericAttribute(node, "w", "colLast");
  if (colLast !== undefined) {
    marker.colLast = colLast;
  }
  return marker;
};

/**
 * `w:author` is required on `CT_MoveBookmark`, so an absent one takes the same
 * `Unknown` fallback a tracked change takes rather than leaving the saved
 * package short of an attribute the schema demands. `w:date` is required too,
 * but inventing a timestamp would state a fact about the document that is not
 * true, so an absent date stays absent.
 */
export const parseMoveBookmarkMarker = (node: XmlElement): MoveBookmarkMarker => {
  const author = (getAttribute(node, "w", "author") ?? "").trim();
  const marker: MoveBookmarkMarker = {
    ...parseBookmarkRangeMarker(node),
    author: author.length > 0 ? author : "Unknown",
  };
  const date = (getAttribute(node, "w", "date") ?? "").trim();
  if (date.length > 0) {
    marker.date = date;
  }
  return marker;
};
