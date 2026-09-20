/**
 * Write the attributes of a range marker in schema order.
 *
 * The counterpart of `../markupRangeMarker.ts`: one writer per schema type, so
 * an attribute the model carries cannot be dropped by whichever marker forgot
 * to mention it. Attribute order follows the declaration order of
 * `CT_MarkupRange` → `CT_Bookmark` → `CT_MoveBookmark`.
 */

import type {
  BookmarkEnd,
  BookmarkRangeMarker,
  BookmarkStart,
  MarkupRangeMarker,
  MoveBookmarkMarker,
} from "../../types/document";
import { escapeXmlAttribute } from "@stll/docx-core";

const attribute = (name: string, value: string | number): string =>
  `${name}="${escapeXmlAttribute(String(value))}"`;

export const markupRangeAttributes = (marker: MarkupRangeMarker): string[] => {
  const attributes = [attribute("w:id", marker.id)];
  if (marker.displacedByCustomXml !== undefined) {
    attributes.push(attribute("w:displacedByCustomXml", marker.displacedByCustomXml));
  }
  return attributes;
};

export const bookmarkRangeAttributes = (marker: BookmarkRangeMarker): string[] => {
  const attributes = markupRangeAttributes(marker);
  if (marker.colFirst !== undefined) {
    attributes.push(attribute("w:colFirst", marker.colFirst));
  }
  if (marker.colLast !== undefined) {
    attributes.push(attribute("w:colLast", marker.colLast));
  }
  attributes.push(attribute("w:name", marker.name));
  return attributes;
};

/**
 * One bookmark marker, wherever it stands.
 *
 * The same element is a paragraph child, a block, a child of `w:tr` and a
 * child of `w:tbl`, so the four serializers share one writer: an attribute
 * added to `CT_Bookmark` reaches every level at once.
 */
export const serializeBookmarkMarker = (marker: BookmarkStart | BookmarkEnd): string =>
  marker.type === "bookmarkStart"
    ? `<w:bookmarkStart ${bookmarkRangeAttributes(marker).join(" ")}/>`
    : `<w:bookmarkEnd ${markupRangeAttributes(marker).join(" ")}/>`;

/** `w:author` is required, so it is always written; `w:date` only when known. */
export const moveBookmarkAttributes = (marker: MoveBookmarkMarker): string[] => {
  const attributes = bookmarkRangeAttributes(marker);
  attributes.push(attribute("w:author", marker.author.trim() || "Unknown"));
  if (marker.date !== undefined && marker.date.trim().length > 0) {
    attributes.push(attribute("w:date", marker.date.trim()));
  }
  return attributes;
};
