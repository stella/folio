import type { BookmarkEnd, BookmarkStart } from "../model/document";
import { escapeXmlAttribute } from "../serialize/xmlEscape";
import {
  scanStreamingXmlElements,
  type XmlElement,
} from "../xml/streamingXmlParser";
import {
  OOXML_NAMESPACE_SCOPE,
  OOXML_NS,
  resolveNamespaceUri,
} from "../xml/xmlNamespaceContext";

type CapturedBookmark =
  | { marker: BookmarkStart; idAttribute: string; nameAttribute: string }
  | { marker: BookmarkEnd; idAttribute: string };

const isWordNamespace = (uri: string | undefined) =>
  uri === OOXML_NS.w ||
  uri === "http://purl.oclc.org/ooxml/wordprocessingml/main";

const nativeAttribute = (node: XmlElement, localName: string) =>
  Object.entries(node.attributes ?? {}).find(([key]) => {
    const index = key.indexOf(":");
    return (
      index > 0 &&
      key.slice(index + 1) === localName &&
      isWordNamespace(
        resolveNamespaceUri(node.namespaceScope, key.slice(0, index)),
      )
    );
  });

export const readCapturedBookmark = (
  node: XmlElement,
): CapturedBookmark | undefined => {
  if (!isWordNamespace(node.namespaceUri)) return undefined;
  const { name = "" } = node;
  const localName = name.slice(name.indexOf(":") + 1);
  if (localName !== "bookmarkStart" && localName !== "bookmarkEnd")
    return undefined;
  const idEntry = nativeAttribute(node, "id");
  if (idEntry === undefined || typeof idEntry[1] !== "string") return undefined;
  const [idAttribute, rawId] = idEntry;
  const spelling = rawId.replace(/^[ \t\r\n]+|[ \t\r\n]+$/gu, "");
  if (!/^[+-]?\d+$/u.test(spelling)) return undefined;
  const id = Number(spelling);
  if (!Number.isSafeInteger(id)) return undefined;
  if (localName === "bookmarkEnd")
    return { marker: { type: localName, id }, idAttribute };
  const nameEntry = nativeAttribute(node, "name");
  if (nameEntry === undefined || typeof nameEntry[1] !== "string")
    return undefined;
  return {
    marker: { type: localName, id, name: nameEntry[1] },
    idAttribute,
    nameAttribute: nameEntry[0],
  };
};

const scanBookmarks = (
  xml: string,
  visit: (
    bookmark: CapturedBookmark,
    attributes: ReadonlyMap<string, { start: number; end: number }>,
  ) => void,
) =>
  scanStreamingXmlElements(
    xml,
    ({ element, attributeValueSpans }) => {
      const bookmark = readCapturedBookmark(element);
      if (bookmark !== undefined) visit(bookmark, attributeValueSpans);
    },
    OOXML_NAMESPACE_SCOPE,
  );

/** Bookmarks inside a captured equation still occupy package identities. */
export const capturedMathBookmarks = (
  xml: string,
): (BookmarkStart | BookmarkEnd)[] => {
  const bookmarks: (BookmarkStart | BookmarkEnd)[] = [];
  const result = scanBookmarks(xml, ({ marker }) => bookmarks.push(marker));
  return result.status === "scanned" ? bookmarks : [];
};

/** Restamp only the known attributes, leaving the captured equation verbatim. */
export const withCapturedBookmarkAliases = (
  xml: string,
  aliases: ReadonlyMap<number, { id: number; name: string }>,
): string => {
  const changes: { start: number; end: number; value: string }[] = [];
  const result = scanBookmarks(xml, (bookmark, attributes) => {
    const alias = aliases.get(bookmark.marker.id);
    if (alias === undefined) return;
    const id = attributes.get(bookmark.idAttribute);
    if (id !== undefined) changes.push({ ...id, value: String(alias.id) });
    if ("nameAttribute" in bookmark) {
      const name = attributes.get(bookmark.nameAttribute);
      if (name !== undefined)
        changes.push({ ...name, value: escapeXmlAttribute(alias.name) });
    }
  });
  if (result.status !== "scanned" || changes.length === 0) return xml;
  const parts: string[] = [];
  let start = 0;
  for (const change of changes.sort((a, b) => a.start - b.start)) {
    parts.push(xml.slice(start, change.start), change.value);
    start = change.end;
  }
  parts.push(xml.slice(start));
  return parts.join("");
};
