// PARSE-WARNING-EXEMPT: the imported-id collision census reports every
// duplicate before this lossless, source-span identity repair.

import { panic } from "better-result";

import { scanSelectedStreamingXmlElements } from "./selectedXmlIdentityScan";
import {
  findAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

const bookmarkIdentity = (element: XmlElement) => {
  if (!WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? ""))
    return null;
  const name = getLocalName(element.name);
  if (name !== "bookmarkStart" && name !== "bookmarkEnd") return null;
  const attribute = findAttributeByNamespaceUri(
    element,
    WORDPROCESSINGML_NAMESPACE_URIS,
    "id",
  );
  if (!attribute) return null;
  const id = Number(attribute.value);
  if (!Number.isSafeInteger(id) || id < 0) return null;
  return { name, attribute: attribute.name, id };
};

/**
 * Give repeated bookmark ranges independent identities before editor pairing.
 * The first range keeps its id. Ends close starts in encounter order, including
 * overlapping ambiguous ranges. Only colliding ids are repaired: those require
 * unique XML ids for a valid save. Every other source byte stays verbatim.
 */
export const normalizeBookmarkIdsInXmlParts = (
  parts: ReadonlyMap<string, string>,
): ReadonlyMap<string, string> => {
  const candidates = [...parts].filter(([, xml]) =>
    /bookmark(?:Start|End)/u.test(xml),
  );
  const reserved = new Set<number>();
  const repeated = new Set<string>();
  const markersByPart = new Map<
    string,
    {
      marker: NonNullable<ReturnType<typeof bookmarkIdentity>>;
      start: number;
      end: number;
    }[]
  >();
  for (const [path, xml] of candidates) {
    const starts = new Set<number>();
    const markers: {
      marker: NonNullable<ReturnType<typeof bookmarkIdentity>>;
      start: number;
      end: number;
    }[] = [];
    const scanned = scanSelectedStreamingXmlElements(
      xml,
      ({ element, attributeValueSpans }) => {
        const marker = bookmarkIdentity(element);
        if (!marker) return;
        const span = attributeValueSpans.get(marker.attribute);
        if (!span) panic("Missing bookmark identity source span");
        markers.push({ marker, ...span });
        reserved.add(marker.id);
        if (marker.name === "bookmarkStart") {
          if (starts.has(marker.id)) repeated.add(path);
          starts.add(marker.id);
        }
      },
    );
    // Unsupported syntax belongs to the regular XML parser's boundary.
    if (scanned.status === "unsupported") {
      repeated.delete(path);
      continue;
    }
    markersByPart.set(path, markers);
  }
  if (repeated.size === 0) return parts;
  let next = 0;
  const allocate = () => {
    while (reserved.has(next)) next += 1;
    const id = next++;
    reserved.add(id);
    return id;
  };
  const normalized = new Map(parts);
  for (const [path, xml] of candidates) {
    if (!repeated.has(path)) continue;
    const seen = new Set<number>();
    const open = new Map<number, { ids: number[]; cursor: number }>();
    const chunks: string[] = [];
    let cursor = 0;
    for (const { marker, start, end } of markersByPart.get(path) ?? []) {
      let id = marker.id;
      if (marker.name === "bookmarkStart") {
        if (seen.has(id)) id = allocate();
        seen.add(marker.id);
        let queue = open.get(marker.id);
        if (!queue) {
          queue = { ids: [], cursor: 0 };
          open.set(marker.id, queue);
        }
        queue.ids.push(id);
      } else {
        const queue = open.get(marker.id);
        const paired = queue?.ids.at(queue.cursor);
        if (paired !== undefined && queue) {
          id = paired;
          queue.cursor += 1;
        }
      }
      if (id === marker.id) continue;
      chunks.push(xml.slice(cursor, start), String(id));
      cursor = end;
    }
    chunks.push(xml.slice(cursor));
    normalized.set(path, chunks.join(""));
  }
  return normalized;
};
