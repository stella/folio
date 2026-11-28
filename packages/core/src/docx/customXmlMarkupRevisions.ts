/**
 * The tracked insertion and deletion of a block custom XML element's own tags.
 *
 * A custom XML element's start and end tags are markup, not content, so no
 * `w:ins` or `w:del` can hold them: ECMA-376 Part 1 records their insertion
 * with a `w:customXmlInsRangeStart` before the element and the matching
 * `w:customXmlInsRangeEnd` after it (§17.13.5.6–7), and their deletion with
 * `w:customXmlDelRangeStart` and `w:customXmlDelRangeEnd` (§17.13.5.4–5).
 * The reader keeps a pair standing directly around an element as that
 * element's record; any other range marker stays the captured markup it was.
 */
import type {
  BlockContent,
  BlockCustomXml,
  CustomXmlMarkupRevision,
  TrackedChangeInfo,
} from "../types/document";
import {
  getAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  resolveAttributeNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

type MarkupKind = "insertion" | "deletion";

const MARKERS: ReadonlyMap<
  string,
  { kind: MarkupKind; edge: "start" | "end" }
> = new Map([
  ["customXmlInsRangeStart", { kind: "insertion", edge: "start" }],
  ["customXmlInsRangeEnd", { kind: "insertion", edge: "end" }],
  ["customXmlDelRangeStart", { kind: "deletion", edge: "start" }],
  ["customXmlDelRangeEnd", { kind: "deletion", edge: "end" }],
]);

/** The attributes a marker may carry for the model to stand for it whole. */
const START_ATTRIBUTES = new Set(["id", "author", "date"]);
const END_ATTRIBUTES = new Set(["id"]);

type ReadMarker = {
  kind: MarkupKind;
  edge: "start" | "end";
  id: number;
  info: TrackedChangeInfo;
};

const readElement = (xml: string): XmlElement | undefined => {
  try {
    // A capture is replayed under the part's own bindings, so it is read in
    // the scope the part's root declares.
    return parseXml(xml, OOXML_NAMESPACE_SCOPE);
  } catch {
    return undefined;
  }
};

/**
 * The range marker a captured block stands for, when the model can write it
 * back as it was read: only the attributes the record carries, a decimal id
 * and an author.
 */
const markerOf = (block: BlockContent | undefined): ReadMarker | undefined => {
  if (block?.type !== "preservedBlock" || !block.xml.includes("customXml"))
    return undefined;
  const element = readElement(block.xml)?.elements?.find(
    ({ type }) => type === "element",
  );
  if (element === undefined) return undefined;
  if (!WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? ""))
    return undefined;
  const marker = MARKERS.get(getLocalName(element.name));
  if (marker === undefined) return undefined;
  if ((element.elements ?? []).length > 0) return undefined;
  const allowed = marker.edge === "start" ? START_ATTRIBUTES : END_ATTRIBUTES;
  for (const name of Object.keys(element.attributes ?? {})) {
    if (name === "xmlns" || name.startsWith("xmlns:")) continue;
    if (
      !allowed.has(getLocalName(name)) ||
      !WORDPROCESSINGML_NAMESPACE_URIS.has(
        resolveAttributeNamespaceUri(element, name) ?? "",
      )
    )
      return undefined;
  }
  const read = (name: string) =>
    getAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_URIS, name);
  const rawId = read("id");
  if (rawId === null || !/^\d+$/u.test(rawId)) return undefined;
  const id = Number.parseInt(rawId, 10);
  if (!Number.isSafeInteger(id)) return undefined;
  const info: TrackedChangeInfo = { id, author: "" };
  if (marker.edge === "start") {
    const author = read("author");
    if (author === null || author.trim() !== author || author.length === 0)
      return undefined;
    info.author = author;
    const date = read("date");
    if (date !== null) {
      if (date.trim() !== date || date.length === 0) return undefined;
      info.date = date;
    }
  }
  return { ...marker, id, info };
};

/**
 * `blocks` with each range-marker pair standing directly around a block
 * custom XML element read as that element's own record, innermost first. A
 * marker that pairs with nothing beside an element, or that carries markup
 * the record does not hold, stays as captured.
 */
export const withCustomXmlMarkupRevisions = (
  blocks: readonly BlockContent[],
): BlockContent[] => {
  const out: BlockContent[] = [];
  let index = 0;
  while (index < blocks.length) {
    const block = blocks[index];
    index += 1;
    if (block === undefined) continue;
    if (block.type !== "blockCustomXml") {
      out.push(block);
      continue;
    }
    let element: BlockCustomXml = block;
    for (;;) {
      const start = markerOf(out.at(-1));
      const end = markerOf(blocks[index]);
      if (
        start?.edge !== "start" ||
        end?.edge !== "end" ||
        start.kind !== end.kind ||
        start.id !== end.id ||
        element[start.kind] !== undefined
      )
        break;
      out.pop();
      index += 1;
      const record: CustomXmlMarkupRevision = { info: start.info };
      element =
        start.kind === "insertion"
          ? Object.assign({}, element, { insertion: record })
          : Object.assign({}, element, { deletion: record });
    }
    out.push(element);
  }
  return out;
};
