import {
  type PreservedChild,
  preservedChildGap,
} from "../model/preservedMarkup";
import { readXml } from "./streamingXmlParser";
import {
  OOXML_NAMESPACE_SCOPE,
  OOXML_NS,
  resolveNamespaceUri,
} from "./xmlNamespaceContext";

const isWordprocessing = (uri: string | undefined): boolean =>
  uri === OOXML_NS.w ||
  uri === "http://purl.oclc.org/ooxml/wordprocessingml/main";

export type PreservedCommentRangeMarker = {
  type: "commentRangeStart" | "commentRangeEnd";
  id: number;
};

/** Read a range boundary captured as a direct child of a table or row. */
export const preservedCommentRangeMarker = (
  child: Pick<PreservedChild, "xml">,
): PreservedCommentRangeMarker | undefined => {
  const marker = capturedMarker(child);
  return marker && { type: marker.type, id: marker.id };
};

/** A captured range boundary under another comment id; another child as it is. */
export const withPreservedCommentRangeId = (
  child: PreservedChild,
  id: number,
): PreservedChild => {
  const marker = capturedMarker(child);
  if (marker === undefined || marker.id === id) return child;
  const attribute = marker.attribute.replaceAll(".", String.raw`\.`);
  return {
    ...child,
    xml: child.xml.replace(
      new RegExp(
        String.raw`(?<before>\s${attribute}\s*=\s*)(?<quote>["'])\d+\k<quote>`,
        "u",
      ),
      `$<before>$<quote>${String(id)}$<quote>`,
    ),
  };
};

const capturedMarker = (
  child: Pick<PreservedChild, "xml">,
): (PreservedCommentRangeMarker & { attribute: string }) | undefined => {
  if (!child.xml.includes("commentRange")) return undefined;
  const parsed = readXml(child.xml, {
    conformance: "ooxml",
    inheritedNamespaceScope: OOXML_NAMESPACE_SCOPE,
  });
  if (parsed.status !== "parsed") return undefined;
  const element = parsed.value.elements?.find(({ type }) => type === "element");
  if (element?.name === undefined || !isWordprocessing(element.namespaceUri))
    return undefined;
  const type = element.name.slice(element.name.indexOf(":") + 1);
  if (type !== "commentRangeStart" && type !== "commentRangeEnd")
    return undefined;
  for (const [name, raw] of Object.entries(element.attributes ?? {})) {
    const colon = name.indexOf(":");
    if (
      colon < 0 ||
      name.slice(colon + 1) !== "id" ||
      resolveNamespaceUri(element.namespaceScope, name.slice(0, colon)) !==
        element.namespaceUri ||
      typeof raw !== "string" ||
      !/^\d+$/u.test(raw)
    )
      continue;
    const id = Number(raw);
    if (Number.isSafeInteger(id)) return { type, id, attribute: name };
  }
  return undefined;
};

/** Read the boundaries before the named row or cell, in authored order. */
export const preservedCommentRangeMarkersAt = (
  children: readonly PreservedChild[] | undefined,
  { at, childCount }: { at: number; childCount: number },
): PreservedCommentRangeMarker[] => {
  const markers: PreservedCommentRangeMarker[] = [];
  for (const child of children ?? []) {
    if (preservedChildGap(child, childCount) !== at) continue;
    const marker = preservedCommentRangeMarker(child);
    if (marker !== undefined) markers.push(marker);
  }
  return markers;
};
