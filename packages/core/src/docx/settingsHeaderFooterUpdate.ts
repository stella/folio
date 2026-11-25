/** Preserve producer settings while updating the document-wide header/footer switch. */
import { panic } from "better-result";

import { SETTINGS_CHILDREN } from "@stll/docx-core/schema";

import { spliceXml } from "./selectiveXmlPatch";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  NAMESPACES,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

const XML_TOKEN =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;
const SETTINGS_CHILD_ORDINALS: ReadonlyMap<string, number> = new Map(
  SETTINGS_CHILDREN.map((name, index) => [name, index] as const),
);

/** The schema ordinals of the settings a compatibility wrapper may stand for. */
const wrappedOrdinals = (wrapper: XmlElement): number[] => {
  const ordinals: number[] = [];
  for (const branch of getChildElements(wrapper)) {
    for (const child of getChildElements(branch)) {
      const ordinal = WORDPROCESSINGML_NAMESPACE_URIS.has(
        getNamespaceUri(child) ?? "",
      )
        ? SETTINGS_CHILD_ORDINALS.get(getLocalName(child.name))
        : undefined;
      if (ordinal !== undefined) ordinals.push(ordinal);
    }
  }
  return ordinals;
};

/**
 * Where a child of `w:settings` sits relative to the flag at `childOrder`.
 * A child of another namespace (the math properties, the schema library,
 * later extensions) follows every on/off child. A compatibility wrapper sits
 * where the settings it stands for sit; one that holds the flag itself, or
 * settings on both sides of it, leaves the flag no place.
 */
const placeOf = (
  child: XmlElement,
  childOrder: number,
): "preceding" | "following" | "unplaceable" => {
  const namespaceUri = getNamespaceUri(child) ?? "";
  if (WORDPROCESSINGML_NAMESPACE_URIS.has(namespaceUri)) {
    return (SETTINGS_CHILD_ORDINALS.get(getLocalName(child.name)) ?? -1) >
      childOrder
      ? "following"
      : "preceding";
  }
  if (namespaceUri !== NAMESPACES.mc) return "following";
  const ordinals = wrappedOrdinals(child);
  const before = ordinals.some((ordinal) => ordinal < childOrder);
  const after = ordinals.some((ordinal) => ordinal > childOrder);
  if (ordinals.includes(childOrder) || (before && after)) return "unplaceable";
  return after ? "following" : "preceding";
};

/** The on/off children of `w:settings` a caller may set without re-serializing the part. */
export type SettingsOnOffChild = "evenAndOddHeaders" | "updateFields";

/**
 * Set (`true`), clear (`false`) or remove (`undefined`) one on/off child of
 * `w:settings`, at its place in the schema's child order; the rest of the part
 * keeps its bytes. Null means the source is malformed or outside the supported
 * namespace profiles.
 */
export const updateSettingsOnOffChild = (
  xml: string,
  localName: SettingsOnOffChild,
  enabled: boolean | undefined,
): string | null => {
  const childOrder = SETTINGS_CHILD_ORDINALS.get(localName);
  const root = parseXmlDocument(xml);
  if (
    !root ||
    getLocalName(root.name) !== "settings" ||
    !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(root) ?? "")
  )
    return null;
  const children = getChildElements(root);
  const spans: { start: number; end: number; newXml: string }[] = [];
  const stack: string[] = [];
  let childIndex = 0;
  let childStart = -1;
  let rootEnd = -1;
  let rootOpenEnd = -1;
  let selfClosingRootStart = -1;
  let firstFollowingChildStart = -1;
  if (childOrder === undefined) {
    return panic(`Generated CT_Settings schema does not define ${localName}`);
  }
  for (const token of xml.matchAll(XML_TOKEN)) {
    const tag = token[0];
    if (tag.startsWith("<!") || tag.startsWith("<?")) continue;
    const closing = tag.startsWith("</");
    const name = tag
      .slice(closing ? 2 : 1)
      .split(/[\s/>]/u)
      .at(0);
    if (!name) return null;
    if (closing) {
      if (stack.pop() !== name) return null;
      if (stack.length === 1) {
        const child = children[childIndex++];
        if (!child || childStart < 0) return null;
        if (
          getLocalName(child.name) === localName &&
          WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "")
        ) {
          spans.push({
            start: childStart,
            end: token.index + tag.length,
            newXml: "",
          });
        }
        childStart = -1;
      }
      if (stack.length === 0) rootEnd = token.index;
      continue;
    }
    if (stack.length === 0) {
      rootOpenEnd = token.index + tag.length;
      if (tag.endsWith("/>")) selfClosingRootStart = token.index;
    }
    if (stack.length === 1) {
      const child = children[childIndex];
      if (child?.name !== name) return null;
      const place = placeOf(child, childOrder);
      if (place === "unplaceable") return null;
      if (firstFollowingChildStart < 0 && place === "following") {
        firstFollowingChildStart = token.index;
      }
      childStart = token.index;
      if (tag.endsWith("/>")) {
        childIndex += 1;
        if (
          getLocalName(child.name) === localName &&
          WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "")
        ) {
          spans.push({
            start: token.index,
            end: token.index + tag.length,
            newXml: "",
          });
        }
        childStart = -1;
      }
    }
    if (!tag.endsWith("/>")) stack.push(name);
  }
  if (stack.length !== 0 || childIndex !== children.length) return null;
  // Bind the new flag locally; attributes need a prefix even under a default namespace.
  const flag =
    enabled === undefined
      ? ""
      : `<folioW:${localName} xmlns:folioW="${getNamespaceUri(root)}"${enabled ? "" : ' folioW:val="0"'}/>`;
  if (selfClosingRootStart >= 0) {
    if (!flag) return xml;
    return spliceXml(xml, [
      {
        start: rootOpenEnd - 2,
        end: rootOpenEnd,
        newXml: `>${flag}</${root.name}>`,
      },
    ]);
  }
  if (rootEnd < 0) return null;
  const existingFlag = spans.at(0);
  if (flag && existingFlag) {
    existingFlag.newXml = flag;
  } else if (flag) {
    const insertAt =
      firstFollowingChildStart >= 0 ? firstFollowingChildStart : rootEnd;
    spans.push({ start: insertAt, end: insertAt, newXml: flag });
  }
  return spliceXml(xml, spans);
};

/** Null means the source is malformed or outside the supported namespace profiles. */
export const updateEvenAndOddHeaders = (
  xml: string,
  enabled: boolean | undefined,
): string | null => updateSettingsOnOffChild(xml, "evenAndOddHeaders", enabled);
