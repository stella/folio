/** Preserve producer settings while updating the document-wide header/footer switch. */
import { panic } from "better-result";

import { spliceXml } from "./selectiveXmlPatch";
import { SETTINGS_CHILDREN } from "@stll/docx-core/schema";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

const XML_TOKEN =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;
const SETTINGS_CHILD_ORDINALS: ReadonlyMap<string, number> = new Map(
  SETTINGS_CHILDREN.map((name, index) => [name, index] as const),
);
const EVEN_AND_ODD_HEADERS_ORDER = SETTINGS_CHILD_ORDINALS.get("evenAndOddHeaders");

/** Null means the source is malformed or outside the supported namespace profiles. */
export const updateEvenAndOddHeaders = (
  xml: string,
  enabled: boolean | undefined,
): string | null => {
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
  if (EVEN_AND_ODD_HEADERS_ORDER === undefined) {
    panic("Generated CT_Settings schema does not define evenAndOddHeaders");
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
          getLocalName(child.name) === "evenAndOddHeaders" &&
          WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "")
        ) {
          spans.push({ start: childStart, end: token.index + tag.length, newXml: "" });
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
      if (
        firstFollowingChildStart < 0 &&
        WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "") &&
        (SETTINGS_CHILD_ORDINALS.get(getLocalName(child.name)) ?? -1) > EVEN_AND_ODD_HEADERS_ORDER
      ) {
        firstFollowingChildStart = token.index;
      }
      childStart = token.index;
      if (tag.endsWith("/>")) {
        childIndex += 1;
        if (
          getLocalName(child.name) === "evenAndOddHeaders" &&
          WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(child) ?? "")
        ) {
          spans.push({ start: token.index, end: token.index + tag.length, newXml: "" });
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
      : `<folioW:evenAndOddHeaders xmlns:folioW="${getNamespaceUri(root)}"${enabled ? "" : ' folioW:val="0"'}/>`;
  if (selfClosingRootStart >= 0) {
    if (!flag) return xml;
    return spliceXml(xml, [
      { start: rootOpenEnd - 2, end: rootOpenEnd, newXml: `>${flag}</${root.name}>` },
    ]);
  }
  if (rootEnd < 0) return null;
  const existingFlag = spans.at(0);
  if (flag && existingFlag) {
    existingFlag.newXml = flag;
  } else if (flag) {
    const insertAt = firstFollowingChildStart >= 0 ? firstFollowingChildStart : rootEnd;
    spans.push({ start: insertAt, end: insertAt, newXml: flag });
  }
  return spliceXml(xml, spans);
};
