/** Preserve producer settings while updating the document-wide header/footer switch. */
import { spliceXml } from "./selectiveXmlPatch";
import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

const XML_TOKEN =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;

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
      if (children[childIndex]?.name !== name) return null;
      childStart = token.index;
      if (tag.endsWith("/>")) {
        const child = children[childIndex++];
        if (!child) return null;
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
  spans.push({ start: rootEnd, end: rootEnd, newXml: flag });
  return spliceXml(xml, spans);
};
