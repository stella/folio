/** Register a package part without rewriting authored content-type entries. */
import { OOXML_NS } from "@stll/docx-utils";
import { escapeXmlAttribute } from "@stll/docx-core";
import { panic } from "better-result";
import { spliceXml } from "./selectiveXmlPatch";
import { resolvePackageRelationshipTarget } from "./packageParts";
import {
  getAttribute,
  getLocalName,
  getNamespacePrefix,
  getNamespaceUri,
  parseXmlDocument,
} from "./xmlParser";

type ContentTypePart = { partName: string; contentType: string };

export const registerContentTypeParts = (
  xml: string,
  parts: readonly ContentTypePart[],
): string => {
  const root = parseXmlDocument(xml);
  if (
    root?.name === undefined ||
    getLocalName(root.name) !== "Types" ||
    getNamespaceUri(root) !== OOXML_NS.ct
  )
    return panic("The package content-type registry must have a valid Types root.");
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  for (const element of root.elements ?? []) {
    if (getNamespaceUri(element) !== OOXML_NS.ct) continue;
    const type = getAttribute(element, null, "ContentType");
    if (type === null) continue;
    if (getLocalName(element.name) === "Override") {
      const name = getAttribute(element, null, "PartName");
      const path =
        name === null ? undefined : resolvePackageRelationshipTarget(name, "_rels/.rels");
      if (path !== undefined) overrides.set(path, type);
    }
    if (getLocalName(element.name) === "Default") {
      const extension = getAttribute(element, null, "Extension");
      if (extension !== null) defaults.set(extension.toLowerCase(), type);
    }
  }
  const prefix = getNamespacePrefix(root.name);
  const name = prefix ? `${prefix}:Override` : "Override";
  const additions: string[] = [];
  for (const part of parts) {
    const path = resolvePackageRelationshipTarget(part.partName, "_rels/.rels");
    if (path === undefined) return panic("A registered part must have a safe package URI.");
    const type = overrides.get(path);
    if (type !== undefined) {
      if (type !== part.contentType)
        return panic("A part's registered content type cannot change.");
      continue;
    }
    const extension = path.slice(path.lastIndexOf(".") + 1);
    if (defaults.get(extension) === part.contentType) continue;
    additions.push(
      `<${name} PartName="${escapeXmlAttribute(part.partName)}" ContentType="${escapeXmlAttribute(part.contentType)}"/>`,
    );
    overrides.set(path, part.contentType);
  }
  if (additions.length === 0) return xml;
  // Locate markup boundaries in validated XML, excluding quoted values and comments.
  const tokens =
    /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;
  for (const token of xml.matchAll(tokens)) {
    const markup = token[0];
    if (markup.startsWith("</") && markup.slice(2, -1).trim() === root.name)
      return (
        spliceXml(xml, [{ start: token.index, end: token.index, newXml: additions.join("") }]) ??
        panic("A registry insertion must preserve XML ranges.")
      );
    if (
      markup.startsWith(`<${root.name}`) &&
      /[\s/>]/u.test(markup.charAt(root.name.length + 1)) &&
      markup.endsWith("/>")
    ) {
      const end = token.index + markup.length - 2;
      return (
        spliceXml(xml, [
          { start: end, end: end + 2, newXml: ">" + additions.join("") + `</${root.name}>` },
        ]) ?? panic("A registry insertion must preserve XML ranges.")
      );
    }
  }
  return panic("A content-type registry needs a closing root boundary.");
};
