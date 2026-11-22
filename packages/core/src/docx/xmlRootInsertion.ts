import { panic } from "better-result";

import { spliceXml } from "./selectiveXmlPatch";

/** Markup boundaries in well-formed XML, excluding quoted values and comments. */
const XML_TOKENS =
  /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;

/**
 * `xml` with `markup` as the last child of its root element, whose qualified
 * name is `rootName`. The authored bytes around it stay as they are.
 */
export const appendToXmlRoot = (
  xml: string,
  { rootName, markup }: { rootName: string; markup: string },
): string => {
  for (const { 0: token, index } of xml.matchAll(XML_TOKENS)) {
    if (token.startsWith("</") && token.slice(2, -1).trim() === rootName)
      return (
        spliceXml(xml, [{ start: index, end: index, newXml: markup }]) ??
        panic("A root insertion must preserve XML ranges.")
      );
    if (
      token.startsWith(`<${rootName}`) &&
      /[\s/>]/u.test(token.charAt(rootName.length + 1)) &&
      token.endsWith("/>")
    ) {
      const end = index + token.length - 2;
      return (
        spliceXml(xml, [
          { start: end, end: end + 2, newXml: `>${markup}</${rootName}>` },
        ]) ?? panic("A root insertion must preserve XML ranges.")
      );
    }
  }
  return panic("A root insertion needs a closing root boundary.");
};
