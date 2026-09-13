import { expect, test } from "bun:test";
import { patchNumberingDefinitions } from "./selectiveXmlPatch";
import {
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const baseline = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const current = baseline
  .replace("singleLevel", "multilevel")
  .replace(
    "</w:numbering>",
    '<w:abstractNum w:abstractNumId="2"/><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>',
  );

for (const prefix of ["w", "n", ""] as const) {
  test(`numbering changes preserve bindings and definition order with prefix '${prefix}'`, () => {
    const original =
      prefix === ""
        ? baseline
            .replaceAll("w:", "")
            .replace("xmlns:w=", "xmlns=")
            .replaceAll('abstractNumId="', 'a:abstractNumId="')
            .replaceAll('numId="', 'a:numId="')
            .replaceAll('val="', 'a:val="')
            .replace("<numbering ", `<numbering xmlns:a="${W}" `)
        : baseline.replaceAll("w:", `${prefix}:`).replace("xmlns:w=", `xmlns:${prefix}=`);
    const patched = patchNumberingDefinitions({
      originalXml: original,
      baselineXml: baseline,
      currentXml: current,
    });
    expect(patched).not.toBeNull();
    const root = parseXmlDocument(patched ?? "");
    if (!root) throw new Error("Expected numbering root");
    const definitions = getChildElements(root);
    expect(definitions.map((child) => getNamespaceUri(child))).toEqual([W, W, W, W]);
    expect(definitions.map((child) => getLocalName(child.name))).toEqual([
      "abstractNum",
      "abstractNum",
      "num",
      "num",
    ]);
    expect(
      definitions.map((child) =>
        getAttributeByNamespaceUri(
          child,
          WORDPROCESSINGML_NAMESPACE_URIS,
          getLocalName(child.name) === "num" ? "numId" : "abstractNumId",
        ),
      ),
    ).toEqual(["0", "2", "1", "3"]);
    expect(patched).toContain('val="multilevel"');
  });
}
