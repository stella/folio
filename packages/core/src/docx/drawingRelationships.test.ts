import { expect, test } from "bun:test";

import { rebindDrawingImageRelationship } from "./drawingRelationships";
import { findAttributeByNamespaceUri, getChildElements, parseXml } from "./xmlParser";

const RELATIONSHIP_NAMESPACES = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
];

test("rebinds the embedded picture by namespace while preserving unrelated attributes", () => {
  for (const namespace of RELATIONSHIP_NAMESPACES) {
    const result = rebindDrawingImageRelationship({
      xml: `<x:blip xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:resource="${namespace}" resource:embed="source" cstate="print"/>`,
      previousId: "source",
      nextId: "destination",
    });
    expect(result).not.toBeNull();
    const root = getChildElements(parseXml(result ?? "")).at(0);
    expect(
      findAttributeByNamespaceUri(root, new Set(RELATIONSHIP_NAMESPACES), "embed")?.value,
    ).toBe("destination");
    expect(root?.attributes?.["cstate"]).toBe("print");
    expect(root?.attributes?.["xmlns:resource"]).toBe(RELATIONSHIP_NAMESPACES.at(0));
  }
});

test("refuses drawings with additional package dependencies", () => {
  for (const dependencies of [
    '<a:blip r:embed="source"/><a:blip r:embed="other"/>',
    '<a:blip r:embed="source"/><a:hlinkClick r:id="hyperlink"/>',
    '<a:blip r:link="source"/>',
  ]) {
    expect(
      rebindDrawingImageRelationship({
        xml: `<w:drawing>${dependencies}</w:drawing>`,
        previousId: "source",
        nextId: "destination",
      }),
    ).toBeNull();
  }
});

test("rebinds VML image references without allowing arbitrary relationship IDs", () => {
  const xml =
    '<v:imagedata xmlns:v="urn:schemas-microsoft-com:vml" xmlns:res="http://schemas.openxmlformats.org/officeDocument/2006/relationships" res:id="source"/>';
  const result = rebindDrawingImageRelationship({ xml, previousId: "source", nextId: "target" });
  expect(result).toContain('res:id="target"');
  expect(
    rebindDrawingImageRelationship({
      xml: xml.replace("urn:schemas-microsoft-com:vml", "urn:foreign"),
      previousId: "source",
      nextId: "target",
    }),
  ).toBeNull();
});
