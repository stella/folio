import { expect, test } from "bun:test";
import { validatePackageGraph } from "./packageGraph";

const contentTypes =
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
const rootRels =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
const fixture = () =>
  new Map([
    ["[Content_Types].xml", contentTypes],
    ["_rels/.rels", rootRels],
    [
      "word/document.xml",
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
    ],
  ]);

test("OPC oracle accepts complete declarations and detects each graph mutation", () => {
  const valid = fixture();
  expect(validatePackageGraph(valid, new Set(valid.keys()))).toBeNull();
  const mutations = [
    (parts: Map<string, string>) => parts.delete("_rels/.rels"),
    (parts: Map<string, string>) =>
      parts.set("_rels/.rels", rootRels.replace("word/document.xml", "word/missing.xml")),
    (parts: Map<string, string>) =>
      parts.set("_rels/.rels", rootRels.replace('Target="', 'TargetMode="invalid" Target="')),
    (parts: Map<string, string>) =>
      parts.set(
        "[Content_Types].xml",
        contentTypes.replace(
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          "",
        ),
      ),
    (parts: Map<string, string>) =>
      parts.set(
        "[Content_Types].xml",
        contentTypes.replace("/word/document.xml", "/word/missing.xml"),
      ),
  ];
  for (const mutate of mutations) {
    const parts = fixture();
    mutate(parts);
    expect(validatePackageGraph(parts, new Set(parts.keys()))).not.toBeNull();
  }
});

test("OPC oracle resolves targets relative to the source part", () => {
  const parts = fixture();
  parts.set(
    "word/_rels/document.xml.rels",
    '<p:Relationships xmlns:p="http://schemas.openxmlformats.org/package/2006/relationships"><p:Relationship Id="x" Type="urn:image" Target="media/image.xml"/></p:Relationships>',
  );
  parts.set("word/media/image.xml", "<image/>");
  expect(validatePackageGraph(parts, new Set(parts.keys()))).toBeNull();
  parts.delete("word/media/image.xml");
  expect(validatePackageGraph(parts, new Set(parts.keys()))).toContain(
    "Missing relationship target",
  );
});

test("OPC defaults use only final-segment extensions; extensionless parts need overrides", () => {
  for (const directory of ["", "customXml/", "customXml/item.d/"]) {
    for (const name of ["data", "data.XML"]) {
      const path = directory + name;
      const parts = fixture();
      parts.set(path, "<data/>");
      // A Default matching the extensionless filename must never cover that part.
      const declarations = contentTypes.replace(
        "</Types>",
        '<Default Extension="data" ContentType="application/xml"/></Types>',
      );
      parts.set("[Content_Types].xml", declarations);
      expect(validatePackageGraph(parts, new Set(parts.keys()))).toBe(
        name === "data" ? `Missing content type for ${path}` : null,
      );
      parts.set(
        "[Content_Types].xml",
        declarations.replace(
          "</Types>",
          `<Override PartName="/${path}" ContentType="application/xml"/></Types>`,
        ),
      );
      expect(validatePackageGraph(parts, new Set(parts.keys()))).toBeNull();
    }
  }
});
