import {expect, test} from "bun:test";
import JSZip from "jszip";

import {createEmptyDocument} from "../utils/createDocument";
import {removeResolvedHeaderFooterParts} from "./removeHeaderFooterParts";
import {parseRelationships, RELATIONSHIP_TYPES} from "./relsParser";

const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

test.each(["retired", "shared"] as const)("resolved references remove only exclusively owned parts (%s)", async (ownership) => {
  const zip = new JSZip();
  zip.file("word/header1.xml", "header");
  zip.file("word/_rels/header1.xml.rels", `<Relationships xmlns="${RELS}"/>`);
  zip.file("word/_rels/document.xml.rels", `<pkg:Relationships xmlns:pkg="${RELS}"><pkg:Relationship Id="rId_header" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/><pkg:Relationship Id="rId_styles" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/></pkg:Relationships>`);
  zip.file("[Content_Types].xml", `<ct:Types xmlns:ct="${TYPES}"><ct:Override PartName="/word/header1.xml" ContentType="header"/><ct:Override PartName="/word/styles.xml" ContentType="styles"/></ct:Types>`);
  if (ownership === "shared") {
    zip.file("customXml/_rels/item1.xml.rels", `<Relationships xmlns="${RELS}"><Relationship Id="shared" Type="urn:example:shared" Target="../word/header1.xml"/></Relationships>`);
  }
  const document = createEmptyDocument();
  await removeResolvedHeaderFooterParts({document, zip, removedReferences: [{part: "header", type: "default", relationshipId: "rId_header"}], compressionLevel: 6});
  const relsFile = zip.file("word/_rels/document.xml.rels");
  const typesFile = zip.file("[Content_Types].xml");
  if (!relsFile || !typesFile) throw new Error("Expected package metadata");
  expect([...parseRelationships(await relsFile.async("text")).keys()]).toEqual(["rId_styles"]);
  expect(zip.file("word/header1.xml") !== null).toBe(ownership === "shared");
  expect(zip.file("word/_rels/header1.xml.rels") !== null).toBe(ownership === "shared");
  const types = await typesFile.async("text");
  expect(types.includes('/word/header1.xml')).toBe(ownership === "shared");
  expect(types).toContain('/word/styles.xml');
});
