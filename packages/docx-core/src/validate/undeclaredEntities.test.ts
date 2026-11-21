/**
 * The validators pass what the engines open and refuse what they refuse: an
 * entity no declaration defines fails every check, while the character
 * references the engines read pass, whatever character they name.
 */

import { describe, expect, test } from "bun:test";

import {
  undeclaredXmlEntityError,
  wellFormedXmlIssue,
} from "../xml/streamingXmlParser";
import { ZipArchive } from "../zip/archive";
import { validateDocxPackage } from "./docx";
import { validatePackageGraph } from "./packageGraph";
import { schemaAttributeIssues } from "./schemaAttributes";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const contentTypes = (extra = "") =>
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"${extra}><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const rootRels =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
const documentWith = (text: string) =>
  `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;

const packageOf = async (documentXml: string) => {
  const zip = new ZipArchive();
  zip.set("[Content_Types].xml", contentTypes());
  zip.set("_rels/.rels", rootRels);
  zip.set("word/document.xml", documentXml);
  return zip.generate();
};

const UNDECLARED = ["&nbsp;", "&copy;", "&amp;nbsp;&ndash;"] as const;
const READ = ["&#xFFFE;", "&#65;", "&#x1F642;", "&amp;"] as const;

describe("an entity no declaration defines", () => {
  test.each(UNDECLARED)("%s fails every validator", async (text) => {
    const xml = documentWith(text);
    // The engines refuse each of these parts.
    expect(undeclaredXmlEntityError("word/document.xml", xml)).not.toBeNull();
    expect(wellFormedXmlIssue(xml, "syntactic")).not.toBeNull();
    expect(schemaAttributeIssues(xml)).not.toEqual([]);
    expect((await validateDocxPackage(await packageOf(xml))).valid).toBe(false);
    const parts = new Map([
      ["[Content_Types].xml", contentTypes(` title="${text}"`)],
      ["_rels/.rels", rootRels],
      ["word/document.xml", documentWith("")],
    ]);
    expect(validatePackageGraph(parts, new Set(parts.keys()))).not.toBeNull();
  });
});

describe("a character reference the engines read", () => {
  test.each(READ)("%s passes every validator", async (text) => {
    const xml = documentWith(text);
    expect(undeclaredXmlEntityError("word/document.xml", xml)).toBeNull();
    expect(wellFormedXmlIssue(xml, "syntactic")).toBeNull();
    expect(schemaAttributeIssues(xml)).toEqual([]);
    expect((await validateDocxPackage(await packageOf(xml))).valid).toBe(true);
    const parts = new Map([
      ["[Content_Types].xml", contentTypes(` title="${text}"`)],
      ["_rels/.rels", rootRels],
      ["word/document.xml", documentWith("")],
    ]);
    expect(validatePackageGraph(parts, new Set(parts.keys()))).toBeNull();
  });
});
