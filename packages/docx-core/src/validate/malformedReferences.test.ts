/**
 * The validators pass what the engines open and refuse what they refuse: a
 * reference XML refuses (an entity no declaration defines, a character XML
 * 1.0 does not allow) fails every check, and so does the engines' refusal.
 */

import { describe, expect, test } from "bun:test";

import {
  malformedXmlReferenceError,
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

const graphWith = (text: string) => {
  const parts = new Map([
    ["[Content_Types].xml", contentTypes(` title="${text}"`)],
    ["_rels/.rels", rootRels],
    ["word/document.xml", documentWith("")],
  ]);
  return validatePackageGraph(parts, new Set(parts.keys()));
};

/** A text, the reference in it XML refuses, and why. */
const MALFORMED = [
  ["&nbsp;", "&nbsp;", "undeclaredEntity"],
  ["&copy;", "&copy;", "undeclaredEntity"],
  ["&amp;nbsp;&ndash;", "&ndash;", "undeclaredEntity"],
  ["&#0;", "&#0;", "invalidCharacterReference"],
  ["&#x1;", "&#x1;", "invalidCharacterReference"],
  ["&#31;", "&#31;", "invalidCharacterReference"],
  ["&#xD800;", "&#xD800;", "invalidCharacterReference"],
  ["&#xFFFE;", "&#xFFFE;", "invalidCharacterReference"],
  ["&#x110000;", "&#x110000;", "invalidCharacterReference"],
  ["&#12ab;", "&#12ab;", "invalidCharacterReference"],
  ["&#X41;", "&#X41;", "invalidCharacterReference"],
] as const;
const ALLOWED = [
  "&#65;",
  "&#x9;",
  "&#x1F642;",
  "&#xfffd;",
  "&#xFFFD;",
  "&amp;",
] as const;

describe("a reference XML refuses", () => {
  test.each(MALFORMED)(
    "%s fails every validator",
    async (text, refused, kind) => {
      const xml = documentWith(text);
      expect(
        malformedXmlReferenceError("word/document.xml", xml),
      ).toMatchObject({
        kind,
        offset: xml.indexOf(refused),
        reference: refused.slice(1, -1),
      });
      expect(wellFormedXmlIssue(xml)).not.toBeNull();
      expect(schemaAttributeIssues(xml)).not.toEqual([]);
      expect((await validateDocxPackage(await packageOf(xml))).valid).toBe(
        false,
      );
      expect(graphWith(text)).not.toBeNull();
    },
  );
});

describe("a reference XML allows", () => {
  test.each(ALLOWED)("%s passes every validator", async (text) => {
    const xml = documentWith(text);
    expect(malformedXmlReferenceError("word/document.xml", xml)).toBeNull();
    expect(wellFormedXmlIssue(xml)).toBeNull();
    expect(schemaAttributeIssues(xml)).toEqual([]);
    expect((await validateDocxPackage(await packageOf(xml))).valid).toBe(true);
    expect(graphWith(text)).toBeNull();
  });
});
