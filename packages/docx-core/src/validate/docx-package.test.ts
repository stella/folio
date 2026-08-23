import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { DOCX_PACKAGE_ISSUE_CODES, validateDocxPackage } from "./docx";

const packageWithDocument = async (documentXml: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types/>");
  zip.file("_rels/.rels", "<Relationships/>");
  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};

describe("validateDocxPackage", () => {
  test("returns a structured issue for an invalid archive", async () => {
    const result = await validateDocxPackage(new Uint8Array([1, 2, 3]));

    expect(result.valid).toBe(false);
    if (result.valid) {
      throw new Error("Expected an invalid package");
    }
    expect(result.code).toBe(DOCX_PACKAGE_ISSUE_CODES.InvalidArchive);
  });

  test("returns a structured issue for a missing required part", async () => {
    const bytes = await new JSZip().generateAsync({ type: "uint8array" });
    const result = await validateDocxPackage(bytes);

    expect(result).toEqual({
      valid: false,
      code: DOCX_PACKAGE_ISSUE_CODES.MissingPackagePart,
      error: "Generated DOCX is missing required package part: [Content_Types].xml",
    });
  });

  test.each([
    [
      "Transitional namespace with an arbitrary prefix",
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ],
    ["Strict namespace", "http://purl.oclc.org/ooxml/wordprocessingml/main"],
  ])("accepts %s without optional styles or document relationships", async (_name, namespace) => {
    const bytes = await packageWithDocument(
      `<x:document xmlns:x="${namespace}"><x:body/></x:document>`,
    );

    expect(await validateDocxPackage(bytes)).toEqual({ valid: true });
  });

  test("requires a body in the same WordprocessingML namespace", async () => {
    const bytes = await packageWithDocument(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    );

    expect(await validateDocxPackage(bytes)).toEqual({
      valid: false,
      code: DOCX_PACKAGE_ISSUE_CODES.InvalidDocumentRoot,
      error: "Generated DOCX document root has no WordprocessingML body.",
    });
  });

  test("rejects a body whose familiar prefix resolves to the wrong namespace", async () => {
    const bytes = await packageWithDocument(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body xmlns:w="urn:not-wordprocessingml"/></w:document>',
    );

    expect(await validateDocxPackage(bytes)).toEqual({
      valid: false,
      code: DOCX_PACKAGE_ISSUE_CODES.InvalidDocumentRoot,
      error: "Generated DOCX document root has no WordprocessingML body.",
    });
  });

  test("rejects oversized document XML before parsing it", async () => {
    const namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const bytes = await packageWithDocument(
      `<w:document xmlns:w="${namespace}"><w:body><w:p>${"x".repeat(32 * 1024 * 1024)}</w:p></w:body></w:document>`,
    );

    const result = await validateDocxPackage(bytes);
    expect(result.valid).toBe(false);
    if (result.valid) {
      throw new Error("Expected oversized document XML to be rejected");
    }
    expect(result.code).toBe(DOCX_PACKAGE_ISSUE_CODES.ArchiveBoundsExceeded);
    expect(result.error).toContain("word/document.xml");
  });

  test("rejects malformed XML even though the parser itself is permissive", async () => {
    const bytes = await packageWithDocument(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        "<w:body></w:document>",
    );

    expect(await validateDocxPackage(bytes)).toEqual({
      valid: false,
      code: DOCX_PACKAGE_ISSUE_CODES.InvalidDocumentRoot,
      error: "Generated DOCX has malformed word/document.xml.",
    });
  });
});
