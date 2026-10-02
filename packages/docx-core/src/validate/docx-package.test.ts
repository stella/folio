import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { DOCX_PACKAGE_ISSUE_CODES, validateDocxPackage } from "./docx";

const packageWithDocument = async (documentXml: string): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
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

  test("saved-package oracle detects a timestamp-sized comment id mutation", async () => {
    const namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const valid = await packageWithDocument(
      `<w:document xmlns:w="${namespace}"><w:body><w:p/></w:body></w:document>`,
    );
    const zip = await JSZip.loadAsync(valid);
    const declarations = await zip.file("[Content_Types].xml")?.async("string");
    if (!declarations) throw new TypeError("Missing mutation fixture declarations");
    zip.file(
      "[Content_Types].xml",
      declarations.replace(
        "</Types>",
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
      ),
    );
    zip.file(
      "word/comments.xml",
      `<w:comments xmlns:w="${namespace}"><w:comment w:id="1" w:author="Reviewer"><w:p/></w:comment></w:comments>`,
    );
    zip.file(
      "word/_rels/document.xml.rels",
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
    );
    const bytes = await zip.generateAsync({ type: "uint8array" });
    expect(await validateDocxPackage(bytes)).toEqual({ valid: true });
    zip.file(
      "word/comments.xml",
      `<w:comments xmlns:w="${namespace}"><w:comment w:id="1790870400000" w:author="Reviewer"><w:p/></w:comment></w:comments>`,
    );
    const mutated = await validateDocxPackage(await zip.generateAsync({ type: "uint8array" }));
    expect(mutated.valid).toBe(false);
    if (mutated.valid) throw new TypeError("Comment mutation escaped the validity oracle");
    expect(mutated.code).toBe(DOCX_PACKAGE_ISSUE_CODES.InvalidSchemaAttribute);
    expect(mutated.error).toContain("word/comments.xml");
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

const CENTRAL_DIRECTORY_SIGNATURE = 0x02_01_4b_50;
const LOCAL_HEADER_SIGNATURE = 0x04_03_4b_50;

/**
 * Rewrite the uncompressed size `entryPath` declares in its local header and
 * central-directory record, leaving its compressed data untouched.
 */
const withDeclaredSize = (source: Uint8Array, entryPath: string, declaredBytes: number) => {
  const bytes = source.slice();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const name = new TextEncoder().encode(entryPath);
  const nameAt = (offset: number): boolean =>
    offset + name.length <= bytes.length &&
    name.every((byte, index) => bytes[offset + index] === byte);
  let patched = 0;
  for (let offset = 0; offset + 46 <= bytes.length; offset += 1) {
    const signature = view.getUint32(offset, true);
    if (
      signature === CENTRAL_DIRECTORY_SIGNATURE &&
      view.getUint16(offset + 28, true) === name.length &&
      nameAt(offset + 46)
    ) {
      view.setUint32(offset + 24, declaredBytes, true);
      patched += 1;
    } else if (
      signature === LOCAL_HEADER_SIGNATURE &&
      view.getUint16(offset + 26, true) === name.length &&
      nameAt(offset + 30)
    ) {
      view.setUint32(offset + 22, declaredBytes, true);
      patched += 1;
    }
  }
  if (patched !== 2) {
    throw new Error(`Expected one local and one central record for ${entryPath}`);
  }
  return bytes;
};

const documentOfSize = (bytes: number) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${"x".repeat(bytes)}</w:p></w:body></w:document>`;

describe("validateDocxPackage archive inflation limits", () => {
  test("stops document XML that inflates past its declared size", async () => {
    const bytes = withDeclaredSize(
      await packageWithDocument(documentOfSize(3 * 1024 * 1024)),
      "word/document.xml",
      64,
    );

    expect(await validateDocxPackage(bytes)).toEqual({
      valid: false,
      code: DOCX_PACKAGE_ISSUE_CODES.ArchiveBoundsExceeded,
      error: 'Generated DOCX entry "word/document.xml" inflates past its declared size.',
    });
  });

  test("refuses an entry whose declared expansion passes the ratio cap", async () => {
    const bytes = await packageWithDocument(documentOfSize(5 * 1024 * 1024));

    expect(await validateDocxPackage(bytes)).toEqual({
      valid: false,
      code: DOCX_PACKAGE_ISSUE_CODES.ArchiveBoundsExceeded,
      error:
        'Generated DOCX entry "word/document.xml" declares more than 200 uncompressed bytes per compressed byte.',
    });
  });

  test("leaves a highly compressible binary entry to the byte and archive limits", async () => {
    const zip = await JSZip.loadAsync(await packageWithDocument(documentOfSize(16)));
    const declarations = await zip.file("[Content_Types].xml")?.async("string");
    if (!declarations) throw new TypeError("Missing binary fixture declarations");
    zip.file(
      "[Content_Types].xml",
      declarations.replace(
        "</Types>",
        '<Default Extension="bmp" ContentType="image/bmp"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>',
      ),
    );
    zip.file("word/media/image1.bmp", new Uint8Array(5 * 1024 * 1024).fill(0xff));
    // Incompressible padding keeps the archive as a whole under the ratio cap.
    zip.file("padding.bin", crypto.getRandomValues(new Uint8Array(64 * 1024)), {
      compression: "STORE",
    });
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });

    expect(await validateDocxPackage(bytes)).toEqual({ valid: true });
  });

  test("counts archive records before the archive is parsed", async () => {
    const zip = new JSZip();
    for (let index = 0; index < 4100; index += 1) {
      zip.file(`customXml/item${String(index)}.xml`, "", { createFolders: false });
    }
    const bytes = await zip.generateAsync({ type: "uint8array" });

    expect(await validateDocxPackage(bytes)).toEqual({
      valid: false,
      code: DOCX_PACKAGE_ISSUE_CODES.ArchiveBoundsExceeded,
      error: "Generated DOCX holds more than 4096 archive entries.",
    });
  });
});
