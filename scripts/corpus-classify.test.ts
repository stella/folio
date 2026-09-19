import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { classifyCorpusFile } from "./lib/corpus-classify";

const TRANSITIONAL = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const OFFICE_DOCUMENT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const relationships = (target: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${OFFICE_DOCUMENT}" Target="${target}"/></Relationships>`;

const documentXml = (namespace: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${namespace}"><w:body/></w:document>`;

type PackageParts = Record<string, string>;

const packageBytes = async (parts: PackageParts): Promise<Uint8Array> => {
  const archive = new JSZip();
  for (const [name, content] of Object.entries(parts)) {
    archive.file(name, content);
  }
  return await archive.generateAsync({ type: "uint8array" });
};

const conventionalPackage = (
  main = "word/document.xml",
  namespace = TRANSITIONAL,
): PackageParts => ({
  "[Content_Types].xml": "<Types/>",
  "_rels/.rels": relationships(main),
  [main]: documentXml(namespace),
});

describe("classifyCorpusFile", () => {
  test("accepts a conventional package", async () => {
    const result = await classifyCorpusFile(await packageBytes(conventionalPackage()));
    expect(result).toEqual({ kind: "docx", documentPart: "word/document.xml" });
  });

  test("accepts a main part that is not named word/document.xml", async () => {
    // Word Online writes `word/document2.xml`; the part is found through the
    // package relationship, never by its conventional name.
    const result = await classifyCorpusFile(
      await packageBytes(conventionalPackage("word/document2.xml")),
    );
    expect(result).toEqual({ kind: "docx", documentPart: "word/document2.xml" });
  });

  test("accepts a Strict-namespace document", async () => {
    const result = await classifyCorpusFile(
      await packageBytes(conventionalPackage("word/document.xml", STRICT)),
    );
    expect(result.kind).toBe("docx");
  });

  test("accepts a relationship target spelled with a leading slash", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": relationships("/word/document.xml"),
        "word/document.xml": documentXml(TRANSITIONAL),
      }),
    );
    expect(result).toEqual({ kind: "docx", documentPart: "word/document.xml" });
  });

  test("accepts relationships written with a namespace prefix", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": `<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships"><r:Relationship Id="rId1" Type="${OFFICE_DOCUMENT}" Target="word/document.xml"/></r:Relationships>`,
        "word/document.xml": documentXml(TRANSITIONAL),
      }),
    );
    expect(result).toEqual({ kind: "docx", documentPart: "word/document.xml" });
  });

  test("accepts the Strict spelling of the officeDocument relationship", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": `<Relationships xmlns="http://purl.oclc.org/ooxml/officeDocument/relationships"><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
        "word/document.xml": documentXml(STRICT),
      }),
    );
    expect(result).toEqual({ kind: "docx", documentPart: "word/document.xml" });
  });

  test("rejects a package with no content types", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "_rels/.rels": relationships("word/document.xml"),
        "word/document.xml": documentXml(TRANSITIONAL),
      }),
    );
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "not-an-opc-package" });
  });

  test("rejects a package with no officeDocument relationship", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": '<Relationships xmlns="x"/>',
        "word/document.xml": documentXml(TRANSITIONAL),
      }),
    );
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "not-an-opc-package" });
  });

  test("rejects a relationship that targets a missing part", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": relationships("word/gone.xml"),
      }),
    );
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "not-a-wordprocessing-package" });
  });

  test("rejects a package whose main part is not a WordprocessingML document", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": relationships("theme/themeManager.xml"),
        "theme/themeManager.xml":
          '<themeManager xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"/>',
      }),
    );
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "not-a-wordprocessing-package" });
  });

  test("rejects a document root in an unrelated namespace", async () => {
    const result = await classifyCorpusFile(
      await packageBytes(conventionalPackage("word/document.xml", "urn:invented")),
    );
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "not-a-wordprocessing-package" });
  });

  test("rejects malformed document markup", async () => {
    const result = await classifyCorpusFile(
      await packageBytes({
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": relationships("word/document.xml"),
        "word/document.xml": `<w:document xmlns:w="${TRANSITIONAL}"><w:body></w:document>`,
      }),
    );
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "malformed-document-xml" });
  });

  test("rejects bytes that are not a ZIP", async () => {
    const result = await classifyCorpusFile(new TextEncoder().encode("not a zip at all"));
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "not-a-zip" });
  });

  test("rejects a truncated archive", async () => {
    const whole = await packageBytes(conventionalPackage());
    const result = await classifyCorpusFile(whole.slice(0, Math.floor(whole.length / 2)));
    expect(result).toMatchObject({ kind: "not-a-docx", reason: "unreadable-archive" });
  });

  test("separates an encrypted package from any other compound file", async () => {
    const signature = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
    const plain = new Uint8Array(512);
    plain.set(signature);
    expect(await classifyCorpusFile(plain)).toMatchObject({ reason: "ole-compound-file" });

    const encrypted = new Uint8Array(512);
    encrypted.set(signature);
    encrypted.set(Buffer.from("EncryptedPackage", "utf16le"), 64);
    expect(await classifyCorpusFile(encrypted)).toMatchObject({ reason: "encrypted-package" });
  });
});
