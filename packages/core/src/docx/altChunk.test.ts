import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { compareDocx } from "../compare/compare";
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { ALT_CHUNK_PLAIN_TEXT_LIMIT_BYTES, getAltChunkRelationshipId } from "./altChunk";
import { inspectDocxCompatibility } from "./compatibility";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { docxToMarkdown } from "./server/docxToMarkdown";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CHUNK = '<w:altChunk r:id="rIdChunk"/>';

type BuildDocxOptions = {
  contentType?: string;
  payload?: string | Uint8Array;
  relationshipMode?: "Internal" | "External";
  relationshipType?: "strict" | "transitional";
  defaultContentType?: string;
};

const buildDocx = async (chunk: string, options: BuildDocxOptions = {}): Promise<ArrayBuffer> => {
  const {
    contentType = "text/html",
    payload = "<html><body>Imported payload sentinel</body></html>",
    relationshipMode = "Internal",
    relationshipType = "transitional",
    defaultContentType,
  } = options;
  const extension = contentType === "text/plain" ? "txt" : "htm";
  const relationshipNamespace =
    relationshipType === "strict"
      ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
      : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const target =
    relationshipMode === "External" ? "https://example.test/afchunk.txt" : `afchunk.${extension}`;
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:body><w:p><w:r><w:t>Visible</w:t></w:r></w:p>${chunk}<w:sectPr/></w:body></w:document>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdChunk" Type="${relationshipNamespace}/aFChunk" Target="${target}"` +
      `${relationshipMode === "External" ? ' TargetMode="External"' : ""}/>` +
      "</Relationships>",
  );
  zip.file(
    "[Content_Types].xml",
    (await zip.file("[Content_Types].xml")?.async("text"))?.replace(
      "</Types>",
      `${defaultContentType ? `<Default Extension="${extension}" ContentType="${defaultContentType}"/>` : ""}` +
        `<Override PartName="/word/afchunk.${extension}" ContentType="${contentType}"/></Types>`,
    ) ?? "",
  );
  if (relationshipMode === "Internal") {
    zip.file(`word/afchunk.${extension}`, payload);
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXml = async (bytes: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(bytes)).file("word/document.xml")?.async("text")) ?? "";

describe("w:altChunk", () => {
  test("resolves the relationship id by expanded name", () => {
    expect(getAltChunkRelationshipId(CHUNK)).toBe("rIdChunk");
    expect(getAltChunkRelationshipId('<x:altChunk xmlns:x="urn:foreign" r:id="rIdChunk"/>')).toBe(
      undefined,
    );
  });

  test("keeps the payload opaque, reports it to readers, and preserves its XML on save", async () => {
    const input = await buildDocx(CHUNK);
    const parsed = await parseDocx(input, { preloadFonts: false });
    const reviewer = await FolioDocxReviewer.fromBuffer(input);

    expect(
      parsed.parseWarnings?.filter(({ code }) => code === PARSE_WARNING_CODES.altChunkUnsupported),
    ).toEqual([
      expect.objectContaining({
        count: 1,
        location: expect.objectContaining({ element: "w:altChunk" }),
      }),
    ]);
    expect(parsed.warnings?.join(" ")).toContain("diagnostic marker");
    expect(inspectDocxCompatibility(parsed).canSafelyEdit).toBe(false);
    expect(reviewer.getContentAsText()).toContain("Visible");
    expect(reviewer.getContentAsText()).toContain("Unsupported w:altChunk content");
    expect(
      reviewer.getContent().some(({ text }) => text.includes("Unsupported w:altChunk content")),
    ).toBe(true);
    expect(
      reviewer
        .snapshot()
        .blocks.some(({ text }) => text.includes("Unsupported w:altChunk content")),
    ).toBe(true);
    expect(await docxToMarkdown(input)).toContain("Unsupported w:altChunk content");

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    expect(await documentXml(saved)).toContain(CHUNK);
    const savedZip = await JSZip.loadAsync(saved);
    expect(await savedZip.file("word/afchunk.htm")?.async("text")).toBe(
      "<html><body>Imported payload sentinel</body></html>",
    );
    expect(await savedZip.file("word/_rels/document.xml.rels")?.async("text")).toContain(
      'Id="rIdChunk"',
    );
  });

  test("extracts a bounded internal text/plain payload for readers", async () => {
    const payload = "Plain imported sentinel.";
    const input = await buildDocx(CHUNK, { contentType: "text/plain", payload });
    const reviewer = await FolioDocxReviewer.fromBuffer(input);
    const reads = [
      reviewer
        .getContent()
        .map(({ text }) => text)
        .join("\n"),
      reviewer
        .snapshot()
        .blocks.map(({ text }) => text)
        .join("\n"),
      reviewer.getContentAsText(),
      await docxToMarkdown(input),
    ];
    for (const read of reads) {
      expect(read.split(payload).length - 1).toBe(1);
    }

    const changed = await buildDocx(CHUNK, {
      contentType: "text/plain",
      payload: "Changed payload.",
    });
    const compared = await compareDocx(input, changed, {
      author: "compare",
      timestamp: "2026-09-27T00:00:00.000Z",
    });
    if (compared.isErr()) {
      throw compared.error;
    }
    expect(compared.value.unsupported).toContainEqual(
      expect.objectContaining({ reason: "unsupported-content" }),
    );
  });

  test("part overrides take precedence over conflicting extension defaults", async () => {
    for (const [contentType, defaultContentType, exposed] of [
      ["text/plain", "text/html", true],
      ["text/html", "text/plain", false],
    ] as const) {
      const input = await buildDocx(CHUNK, {
        contentType,
        defaultContentType,
        payload: "Override precedence sentinel",
      });
      const text = (await FolioDocxReviewer.fromBuffer(input)).getContentAsText();
      expect(text.includes("Override precedence sentinel")).toBe(exposed);
    }
  });

  test("extracts each secondary story through its own relationship part", async () => {
    const zip = await JSZip.loadAsync(await createEmptyDocx());
    const stories = [
      {
        kind: "header",
        id: "rIdHeader",
        file: "header1.xml",
        root: "hdr",
        text: "Header payload",
        relationshipKind: "header",
      },
      {
        kind: "footer",
        id: "rIdFooter",
        file: "footer1.xml",
        root: "ftr",
        text: "Footer payload",
        relationshipKind: "footer",
      },
      {
        kind: "footnote",
        id: "rIdFootnote",
        file: "footnotes.xml",
        root: "footnotes",
        text: "Footnote payload",
        relationshipKind: "footnotes",
      },
      {
        kind: "endnote",
        id: "rIdEndnote",
        file: "endnotes.xml",
        root: "endnotes",
        text: "Endnote payload",
        relationshipKind: "endnotes",
      },
    ] as const;
    zip.file(
      "word/document.xml",
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p/>${CHUNK}` +
        `<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>` +
        `<w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`,
    );
    const documentRels = await zip.file("word/_rels/document.xml.rels")?.async("text");
    const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
    if (!documentRels || !contentTypes) {
      throw new Error("The fixture is missing its package declarations");
    }
    zip.file(
      "word/_rels/document.xml.rels",
      documentRels.replace(
        "</Relationships>",
        `<Relationship Id="rIdChunk" Type="${R}/aFChunk" Target="body.txt"/>` +
          stories
            .map(
              ({ relationshipKind, id, file }) =>
                `<Relationship Id="${id}" Type="${R}/${relationshipKind}" Target="${file}"/>`,
            )
            .join("") +
          "</Relationships>",
      ),
    );
    zip.file(
      "[Content_Types].xml",
      contentTypes.replace(
        "</Types>",
        '<Default Extension="txt" ContentType="text/plain"/>' +
          stories
            .map(
              ({ kind, file }) =>
                `<Override PartName="/word/${file}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`,
            )
            .join("") +
          "</Types>",
      ),
    );
    zip.file("word/body.txt", "Body payload");
    for (const { kind, file, root, text } of stories) {
      const content =
        kind === "footnote" || kind === "endnote"
          ? `<w:${kind} w:id="1">${CHUNK}</w:${kind}>`
          : CHUNK;
      zip.file(`word/${file}`, `<w:${root} xmlns:w="${W}" xmlns:r="${R}">${content}</w:${root}>`);
      zip.file(
        `word/_rels/${file}.rels`,
        `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}">` +
          `<Relationship Id="rIdChunk" Type="${R}/aFChunk" Target="${kind}.txt"/>` +
          "</Relationships>",
      );
      zip.file(`word/${kind}.txt`, text);
    }
    const input = await zip.generateAsync({ type: "arraybuffer" });
    const parsed = await parseDocx(input, {
      preloadFonts: false,
    });
    expect(
      parsed.package.document.content.find(({ type }) => type === "preservedBlock"),
    ).toMatchObject({
      type: "preservedBlock",
      readerText: "Body payload",
    });
    expect(parsed.package.headers?.get("rIdHeader")?.content.at(0)).toMatchObject({
      type: "preservedBlock",
      readerText: "Header payload",
    });
    expect(parsed.package.footers?.get("rIdFooter")?.content.at(0)).toMatchObject({
      type: "preservedBlock",
      readerText: "Footer payload",
    });
    expect(parsed.package.footnotes?.at(0)?.content.at(0)).toMatchObject({
      type: "preservedBlock",
      readerText: "Footnote payload",
    });
    expect(parsed.package.endnotes?.at(0)?.content.at(0)).toMatchObject({
      type: "preservedBlock",
      readerText: "Endnote payload",
    });
    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const savedZip = await JSZip.loadAsync(saved);
    expect(await savedZip.file("word/header1.xml")?.async("text")).toBe(
      await zip.file("word/header1.xml")?.async("text"),
    );
    expect(await savedZip.file("word/footer1.xml")?.async("text")).toBe(
      await zip.file("word/footer1.xml")?.async("text"),
    );
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(reopened.package.footnotes?.at(0)?.content.at(0)).toMatchObject({
      type: "preservedBlock",
      readerText: "Footnote payload",
    });
    expect(reopened.package.endnotes?.at(0)?.content.at(0)).toMatchObject({
      type: "preservedBlock",
      readerText: "Endnote payload",
    });
  });

  test("shares the plain-text byte limit between document and header stories", async () => {
    const largeText = "x".repeat(Math.floor(ALT_CHUNK_PLAIN_TEXT_LIMIT_BYTES / 2) + 1);
    const zip = await JSZip.loadAsync(
      await buildDocx(CHUNK, { contentType: "text/plain", payload: largeText }),
    );
    const documentRels = await zip.file("word/_rels/document.xml.rels")?.async("text");
    const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
    if (!documentRels || !contentTypes) {
      throw new Error("The fixture is missing its package declarations");
    }
    zip.file(
      "word/_rels/document.xml.rels",
      documentRels.replace(
        "</Relationships>",
        `<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>` +
          "</Relationships>",
      ),
    );
    zip.file(
      "[Content_Types].xml",
      contentTypes.replace(
        "</Types>",
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
          '<Override PartName="/word/header.txt" ContentType="text/plain"/></Types>',
      ),
    );
    zip.file("word/header1.xml", `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${CHUNK}</w:hdr>`);
    zip.file(
      "word/_rels/header1.xml.rels",
      `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}">` +
        `<Relationship Id="rIdChunk" Type="${R}/aFChunk" Target="header.txt"/>` +
        "</Relationships>",
    );
    zip.file("word/header.txt", largeText);
    const parsed = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    expect(
      parsed.package.document.content.find(({ type }) => type === "preservedBlock"),
    ).toMatchObject({
      type: "preservedBlock",
      readerText: largeText,
    });
    expect(parsed.package.headers?.get("rIdHeader")?.content.at(0)).toMatchObject({
      type: "preservedBlock",
      xml: CHUNK,
    });
    expect(parsed.package.headers?.get("rIdHeader")?.content.at(0)).not.toHaveProperty(
      "readerText",
    );
  });

  test("limits extracted plain text across repeated chunks", async () => {
    const payload = "x".repeat(Math.floor(ALT_CHUNK_PLAIN_TEXT_LIMIT_BYTES / 2) + 1);
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await buildDocx(CHUNK + CHUNK, { contentType: "text/plain", payload }),
    );
    const text = reviewer.getContentAsText();
    expect(text.split(payload).length - 1).toBe(1);
    expect(text.split("Unsupported w:altChunk content").length - 1).toBe(2);
  });

  test("extracts nested and Strict text/plain parts, and refuses external, oversized, or invalid UTF-8 payloads", async () => {
    const nested =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr/>${CHUNK}</w:tc></w:tr></w:tbl>`;
    const nestedPayload = "Nested imported text.";
    const nestedReader = await FolioDocxReviewer.fromBuffer(
      await buildDocx(nested, { contentType: "text/plain", payload: nestedPayload }),
    );
    expect(nestedReader.getContentAsText()).toContain(nestedPayload);

    const strictPayload = "Strict imported text.";
    const strictReader = await FolioDocxReviewer.fromBuffer(
      await buildDocx(CHUNK, {
        contentType: "text/plain",
        payload: strictPayload,
        relationshipType: "strict",
      }),
    );
    expect(strictReader.getContentAsText()).toContain(strictPayload);

    const externalReader = await FolioDocxReviewer.fromBuffer(
      await buildDocx(CHUNK, {
        contentType: "text/plain",
        payload: "Do not read external content.",
        relationshipMode: "External",
      }),
    );
    expect(externalReader.getContentAsText()).toContain("Unsupported w:altChunk content");
    expect(externalReader.getContentAsText()).not.toContain("Do not read external content.");

    const oversizedReader = await FolioDocxReviewer.fromBuffer(
      await buildDocx(CHUNK, {
        contentType: "text/plain",
        payload: "x".repeat(ALT_CHUNK_PLAIN_TEXT_LIMIT_BYTES + 1),
      }),
    );
    expect(oversizedReader.getContentAsText()).toContain("Unsupported w:altChunk content");
    expect(oversizedReader.getContentAsText()).not.toContain(
      "x".repeat(ALT_CHUNK_PLAIN_TEXT_LIMIT_BYTES),
    );

    const invalidUtf8Reader = await FolioDocxReviewer.fromBuffer(
      await buildDocx(CHUNK, { contentType: "text/plain", payload: new Uint8Array([0xff, 0xfe]) }),
    );
    expect(invalidUtf8Reader.getContentAsText()).toContain("Unsupported w:altChunk content");
    expect(invalidUtf8Reader.getContentAsText()).not.toContain("��");
  });

  test("comparison marks either side containing an altChunk as unsupported", async () => {
    const base = await buildDocx(CHUNK);
    const target = await buildDocx('<w:altChunk r:id="rIdOther"/>');
    const compared = await compareDocx(base, target, {
      author: "compare",
      timestamp: "2026-09-27T00:00:00.000Z",
    });
    if (compared.isErr()) {
      throw compared.error;
    }

    expect(compared.value.unsupported).toContainEqual(
      expect.objectContaining({ reason: "unsupported-content" }),
    );
  });
});
