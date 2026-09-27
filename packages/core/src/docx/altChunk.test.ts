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
const CHUNK = '<w:altChunk r:id="rIdChunk"/>';

type BuildDocxOptions = {
  contentType?: string;
  payload?: string | Uint8Array;
  relationshipMode?: "Internal" | "External";
  relationshipType?: "strict" | "transitional";
};

const buildDocx = async (chunk: string, options: BuildDocxOptions = {}): Promise<ArrayBuffer> => {
  const {
    contentType = "text/html",
    payload = "<html><body>Imported payload sentinel</body></html>",
    relationshipMode = "Internal",
    relationshipType = "transitional",
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
