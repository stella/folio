import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { DocxSecurityError, extractFile, getFileList, unzipDocx } from "./unzip";

const UTF16_BYTE_ORDERS = {
  big: "big",
  little: "little",
} as const;

const BYTE_ORDER_MARKS = {
  absent: "absent",
  present: "present",
} as const;

type EncodeUtf16Options = {
  byteOrder: (typeof UTF16_BYTE_ORDERS)[keyof typeof UTF16_BYTE_ORDERS];
  byteOrderMark: (typeof BYTE_ORDER_MARKS)[keyof typeof BYTE_ORDER_MARKS];
};

const encodeUtf16 = (
  value: string,
  { byteOrder, byteOrderMark }: EncodeUtf16Options,
): Uint8Array => {
  const prefixLength = byteOrderMark === BYTE_ORDER_MARKS.present ? 2 : 0;
  const bytes = new Uint8Array(prefixLength + value.length * 2);
  if (byteOrderMark === BYTE_ORDER_MARKS.present) {
    bytes.set(byteOrder === UTF16_BYTE_ORDERS.little ? [0xff, 0xfe] : [0xfe, 0xff]);
  }
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < value.length; index += 1) {
    view.setUint16(
      prefixLength + index * 2,
      value.charCodeAt(index),
      byteOrder === UTF16_BYTE_ORDERS.little,
    );
  }
  return bytes;
};

const UTF16_CASES = [
  { byteOrder: UTF16_BYTE_ORDERS.little, byteOrderMark: BYTE_ORDER_MARKS.present },
  { byteOrder: UTF16_BYTE_ORDERS.big, byteOrderMark: BYTE_ORDER_MARKS.present },
  { byteOrder: UTF16_BYTE_ORDERS.little, byteOrderMark: BYTE_ORDER_MARKS.absent },
  { byteOrder: UTF16_BYTE_ORDERS.big, byteOrderMark: BYTE_ORDER_MARKS.absent },
] as const satisfies readonly EncodeUtf16Options[];

describe("unzipDocx security limits", () => {
  test("rejects input larger than the configured limit", async () => {
    const error = await getRejectedError(unzipDocx(new ArrayBuffer(2), { maxInputBytes: 1 }));

    expect(error).toBeInstanceOf(DocxSecurityError);
  });

  test("rejects archives with too many file entries", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");

    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    const error = await getRejectedError(unzipDocx(buffer, { maxFiles: 1 }));

    expect(error).toBeInstanceOf(DocxSecurityError);
  });

  test("accepts large document XML entries within the default limit", async () => {
    const zip = new JSZip();
    const largeBody = "x".repeat(26 * 1024 * 1024);
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", `<w:document>${largeBody}</w:document>`);

    const content = await unzipDocx(
      await zip.generateAsync({
        compression: "DEFLATE",
        type: "arraybuffer",
      }),
    );

    expect(content.documentXml?.length).toBe("<w:document></w:document>".length + largeBody.length);
  });

  test("accepts large header XML entries within the default limit", async () => {
    const zip = new JSZip();
    const largeHeader = "x".repeat(65 * 1024 * 1024);
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/header3.xml", `<w:hdr>${largeHeader}</w:hdr>`);

    const content = await unzipDocx(
      await zip.generateAsync({
        compression: "DEFLATE",
        type: "arraybuffer",
      }),
    );

    expect(content.headers.get("header3.xml")?.length).toBe(
      "<w:hdr></w:hdr>".length + largeHeader.length,
    );
  });

  for (const utf16 of UTF16_CASES) {
    test(`decodes UTF-16 ${utf16.byteOrder} endian XML with byte-order mark ${utf16.byteOrderMark}`, async () => {
      const zip = new JSZip();
      const documentXml =
        '<?xml version="1.0" encoding="UTF-16"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Žluťoučký kůň</w:t></w:r></w:p></w:body></w:document>';
      zip.file("[Content_Types].xml", "<Types />");
      zip.file("word/document.xml", encodeUtf16(documentXml, utf16));

      const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));

      expect(content.documentXml).toBe(documentXml);
      expect(await extractFile(content, "word/document.xml")).toBe(documentXml);
    });
  }

  test("accepts media-heavy packages within the default file-count limit", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    for (let index = 0; index < 3800; index += 1) {
      zip.file(`word/media/image${index}.png`, new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x00]));
    }

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));

    expect(content.media.size).toBe(3800);
  });

  test("repairs archives missing the end-of-central-directory tail", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");

    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    const truncated = truncateAfterEndOfCentralDirectoryCounts(buffer);
    const content = await unzipDocx(truncated);

    expect(content.documentXml).toBe("<w:document />");
    expect(content.originalBuffer.byteLength).toBe(buffer.byteLength);
    expect(getFileList(content)).toEqual(["[Content_Types].xml", "word/document.xml"]);
  });

  test("rejects unsafe archive paths", async () => {
    const zip = new JSZip();
    zip.file("/word/document.xml", "<w:document />");

    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    const error = await getRejectedError(unzipDocx(buffer));

    expect(error).toBeInstanceOf(DocxSecurityError);
  });

  test("counts an entry the parser never reads toward the expansion ceiling", async () => {
    // A save carries every part of the package, so an entry folio does not
    // model still passes through the host. The ceiling has to see it.
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/embeddings/workbook.xlsx", "\0".repeat(4 * 1024 * 1024));

    const buffer = await zip.generateAsync({ compression: "DEFLATE", type: "arraybuffer" });
    const error = await getRejectedError(
      unzipDocx(buffer, { maxTotalUncompressedBytes: 1024 * 1024 }),
    );

    expect(error).toBeInstanceOf(DocxSecurityError);
  });

  test("carries a macro project through the archive without reading it", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/vbaProject.bin", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]));

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));

    expect(content.allXml.has("word/vbaProject.bin")).toBe(false);
    expect(getFileList(content)).not.toContain("word/vbaProject.bin");
    expect(await extractFile(content, "word/vbaProject.bin")).toBeNull();
    expect(content.originalZip.file("word/vbaProject.bin")).not.toBeNull();
  });

  test("skips media entries with mismatched content signatures", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/media/image1.png", new Uint8Array([0x00, 0x00, 0x00, 0x00]));

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));

    expect(content.media.size).toBe(0);
  });

  test("loads valid media without eager data URL conversion", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/media/image1.png", new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x00]));

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));

    expect(content.media.has("word/media/image1.png")).toBe(true);
  });

  test("can keep unused XML in the source package without decompressing it", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("customXml/item1.xml", "<data />");

    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    const fullContent = await unzipDocx(buffer);
    const content = await unzipDocx(buffer, { extractAllXml: false });

    expect(fullContent.allXml.get("customXml/item1.xml")).toBe("<data />");
    expect(content.originalBuffer).toBe(buffer);
    expect(content.allXml.has("customXml/item1.xml")).toBe(false);
    expect(content.originalZip.file("customXml/item1.xml")).not.toBeNull();
  });

  test("skips oversized raster media instead of rejecting the document", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/media/image1.png", new Uint8Array([0x89, 0x50, 0x4e, 0x47]));

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      maxMediaBytes: 3,
    });

    expect(content.media.size).toBe(0);
    expect(getFileList(content)).toContain("word/media/image1.png");
    expect(content.warnings).toContain(
      "Skipped oversized media file: word/media/image1.png; original entry preserved for round-trip.",
    );
  });

  test("does not expose active or embedded payload entries for preservation", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/vbaProject.bin", new Uint8Array([0x00]));
    zip.file("word/embeddings/oleObject1.bin", new Uint8Array([0x00]));

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));

    expect(content.media.size).toBe(0);
    expect(getFileList(content)).toEqual(["[Content_Types].xml", "word/document.xml"]);
    expect(await extractFile(content, "word/vbaProject.bin")).toBeNull();
  });

  test("preserves known vector image entries without loading them by default", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/media/image1.emf", new Uint8Array([0x01, 0x02]));

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));

    expect(content.media.size).toBe(0);
    expect(getFileList(content)).toContain("word/media/image1.emf");
  });

  test("preserves large vector image entries without extracting them", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/media/image1.emf", new Uint8Array(26 * 1024 * 1024));

    const content = await unzipDocx(
      await zip.generateAsync({
        compression: "DEFLATE",
        type: "arraybuffer",
      }),
    );

    expect(content.media.size).toBe(0);
    expect(getFileList(content)).toContain("word/media/image1.emf");
  });

  test("skips oversized embedded fonts without rejecting the document", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("word/document.xml", "<w:document />");
    zip.file("word/fonts/font1.odttf", new Uint8Array([1, 2, 3, 4]));

    const content = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      maxFontBytes: 3,
    });

    expect(content.fonts.size).toBe(0);
    expect(getFileList(content)).toContain("word/fonts/font1.odttf");
  });
});

function truncateAfterEndOfCentralDirectoryCounts(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  for (let offset = bytes.byteLength - 22; offset >= 0; offset -= 1) {
    if (
      bytes[offset] === 0x50 &&
      bytes[offset + 1] === 0x4b &&
      bytes[offset + 2] === 0x05 &&
      bytes[offset + 3] === 0x06
    ) {
      return bytes.slice(0, offset + 12).buffer;
    }
  }

  throw new Error("Could not find ZIP end-of-central-directory record");
}

const getRejectedError = async (promise: Promise<unknown>) =>
  promise.then(() => null).catch((error: unknown) => error);
