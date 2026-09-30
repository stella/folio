import { describe, expect, spyOn, test } from "bun:test";
import JSZip from "jszip";

import {
  archiveWithRawDeflate,
  compressibleArchive,
  manyEntryArchive,
  rawDeflateOfZeros,
  understatedEntryArchive,
  withDeclaredSize,
} from "./__tests__/archiveInflationFixtures";
import { DocxParseError, parseDocx } from "./parser";
import { DocxSecurityError, extractFile, getFileList, unzipDocx } from "./unzip";

/**
 * The filler in the size tests below is one repeated byte, which compresses
 * far past the ratio cap; lifting the cap keeps those tests on the size limits.
 */
const WITHOUT_RATIO_LIMIT = { maxCompressionRatio: Number.MAX_SAFE_INTEGER } as const;

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer => bytes.slice().buffer;

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
      WITHOUT_RATIO_LIMIT,
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
      WITHOUT_RATIO_LIMIT,
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
      WITHOUT_RATIO_LIMIT,
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

describe("unzipDocx archive inflation limits", () => {
  test("stops a part that inflates past its declared size", async () => {
    const buffer = toArrayBuffer(await understatedEntryArchive("word/header1.xml"));

    const error = await getRejectedError(unzipDocx(buffer));

    expect(error).toBeInstanceOf(DocxSecurityError);
    expect(error).toHaveProperty(
      "message",
      "DOCX entry inflates past its declared size: word/header1.xml",
    );
  });

  test("stops a part the parser never reads when it inflates past its declared size", async () => {
    const buffer = toArrayBuffer(await understatedEntryArchive("word/embeddings/object1.bin"));

    const error = await getRejectedError(unzipDocx(buffer, { extractAllXml: false }));

    expect(error).toHaveProperty(
      "message",
      "DOCX entry inflates past its declared size: word/embeddings/object1.bin",
    );
  });

  test("still opens a package whose unread part is shorter than it declares", async () => {
    const source = await compressibleArchive({
      entryPath: "word/embeddings/object1.bin",
      inflatedMebibytes: 1,
    });
    const buffer = toArrayBuffer(
      withDeclaredSize(source, "word/embeddings/object1.bin", 1024 * 1024 + 16),
    );

    const content = await unzipDocx(buffer);

    expect(content.documentXml).toBe("<w:document/>");
  });

  test("refuses a part whose declared expansion passes the ratio cap", async () => {
    const buffer = toArrayBuffer(
      await compressibleArchive({ entryPath: "word/header1.xml", inflatedMebibytes: 5 }),
    );

    const error = await getRejectedError(unzipDocx(buffer));

    expect(error).toHaveProperty(
      "message",
      "DOCX entry exceeds the maximum compression ratio: word/header1.xml",
    );
  });

  test("refuses a package whose parts together pass the ratio cap", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file("word/document.xml", "<w:document/>");
    for (const index of [1, 2, 3]) {
      zip.file(`word/header${String(index)}.xml`, "x".repeat(2 * 1024 * 1024));
    }
    const buffer = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });

    const error = await getRejectedError(unzipDocx(buffer));

    expect(error).toHaveProperty("message", "DOCX file exceeds the maximum compression ratio");
  });

  test("lets an integrator tighten the ratio cap", async () => {
    const buffer = toArrayBuffer(
      await compressibleArchive({ entryPath: "word/header1.xml", inflatedMebibytes: 5 }),
    );

    await expect(unzipDocx(buffer, WITHOUT_RATIO_LIMIT)).resolves.toBeDefined();
    const error = await getRejectedError(unzipDocx(buffer, { maxCompressionRatio: 2 }));
    expect(error).toBeInstanceOf(DocxSecurityError);
  });

  test("counts archive records before the archive is parsed", async () => {
    const buffer = toArrayBuffer(await manyEntryArchive(40));
    const loadAsync = spyOn(JSZip, "loadAsync");

    try {
      const error = await getRejectedError(unzipDocx(buffer, { maxFiles: 10 }));

      expect(error).toHaveProperty("message", "DOCX file contains too many entries");
      expect(loadAsync).not.toHaveBeenCalled();
    } finally {
      loadAsync.mockRestore();
    }
  });

  test("stops a part declared at 64 bytes that would inflate to a gibibyte", async () => {
    const buffer = toArrayBuffer(
      await archiveWithRawDeflate({
        entryPath: "word/embeddings/object1.bin",
        deflated: rawDeflateOfZeros(1024),
        declaredBytes: 64,
      }),
    );

    const error = await getRejectedError(unzipDocx(buffer));

    expect(error).toHaveProperty(
      "message",
      "DOCX entry inflates past its declared size: word/embeddings/object1.bin",
    );
  });

  test("refuses parts that together pass the cumulative limit", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file("word/document.xml", `<w:document>${"a".repeat(600)}</w:document>`);
    zip.file("word/header1.xml", `<w:hdr>${"b".repeat(600)}</w:hdr>`);
    const buffer = await zip.generateAsync({ type: "arraybuffer" });

    await expect(unzipDocx(buffer, { maxTotalUncompressedBytes: 1000 })).rejects.toHaveProperty(
      "message",
      "DOCX file expands beyond the maximum allowed size",
    );
    await expect(unzipDocx(buffer, { maxTotalUncompressedBytes: 2000 })).resolves.toBeDefined();
  });

  test("still opens a package whose unread part cannot be decompressed", async () => {
    const buffer = toArrayBuffer(
      await archiveWithRawDeflate({
        entryPath: "word/embeddings/object1.bin",
        deflated: new Uint8Array([0xff, 0xff, 0xff, 0xff]),
        declaredBytes: 100,
      }),
    );

    const content = await unzipDocx(buffer);

    expect(content.documentXml).toBe("<w:document/>");
  });

  test("leaves a highly compressible binary part to the byte and package limits", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types/>");
    zip.file("word/document.xml", "<w:document/>");
    zip.file("word/media/image1.bmp", new Uint8Array(5 * 1024 * 1024).fill(0xff), {
      compression: "DEFLATE",
    });
    // Incompressible padding keeps the package as a whole under the ratio cap.
    const padding = crypto.getRandomValues(new Uint8Array(64 * 1024));
    zip.file("word/embeddings/padding.bin", padding);
    const buffer = await zip.generateAsync({ type: "arraybuffer" });

    await expect(unzipDocx(buffer)).resolves.toBeDefined();
  });

  test("rejects an invalid ratio limit", async () => {
    const buffer = toArrayBuffer(await manyEntryArchive(1));

    for (const maxCompressionRatio of [Number.NaN, -1, 1.5]) {
      // oxlint-disable-next-line no-await-in-loop -- each case is independent and tiny
      expect(await getRejectedError(unzipDocx(buffer, { maxCompressionRatio }))).toBeInstanceOf(
        RangeError,
      );
    }
  });

  test("can leave unread parts unchecked for a caller that discards the package", async () => {
    const buffer = toArrayBuffer(await understatedEntryArchive("word/embeddings/object1.bin"));

    const content = await unzipDocx(buffer, {}, { verifyUnreadEntries: false });

    expect(content.documentXml).toBe("<w:document/>");
  });

  test("parseDocx refuses the same packages", async () => {
    const understated = await getRejectedError(
      parseDocx(await understatedEntryArchive("word/header1.xml")),
    );
    const compressed = await getRejectedError(
      parseDocx(await compressibleArchive({ entryPath: "word/header1.xml", inflatedMebibytes: 5 })),
    );

    for (const error of [understated, compressed]) {
      expect(error).toBeInstanceOf(DocxParseError);
      expect(error).toHaveProperty("cause", expect.any(DocxSecurityError));
    }
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
