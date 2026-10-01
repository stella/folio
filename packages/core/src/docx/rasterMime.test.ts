import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import type { RasterMimeType } from "./rasterMime";
import { unzipDocx } from "./unzip";

const RASTER_HEADERS = {
  "image/png": [[0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]],
  "image/jpeg": [[0xff, 0xd8, 0xff, 0xe0]],
  "image/gif": [[0x47, 0x49, 0x46, 0x38, 0x39, 0x61]],
  "image/bmp": [[0x42, 0x4d, 0x20, 0, 0, 0]],
  "image/webp": [[0x52, 0x49, 0x46, 0x46, 0x10, 0, 0, 0, 0x57, 0x45, 0x42, 0x50]],
  "image/tiff": [
    [0x49, 0x49, 0x2a, 0],
    [0x4d, 0x4d, 0, 0x2a],
    [0x49, 0x49, 0x2b, 0],
    [0x4d, 0x4d, 0, 0x2b],
  ],
} as const satisfies Record<RasterMimeType, readonly (readonly number[])[]>;

const RASTER_EXTENSIONS = ["png", "jpg", "jpeg", "gif", "bmp", "webp", "tif", "tiff"];

const createPackage = () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    "word/document.xml",
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
  );
  return zip;
};

describe("raster package MIME detection", () => {
  test("uses the signature for every raster format and file extension", async () => {
    const zip = createPackage();
    const cases = [];
    for (const [mimeType, variants] of Object.entries(RASTER_HEADERS)) {
      for (const signature of variants) {
        for (const extension of RASTER_EXTENSIONS) {
          const path = `word/media/raster${cases.length}.${extension}`;
          const bytes = new Uint8Array(signature);
          cases.push({ path, mimeType, bytes });
          zip.file(path, bytes);
        }
      }
    }
    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    const raw = await unzipDocx(buffer);
    const document = await parseDocx(buffer, { preloadFonts: false });

    expect(raw.media.size).toBe(cases.length);
    for (const { path, mimeType, bytes } of cases) {
      expect(raw.media.get(path)).toEqual(bytes.buffer);
      const media = document.package.media?.get(path);
      expect(media?.mimeType).toBe(mimeType);
      expect(media?.dataUrl).toStartWith(`data:${mimeType};base64,`);
      expect(media?.data).toEqual(bytes.buffer);
    }
  });

  test("requires both declared and detected raster types to satisfy the allowlist", async () => {
    const zip = createPackage();
    zip.file("word/media/declared.png", new Uint8Array(RASTER_HEADERS["image/gif"][0]));
    zip.file("word/media/declared.gif", new Uint8Array(RASTER_HEADERS["image/png"][0]));
    const raw = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      allowedMediaMimeTypes: new Set(["image/png"]),
    });
    expect(raw.media.size).toBe(0);
  });

  test("rejects raster and metafile class changes without removing original entries", async () => {
    const zip = createPackage();
    const emf = new Uint8Array(44);
    emf[0] = 1;
    emf.set([0x20, 0x45, 0x4d, 0x46], 40);
    const rejected = {
      "word/media/metafile.png": emf,
      "word/media/raster.emf": new Uint8Array(RASTER_HEADERS["image/png"][0]),
      "word/media/vector.png": new TextEncoder().encode(
        '<svg xmlns="http://www.w3.org/2000/svg"/>',
      ),
      "word/media/unknown.gif": new Uint8Array([0, 1, 2, 3]),
      "word/media/truncated.png": new Uint8Array([0x89, 0x50]),
      "word/media/malformed.tiff": new Uint8Array([0x49, 0x49, 0, 0]),
    };
    for (const [path, bytes] of Object.entries(rejected)) {
      zip.file(path, bytes);
    }
    const raw = await unzipDocx(await zip.generateAsync({ type: "arraybuffer" }));
    expect(raw.media.size).toBe(0);
    for (const path of Object.keys(rejected)) {
      expect(raw.originalZip.file(path)).not.toBeNull();
    }
  });
});
