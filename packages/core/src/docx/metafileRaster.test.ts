import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { extractMetafileRaster, isMetafileMimeType } from "./metafileRaster";

const FIXTURE = resolve(import.meta.dir, "__fixtures__/header-vml-emf.docx");

describe("extractMetafileRaster", () => {
  test("isMetafileMimeType recognizes EMF and WMF MIME types", () => {
    expect(isMetafileMimeType("image/x-emf")).toBe(true);
    expect(isMetafileMimeType("image/x-wmf")).toBe(true);
    expect(isMetafileMimeType("image/png")).toBe(false);
  });

  test("extracts the embedded PNG from the header-vml-emf fixture EMF", async () => {
    const buf = readFileSync(FIXTURE);
    const zip = await JSZip.loadAsync(buf);
    const emf = await zip.file("word/media/image1.emf")!.async("uint8array");

    const raster = extractMetafileRaster(emf);
    expect(raster).not.toBeNull();
    expect(raster!.mimeType).toBe("image/png");
    expect(raster!.bytes[0]).toBe(0x89);
    expect(raster!.bytes[1]).toBe(0x50);
    const tail = raster!.bytes.slice(-8);
    expect(String.fromCharCode(tail[0]!, tail[1]!, tail[2]!, tail[3]!)).toBe("IEND");
  });

  test("returns null for non-metafile bytes", () => {
    expect(extractMetafileRaster(new Uint8Array(64).fill(0x20))).toBeNull();
  });
});
