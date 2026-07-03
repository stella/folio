import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { getHeaderFooterVerbatimXml } from "./headerFooterVerbatim";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

const FIXTURE = resolve(import.meta.dir, "__fixtures__/header-vml-emf.docx");

function load(path: string): ArrayBuffer {
  const buf = readFileSync(path);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

describe("parseDocx — header with EMF media", () => {
  test("loads EMF media with a browser-renderable PNG data URL", async () => {
    const doc = await parseDocx(load(FIXTURE), { preloadFonts: false });

    const media = doc.package.media?.get("word/media/image1.emf");
    expect(media).toBeDefined();
    expect(media!.mimeType).toBe("image/x-emf");
    expect(media!.dataUrl?.startsWith("data:image/png")).toBe(true);
  });

  test("captures verbatim header XML for byte-identical re-export", async () => {
    const doc = await parseDocx(load(FIXTURE), { preloadFonts: false });
    const header = doc.package.headers?.get("rId6");
    expect(header).toBeDefined();
    expect(getHeaderFooterVerbatimXml(header!)).toContain("<w:hdr");
    expect(getHeaderFooterVerbatimXml(header!)).toContain('r:id="rId1"');
    expect(getHeaderFooterVerbatimXml(header!)).toContain("<w:object");
  });

  test("mediaResolver hook can override the display URL", async () => {
    const seen: string[] = [];
    const doc = await parseDocx(load(FIXTURE), {
      preloadFonts: false,
      mediaResolver: async (file) => {
        seen.push(file.mimeType);
        return file.mimeType === "image/x-emf" ? "data:image/png;base64,OVERRIDE" : undefined;
      },
    });
    expect(seen).toContain("image/x-emf");
    const media = doc.package.media!.get("word/media/image1.emf")!;
    expect(media.dataUrl).toBe("data:image/png;base64,OVERRIDE");
  });
});

describe("repackDocx — header round-trip", () => {
  test("unedited header/footer XML is byte-identical after round-trip", async () => {
    const inBuf = load(FIXTURE);
    const inZip = await JSZip.loadAsync(inBuf);
    const headerIn = await inZip.file("word/header1.xml")!.async("string");
    const footerIn = await inZip.file("word/footer1.xml")!.async("string");
    const emfIn = await inZip.file("word/media/image1.emf")!.async("uint8array");

    const doc = await parseDocx(inBuf, { preloadFonts: false });
    const outBuf = await repackDocx(doc);
    const outZip = await JSZip.loadAsync(outBuf);

    const headerOut = await outZip.file("word/header1.xml")!.async("string");
    const footerOut = await outZip.file("word/footer1.xml")!.async("string");
    const emfOut = await outZip.file("word/media/image1.emf")!.async("uint8array");

    expect(headerOut).toBe(headerIn);
    expect(footerOut).toBe(footerIn);
    expect(emfOut.byteLength).toBe(emfIn.byteLength);
    expect(emfOut).toEqual(emfIn);
  });
});
