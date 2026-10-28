import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

import { getHeaderFooterVerbatimXml } from "./headerFooterVerbatim";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";

const FIXTURE = resolve(import.meta.dir, "__fixtures__/header-vml-emf.docx");

function load(path: string): ArrayBuffer {
  const buf = readFileSync(path);
  return buf.buffer.slice(
    buf.byteOffset,
    buf.byteOffset + buf.byteLength,
  ) as ArrayBuffer;
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

  test("media outside the relationship graph is stored without conversion", async () => {
    // Conversion is driven by references, not by what the ZIP happens to hold:
    // a `word/media/` entry no relationship points at is kept byte for byte so
    // the round-trip is unchanged, but nothing decodes or re-encodes it.
    const zip = await ZipArchive.read(load(FIXTURE));
    const referencedBytes = await zip.entry("word/media/image1.emf")!.bytes();
    zip.set("word/media/orphan.emf", referencedBytes);
    const seeded = (await zip.generate({ compression: "store" })).buffer;

    const doc = await parseDocx(seeded, { preloadFonts: false });

    const referenced = doc.package.media?.get("word/media/image1.emf");
    expect(referenced?.dataUrl?.startsWith("data:image/png")).toBe(true);

    const orphan = doc.package.media?.get("word/media/orphan.emf");
    expect(orphan).toBeDefined();
    expect(orphan!.data.byteLength).toBe(referencedBytes.byteLength);
    expect(orphan!.dataUrl?.startsWith("data:image/x-emf")).toBe(true);
  });

  test("mediaResolver hook can override the display URL", async () => {
    const seen: string[] = [];
    const doc = await parseDocx(load(FIXTURE), {
      preloadFonts: false,
      mediaResolver: async (file) => {
        seen.push(file.mimeType);
        return file.mimeType === "image/x-emf"
          ? "data:image/png;base64,OVERRIDE"
          : undefined;
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
    const inZip = await ZipArchive.read(inBuf);
    const headerIn = await inZip.entry("word/header1.xml")!.text();
    const footerIn = await inZip.entry("word/footer1.xml")!.text();
    const emfIn = await inZip.entry("word/media/image1.emf")!.bytes();

    const doc = await parseDocx(inBuf, { preloadFonts: false });
    const outBuf = await repackDocx(doc);
    const outZip = await ZipArchive.read(outBuf);

    const headerOut = await outZip.entry("word/header1.xml")!.text();
    const footerOut = await outZip.entry("word/footer1.xml")!.text();
    const emfOut = await outZip.entry("word/media/image1.emf")!.bytes();

    expect(headerOut).toBe(headerIn);
    expect(footerOut).toBe(footerIn);
    expect(emfOut.byteLength).toBe(emfIn.byteLength);
    expect(emfOut).toEqual(emfIn);
  });
});
