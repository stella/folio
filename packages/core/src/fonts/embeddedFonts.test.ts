import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import { extractEmbeddedFonts, getEmbeddedFontFaces } from "./embeddedFonts";
import { parseEmbeddedFontTable } from "./embeddedFontTable";
import { deobfuscateFont, isValidFontKey } from "./fontDeobfuscation";

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const GUID = "{001B70DC-AA60-4AD5-90EC-18A0948E1EAE}";

/**
 * A 64-byte stand-in font: a TrueType sfnt signature (`00 01 00 00`) followed
 * by deterministic filler. We obfuscate it so the extract path must
 * de-obfuscate to recover the recognizable header.
 */
function makeFont(): Uint8Array {
  const font = new Uint8Array(64);
  font.set([0x00, 0x01, 0x00, 0x00], 0);
  for (let i = 4; i < font.length; i++) {
    font[i] = (i * 17 + 5) & 0xff;
  }
  return font;
}

/** Whether the first four bytes are a known SFNT/OpenType signature. */
function hasSfntSignature(bytes: Uint8Array): boolean {
  const tag =
    ((bytes[0] ?? 0) << 24) | ((bytes[1] ?? 0) << 16) | ((bytes[2] ?? 0) << 8) | (bytes[3] ?? 0);
  const SIGNATURES = [
    0x00010000, // TrueType
    0x4f54544f, // 'OTTO' (CFF/OpenType)
    0x74727565, // 'true'
    0x74746366, // 'ttcf' (TrueType collection)
  ];
  return SIGNATURES.includes(tag >>> 0);
}

const CONTENT_TYPES = `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
</Types>`;

const PACKAGE_RELS = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_XML = `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:rFonts w:ascii="My Brand Sans" w:hAnsi="My Brand Sans"/></w:rPr><w:t>Hello</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const FONT_TABLE_XML = `${XML_DECLARATION}
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="My Brand Sans">
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:embedRegular r:id="rId1" w:fontKey="${GUID}" w:subsetted="true"/>
    <w:embedBold r:id="rId2" w:fontKey="${GUID}"/>
  </w:font>
</w:fonts>`;

const FONT_TABLE_RELS = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font2.odttf"/>
</Relationships>`;

async function buildDocx(options?: {
  regular?: Uint8Array;
  bold?: Uint8Array;
  fontTableRels?: string;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", PACKAGE_RELS);
  zip.file("word/document.xml", DOCUMENT_XML);
  zip.file("word/fontTable.xml", FONT_TABLE_XML);
  zip.file("word/_rels/fontTable.xml.rels", options?.fontTableRels ?? FONT_TABLE_RELS);
  if (options?.regular) {
    zip.file("word/fonts/font1.odttf", options.regular);
  }
  if (options?.bold) {
    zip.file("word/fonts/font2.odttf", options.bold);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

describe("deobfuscateFont", () => {
  test("XOR is symmetric: obfuscate then de-obfuscate returns the original", () => {
    const font = makeFont();
    const obfuscated = deobfuscateFont(font, GUID);
    // The header bytes change; the tail past byte 32 is untouched.
    expect(Array.from(obfuscated.slice(0, 4))).not.toEqual([0x00, 0x01, 0x00, 0x00]);
    expect(Array.from(obfuscated.slice(32))).toEqual(Array.from(font.slice(32)));

    const restored = deobfuscateFont(obfuscated, GUID);
    expect(Array.from(restored)).toEqual(Array.from(font));
  });

  test("does not mutate its input", () => {
    const font = makeFont();
    const copy = new Uint8Array(font);
    deobfuscateFont(font, GUID);
    expect(Array.from(font)).toEqual(Array.from(copy));
  });

  test("throws on an invalid key", () => {
    expect(() => deobfuscateFont(makeFont(), "not-a-guid")).toThrow();
  });

  test("isValidFontKey accepts a GUID with or without braces/hyphens, rejects junk", () => {
    expect(isValidFontKey(GUID)).toBe(true);
    expect(isValidFontKey("001B70DCAA604AD590EC18A0948E1EAE")).toBe(true);
    expect(isValidFontKey(undefined)).toBe(false);
    expect(isValidFontKey("{too-short}")).toBe(false);
  });
});

describe("parseEmbeddedFontTable", () => {
  test("reads the font name and each embed's rel id / key / subsetted flag", () => {
    const entries = parseEmbeddedFontTable(FONT_TABLE_XML);
    expect(entries).toHaveLength(1);
    const brand = entries[0];
    expect(brand?.name).toBe("My Brand Sans");
    expect(brand?.embedRegular).toEqual({ relId: "rId1", fontKey: GUID, subsetted: true });
    expect(brand?.embedBold).toEqual({ relId: "rId2", fontKey: GUID });
    expect(brand?.embedItalic).toBeUndefined();
  });

  test("returns an empty list for missing/blank input", () => {
    expect(parseEmbeddedFontTable(null)).toEqual([]);
    expect(parseEmbeddedFontTable("   ")).toEqual([]);
  });
});

describe("getEmbeddedFontFaces", () => {
  const font = makeFont();
  const obfuscated = deobfuscateFont(font, GUID);

  test("de-obfuscates each declared face with the right family/style/weight", () => {
    const fonts = new Map<string, ArrayBuffer>([
      ["word/fonts/font1.odttf", obfuscated.buffer],
      ["word/fonts/font2.odttf", obfuscated.buffer],
    ]);
    const faces = getEmbeddedFontFaces({
      fontTableXml: FONT_TABLE_XML,
      fontTableRelsXml: FONT_TABLE_RELS,
      fonts,
    });

    expect(faces).toHaveLength(2);
    const regular = faces.find((f) => f.weight === 400 && f.style === "normal");
    const bold = faces.find((f) => f.weight === 700 && f.style === "normal");
    expect(regular?.family).toBe("My Brand Sans");
    expect(regular?.subsetted).toBe(true);
    expect(bold?.subsetted).toBe(false);
    // The de-obfuscated bytes are a valid SFNT font, byte-identical to the input.
    expect(hasSfntSignature(regular?.bytes ?? new Uint8Array())).toBe(true);
    expect(Array.from(regular?.bytes ?? new Uint8Array())).toEqual(Array.from(font));
  });

  test("skips a face whose binary is missing, keeps the resolvable one", () => {
    const fonts = new Map<string, ArrayBuffer>([["word/fonts/font1.odttf", obfuscated.buffer]]);
    const faces = getEmbeddedFontFaces({
      fontTableXml: FONT_TABLE_XML,
      fontTableRelsXml: FONT_TABLE_RELS,
      fonts,
    });
    expect(faces.map((f) => f.weight)).toEqual([400]);
  });

  test("no rels or no font binaries → no faces", () => {
    const fonts = new Map<string, ArrayBuffer>([["word/fonts/font1.odttf", obfuscated.buffer]]);
    expect(
      getEmbeddedFontFaces({ fontTableXml: FONT_TABLE_XML, fontTableRelsXml: null, fonts }),
    ).toEqual([]);
    expect(
      getEmbeddedFontFaces({
        fontTableXml: FONT_TABLE_XML,
        fontTableRelsXml: FONT_TABLE_RELS,
        fonts: new Map(),
      }),
    ).toEqual([]);
  });
});

describe("extractEmbeddedFonts (buffer)", () => {
  test("end-to-end: unzips, resolves rels, de-obfuscates to a valid font", async () => {
    const font = makeFont();
    const obfuscated = deobfuscateFont(font, GUID);
    const buffer = await buildDocx({ regular: obfuscated, bold: obfuscated });

    const faces = await extractEmbeddedFonts(buffer);
    expect(faces).toHaveLength(2);
    const regular = faces.find((f) => f.weight === 400 && f.style === "normal");
    expect(regular?.family).toBe("My Brand Sans");
    expect(hasSfntSignature(regular?.bytes ?? new Uint8Array())).toBe(true);
    expect(Array.from(regular?.bytes ?? new Uint8Array())).toEqual(Array.from(font));
  });

  test("a document with no embedded fonts yields an empty list", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", CONTENT_TYPES);
    zip.file("_rels/.rels", PACKAGE_RELS);
    zip.file("word/document.xml", DOCUMENT_XML);
    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    expect(await extractEmbeddedFonts(buffer)).toEqual([]);
  });
});

describe("round-trip preservation", () => {
  test("save keeps fontTable.xml and word/fonts/* byte-identical", async () => {
    const font = makeFont();
    const obfuscated = deobfuscateFont(font, GUID);
    const buffer = await buildDocx({ regular: obfuscated, bold: obfuscated });

    const originalZip = await JSZip.loadAsync(buffer);
    const originalFontTable = await originalZip.file("word/fontTable.xml")!.async("text");
    const originalFontBytes = new Uint8Array(
      await originalZip.file("word/fonts/font1.odttf")!.async("arraybuffer"),
    );

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const repacked = await repackDocx(doc, { updateModifiedDate: false });

    const outZip = await JSZip.loadAsync(repacked);
    const outFontTable = await outZip.file("word/fontTable.xml")!.async("text");
    const outFontBytes = new Uint8Array(
      await outZip.file("word/fonts/font1.odttf")!.async("arraybuffer"),
    );

    expect(outFontTable).toBe(originalFontTable);
    // Still obfuscated, byte-identical to what we packed.
    expect(Array.from(outFontBytes)).toEqual(Array.from(originalFontBytes));
  });
});
