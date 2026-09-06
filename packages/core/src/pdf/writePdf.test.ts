import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { inflateSync } from "node:zlib";
import { BLACK, DISPLAY_PRIMITIVE_KINDS } from "../display-list/primitives";
import type { DisplayFontFace, DisplayList, DisplayPrimitive } from "../display-list/types";
import { displayPointToPdf } from "./pageSpace";
import { writePdf, type PdfFontSource } from "./writePdf";

const TIMESTAMP = "2026-01-02T03:04:05Z";

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

const FACE: DisplayFontFace = {
  family: "Helvetica",
  weight: 400,
  italic: false,
  generic: "sans-serif",
  fontBoxAscentRatio: 0.9,
  fontBoxDescentRatio: 0.2,
};

const NO_FONTS: PdfFontSource = { load: () => null };

const PAGE_WIDTH_PX = 816;
const PAGE_HEIGHT_PX = 1056;

const listWith = (primitives: readonly DisplayPrimitive[]): DisplayList => ({
  pages: [
    {
      pageNumber: 1,
      widthPx: PAGE_WIDTH_PX,
      heightPx: PAGE_HEIGHT_PX,
      orientation: "portrait",
      primitives,
      links: [
        {
          rect: { xPx: 72, yPx: 90, widthPx: 100, heightPx: 20 },
          target: { kind: "external", href: "https://example.com/a?b=1" },
          tooltip: "Example",
        },
        {
          rect: { xPx: 72, yPx: 120, widthPx: 100, heightPx: 20 },
          target: { kind: "page", pageIndex: 0, yPx: 400 },
        },
      ],
    },
  ],
  fonts: [FACE],
  images: [],
  outline: [
    { title: "One", level: 0, pageIndex: 0, yPx: 100 },
    { title: "One.a", level: 1, pageIndex: 0, yPx: 200 },
    { title: "Two", level: 0, pageIndex: 0, yPx: 300 },
  ],
  metadata: { title: "Fixture", author: "folio" },
  unsupported: [],
});

const textRun = (text: string): DisplayPrimitive => ({
  kind: "glyphRun",
  font: 0,
  fontSizePx: 16,
  color: BLACK,
  xPx: 72,
  baselineYPx: 100,
  text,
  advancesPx: [...text].map(() => 9),
  direction: "ltr",
});

const FIXTURE = listWith([
  { kind: "rect", rect: { xPx: 10, yPx: 20, widthPx: 100, heightPx: 30 }, fill: BLACK },
  {
    kind: "line",
    x1Px: 10,
    y1Px: 200,
    x2Px: 300,
    y2Px: 200,
    stroke: { color: BLACK, thicknessPx: 2, pattern: "wavy" },
  },
  textRun("Hello PDF"),
]);

const write = (list: DisplayList, timestamp = TIMESTAMP) => {
  const result = writePdf(list, { fonts: NO_FONTS, timestamp });
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

/**
 * Every text stream in the file, decompressed. Image and font-program
 * streams are skipped: their dictionaries name a payload that is not text.
 */
const textStreams = (bytes: Uint8Array): readonly string[] => {
  const text = latin1(bytes);
  const out: string[] = [];
  // The dictionary capture must not run back through an earlier `<<`, or it
  // picks up keys from the objects before this one.
  const pattern = /<<((?:(?!<<)[^])*?)\/Length (\d+)>>\nstream\n/g;
  let match = pattern.exec(text);
  while (match !== null) {
    const [, dict = "", length = "0"] = match;
    const start = match.index + match[0].length;
    if (dict.includes("/FlateDecode") && !dict.includes("/Subtype") && !dict.includes("/Length1")) {
      out.push(latin1(inflateSync(bytes.subarray(start, start + Number(length)))));
    }
    match = pattern.exec(text);
  }
  return out;
};

/** The page's own stream, told apart by the base matrix it opens with. */
const contentStreamOf = (bytes: Uint8Array): string =>
  textStreams(bytes).find((stream) => stream.includes(" cm\n")) ?? "";

describe("determinism", () => {
  test("two runs with the same timestamp produce identical bytes", () => {
    const first = write(FIXTURE);
    const second = write(FIXTURE);
    expect(first.bytes.length).toBe(second.bytes.length);
    expect(Buffer.from(first.bytes).equals(Buffer.from(second.bytes))).toBe(true);
  });

  test("a different timestamp changes the bytes", () => {
    const first = write(FIXTURE);
    const second = write(FIXTURE, "2026-01-02T03:04:06Z");
    expect(Buffer.from(first.bytes).equals(Buffer.from(second.bytes))).toBe(false);
  });

  test("stamps the timestamp as a PDF date in UTC", () => {
    expect(latin1(write(FIXTURE).bytes)).toContain("/CreationDate (D:20260102030405+00'00')");
  });

  test("an offset timestamp and the same instant in Z agree byte for byte", () => {
    const zulu = write(FIXTURE, "2026-01-02T03:04:05Z");
    const offset = write(FIXTURE, "2026-01-02T05:04:05+02:00");
    expect(Buffer.from(zulu.bytes).equals(Buffer.from(offset.bytes))).toBe(true);
  });

  test("refuses a timestamp that is not an instant", () => {
    const result = writePdf(FIXTURE, { fonts: NO_FONTS, timestamp: "not a date" });
    expect(result.isErr()).toBe(true);
  });
});

describe("file structure", () => {
  const bytes = write(FIXTURE).bytes;
  const text = latin1(bytes);

  test("starts with a 1.7 header and a binary marker", () => {
    expect(text.startsWith("%PDF-1.7\n%")).toBe(true);
    expect(bytes[10]).toBeGreaterThan(127);
  });

  test("ends with startxref and %%EOF", () => {
    expect(text.endsWith("\n%%EOF\n")).toBe(true);
  });

  test("every xref offset resolves to the object it numbers", () => {
    const startxrefAt = text.lastIndexOf("startxref\n");
    expect(startxrefAt).toBeGreaterThan(0);
    const xrefOffset = Number.parseInt(text.slice(startxrefAt + "startxref\n".length), 10);
    expect(text.slice(xrefOffset, xrefOffset + 4)).toBe("xref");

    const headerAt = xrefOffset + "xref\n".length;
    const headerEnd = text.indexOf("\n", headerAt);
    const [, countText = "0"] = text.slice(headerAt, headerEnd).split(" ");
    const count = Number(countText);
    expect(count).toBeGreaterThan(1);

    const entriesStart = headerEnd + 1;
    const ENTRY_BYTES = 20;
    for (let id = 1; id < count; id += 1) {
      const entry = text.slice(
        entriesStart + id * ENTRY_BYTES,
        entriesStart + (id + 1) * ENTRY_BYTES,
      );
      expect(entry.endsWith(" n \n")).toBe(true);
      const offset = Number.parseInt(entry.slice(0, 10), 10);
      expect(text.slice(offset, offset + `${id} 0 obj`.length)).toBe(`${id} 0 obj`);
    }
    expect(text.slice(entriesStart, entriesStart + ENTRY_BYTES)).toBe("0000000000 65535 f \n");
  });

  test("the trailer names the catalog, the info dictionary and an /ID", () => {
    expect(text).toContain("/Size ");
    expect(text).toMatch(/\/Root \d+ 0 R/u);
    expect(text).toMatch(/\/Info \d+ 0 R/u);
    expect(text).toMatch(/\/ID \[<[0-9A-F]{64}> <[0-9A-F]{64}>\]/u);
  });

  test("writes the page box in points", () => {
    expect(text).toContain(`/MediaBox [0 0 ${PAGE_WIDTH_PX * 0.75} ${PAGE_HEIGHT_PX * 0.75}]`);
  });
});

describe("the y flip", () => {
  test("puts a display-list point at the matching PDF-space point", () => {
    const pageHeightPx = 200;
    expect(displayPointToPdf(pageHeightPx, 0, 0)).toEqual({ x: 0, y: 150 });
    expect(displayPointToPdf(pageHeightPx, 100, 200)).toEqual({ x: 75, y: 0 });
    expect(displayPointToPdf(pageHeightPx, 40, 80)).toEqual({ x: 30, y: 90 });
  });

  test("opens the content stream with the one base matrix and paints in px", () => {
    const content = contentStreamOf(write(FIXTURE).bytes);
    expect(content.startsWith(`0.75 0 0 -0.75 0 ${PAGE_HEIGHT_PX * 0.75} cm\n`)).toBe(true);
    // The rect keeps its display-list coordinates: nothing below the base
    // matrix does arithmetic on the page height.
    expect(content).toContain("10 20 100 30 re");
  });

  test("places a link annotation in default user space, not content space", () => {
    const text = latin1(write(FIXTURE).bytes);
    const top = displayPointToPdf(PAGE_HEIGHT_PX, 72, 90);
    const bottom = displayPointToPdf(PAGE_HEIGHT_PX, 172, 110);
    expect(text).toContain(`/Rect [${top.x} ${bottom.y} ${bottom.x} ${top.y}]`);
  });
});

describe("annotations and outline", () => {
  const text = latin1(write(FIXTURE).bytes);

  test("writes an external link as a URI action with no border", () => {
    expect(text).toContain("/Subtype /Link");
    expect(text).toContain("/Border [0 0 0]");
    expect(text).toContain("/S /URI /URI (https://example.com/a?b=1)");
  });

  test("writes a page link as a destination", () => {
    expect(text).toMatch(/\/Dest \[\d+ 0 R \/XYZ null 492 null\]/u);
  });

  test("nests the outline by level and counts every open node", () => {
    expect(text).toContain("/Type /Outlines");
    // Two roots and one child: the root count is the whole open subtree.
    expect(text).toMatch(/\/Type \/Outlines[^>]*\/Count 3/u);
    expect(text).toMatch(/\/Title <FEFF004F006E0065>[^>]*\/Count 1/u);
  });
});

describe("substitution", () => {
  test("reports a face the source cannot supply instead of throwing", () => {
    const result = write(FIXTURE);
    expect(result.substitutions).toEqual([
      {
        family: "Helvetica",
        weight: 400,
        italic: false,
        reason: "the font source supplied no bytes for this face",
      },
    ]);
    expect(latin1(result.bytes)).toContain("/BaseFont /Helvetica");
  });

  test("reports a face whose bytes will not parse", () => {
    const result = writePdf(FIXTURE, {
      fonts: { load: () => new Uint8Array([1, 2, 3, 4]) },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.substitutions).toHaveLength(1);
    expect(result.value.substitutions[0]?.reason).toContain("font parsing failed");
  });

  test("does not report a face nothing paints", () => {
    const unused: DisplayList = {
      ...FIXTURE,
      pages: [{ ...FIXTURE.pages[0]!, primitives: [], links: [] }],
    };
    expect(write(unused).substitutions).toEqual([]);
  });

  test("chooses the base-14 face from the generic category", () => {
    const serif: DisplayList = {
      ...FIXTURE,
      fonts: [{ ...FACE, generic: "serif", weight: 700, italic: true }],
    };
    expect(latin1(write(serif).bytes)).toContain("/BaseFont /Times-BoldItalic");
  });
});

describe("malformed display lists", () => {
  test("refuses a run whose advances do not match its code points", () => {
    const broken = listWith([
      {
        kind: "glyphRun",
        font: 0,
        fontSizePx: 12,
        color: BLACK,
        xPx: 0,
        baselineYPx: 0,
        text: "abc",
        advancesPx: [1, 2],
        direction: "ltr",
      },
    ]);
    expect(writePdf(broken, { fonts: NO_FONTS, timestamp: TIMESTAMP }).isErr()).toBe(true);
  });

  test("refuses a run that names a font outside the table", () => {
    const broken = listWith([{ ...textRun("x"), font: 7 }]);
    expect(writePdf(broken, { fonts: NO_FONTS, timestamp: TIMESTAMP }).isErr()).toBe(true);
  });

  test("refuses a document with no pages", () => {
    expect(
      writePdf({ ...FIXTURE, pages: [] }, { fonts: NO_FONTS, timestamp: TIMESTAMP }).isErr(),
    ).toBe(true);
  });
});

describe("graphics state", () => {
  test("names one ExtGState per distinct alpha, sorted", () => {
    const translucent = listWith([
      {
        kind: "rect",
        rect: { xPx: 0, yPx: 0, widthPx: 10, heightPx: 10 },
        fill: { r: 0, g: 0, b: 0, a: 0.5 },
      },
      {
        kind: "opacityGroup",
        opacity: 0.5,
        children: [
          {
            kind: "rect",
            rect: { xPx: 0, yPx: 0, widthPx: 10, heightPx: 10 },
            fill: { r: 0, g: 0, b: 0, a: 0.5 },
          },
        ],
      },
    ]);
    const bytes = write(translucent).bytes;
    const text = latin1(bytes);
    // 0.25 from the nested pair and 0.5 from both the group and the plain
    // fill: the group's alpha multiplies down rather than replacing.
    expect(text).toContain("/GS1 <</Type /ExtGState /ca 0.25 /CA 0.25>>");
    expect(text).toContain("/GS2 <</Type /ExtGState /ca 0.5 /CA 0.5>>");
    expect(text).not.toContain("/GS3");
  });
});

describe("right-to-left runs", () => {
  test("places the first logical code point at the run's right end", () => {
    const rtl = listWith([
      {
        kind: "glyphRun",
        font: 0,
        fontSizePx: 12,
        color: BLACK,
        xPx: 100,
        baselineYPx: 50,
        text: "abc",
        advancesPx: [10, 20, 30],
        direction: "rtl",
      },
    ]);
    const content = contentStreamOf(write(rtl).bytes);
    // The run occupies [100, 160]; logical "a" ends at 160, "b" at 150 and
    // "c" at 130, so their left edges are 150, 130 and 100.
    expect(content).toContain("1 0 0 -1 150 50 Tm");
    expect(content).toContain("1 0 0 -1 130 50 Tm");
    expect(content).toContain("1 0 0 -1 100 50 Tm");
  });
});

/** SOI, a baseline frame header for a 96x64 RGB image, then a scan marker. */
const TINY_JPEG = Uint8Array.from([
  0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x40, 0x00, 0x60, 0x03, 0x01, 0x11, 0x00, 0x02,
  0x11, 0x01, 0x03, 0x11, 0x01, 0xff, 0xda, 0x00, 0x02,
]);

const SAMPLE_BY_KIND = {
  glyphRun: textRun("a"),
  rect: { kind: "rect", rect: { xPx: 1, yPx: 2, widthPx: 3, heightPx: 4 }, fill: BLACK },
  line: {
    kind: "line",
    x1Px: 0,
    y1Px: 0,
    x2Px: 10,
    y2Px: 10,
    stroke: { color: BLACK, thicknessPx: 1, pattern: "double" },
  },
  image: {
    kind: "image",
    image: 0,
    rect: { xPx: 5, yPx: 6, widthPx: 20, heightPx: 10 },
    crop: { l: 0.1, t: 0.1, r: 0.1, b: 0.1 },
    opacity: 0.5,
  },
  clipGroup: {
    kind: "clipGroup",
    rect: { xPx: 0, yPx: 0, widthPx: 10, heightPx: 10 },
    children: [{ kind: "rect", rect: { xPx: 1, yPx: 1, widthPx: 2, heightPx: 2 }, fill: BLACK }],
  },
  rotateGroup: {
    kind: "rotateGroup",
    degrees: 90,
    originXPx: 10,
    originYPx: 20,
    children: [{ kind: "rect", rect: { xPx: 1, yPx: 1, widthPx: 2, heightPx: 2 }, fill: BLACK }],
  },
  opacityGroup: {
    kind: "opacityGroup",
    opacity: 0.5,
    children: [{ kind: "rect", rect: { xPx: 1, yPx: 1, widthPx: 2, heightPx: 2 }, fill: BLACK }],
  },
} as const satisfies Record<DisplayPrimitive["kind"], DisplayPrimitive>;

describe("primitive coverage", () => {
  // Total over the display list's own list of kinds: a kind added there and
  // not painted here fails to compile rather than silently painting nothing.
  test.each(DISPLAY_PRIMITIVE_KINDS)("paints a %s", (kind) => {
    const list: DisplayList = {
      ...listWith([SAMPLE_BY_KIND[kind]]),
      images: [{ format: "jpeg", bytes: TINY_JPEG, pixelWidth: 96, pixelHeight: 64 }],
    };
    const content = contentStreamOf(write(list).bytes);
    const operators = content.trimEnd().split("\n");
    // More than the base matrix: the primitive reached the page.
    expect(operators.length).toBeGreaterThan(1);
  });

  test("rotates about the given origin", () => {
    const content = contentStreamOf(write(listWith([SAMPLE_BY_KIND.rotateGroup])).bytes);
    expect(content).toContain("0 1 -1 0 30 10 cm");
  });

  test("clips a cropped image and scales it past the destination box", () => {
    const list: DisplayList = {
      ...listWith([SAMPLE_BY_KIND.image]),
      images: [{ format: "jpeg", bytes: TINY_JPEG, pixelWidth: 96, pixelHeight: 64 }],
    };
    const content = contentStreamOf(write(list).bytes);
    expect(content).toContain("5 6 20 10 re\nW n");
    // A 10% crop on each side scales the full image to 25x12.5 px and puts
    // its top-left at (2.5, 4.75), so the kept region fills the 20x10 box.
    expect(content).toContain("25 0 0 -12.5 2.5 17.25 cm");
    expect(content).toContain("/Im1 Do");
  });
});

const mutool = Bun.which("mutool");

describe.skipIf(mutool === null)("mutool", () => {
  test("extracts the text of a glyph run painted with a base-14 stand-in", async () => {
    const directory = mkdtempSync(join(tmpdir(), "folio-pdf-"));
    const file = join(directory, "out.pdf");
    await Bun.write(file, write(FIXTURE).bytes);
    const process = Bun.spawn({
      cmd: [mutool ?? "mutool", "draw", "-F", "stext", "-o", join(directory, "out.xml"), file],
      stdout: "pipe",
      stderr: "pipe",
    });
    const stderr = await new Response(process.stderr).text();
    expect(await process.exited).toBe(0);
    expect(stderr).not.toContain("error");
    const extracted = await Bun.file(join(directory, "out.xml")).text();
    const characters = [...extracted.matchAll(/<char c="(.)"/gu)].map(([, character]) => character);
    // A base-14 stand-in is positioned per code point, so the extractor
    // inserts synthetic spaces where the measurer's advances are wider than
    // the stand-in's: the characters and their order are what survive.
    expect(characters.join("").replace(/\s/gu, "")).toBe("HelloPDF");
  });
});
