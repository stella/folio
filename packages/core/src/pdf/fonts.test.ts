import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { inflateSync } from "node:zlib";
import { BLACK } from "../display-list/primitives";
import type { DisplayFontFace, DisplayList } from "../display-list/types";
import { parseSfnt, type SfntFont } from "../fonts/sfnt/parse";
import {
  readTestFont,
  TEST_FONTS_INSTALLED,
  TEST_FONTS_SKIP_REASON,
} from "../fonts/sfnt/__tests__/testFonts";
import { toSfntBytes } from "../fonts/sfnt/woff";
import { writePdf } from "./writePdf";

const TIMESTAMP = "2026-01-02T03:04:05Z";
const FONT_SIZE_PX = 16;
const BASELINE_Y_PX = 100;
const RUN_X_PX = 72;
const POINTS_PER_PIXEL = 0.75;
const TEXT = "Fidelity";

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

const FACE: DisplayFontFace = {
  family: "Arimo",
  weight: 400,
  italic: false,
  generic: "sans-serif",
  fallbacks: [],
  fontBoxAscentRatio: 0.9,
  fontBoxDescentRatio: 0.2,
};

type Fixture = {
  readonly woff: Uint8Array;
  readonly font: SfntFont;
};

const loadFixture = async (): Promise<Fixture> => {
  const woff = await readTestFont({ family: "arimo" });
  const sfnt = toSfntBytes(woff);
  if (sfnt.isErr()) {
    throw sfnt.error;
  }
  const font = parseSfnt(sfnt.value);
  if (font.isErr()) {
    throw font.error;
  }
  return { woff, font: font.value };
};

/** The advance the face itself would use, in px, for one code point. */
const naturalAdvance = (font: SfntFont, codePoint: number): number =>
  (Math.round((font.advanceWidthFor(font.glyphIdFor(codePoint)) * 1000) / font.unitsPerEm) *
    FONT_SIZE_PX) /
  1000;

const listFor = (advances: readonly number[]): DisplayList => ({
  pages: [
    {
      pageNumber: 1,
      widthPx: 816,
      heightPx: 1056,
      orientation: "portrait",
      regions: [],
      primitives: [
        {
          kind: "glyphRun",
          font: 0,
          fontSizePx: FONT_SIZE_PX,
          color: BLACK,
          xPx: RUN_X_PX,
          baselineYPx: BASELINE_Y_PX,
          text: TEXT,
          advancesPx: advances,
          direction: "ltr",
          kerning: false,
          smallCaps: false,
        },
      ],
      links: [],
    },
  ],
  fonts: [FACE],
  images: [],
  outline: [],
  metadata: {},
  unsupported: [],
});

const mutool = Bun.which("mutool");

describe.skipIf(!TEST_FONTS_INSTALLED)(`embedded faces (${TEST_FONTS_SKIP_REASON})`, () => {
  test("embeds a subset as a CIDFontType2 with Identity-H", async () => {
    const { woff, font } = await loadFixture();
    const advances = [...TEXT].map((character) =>
      naturalAdvance(font, character.codePointAt(0) ?? 0),
    );
    const result = await writePdf(listFor(advances), {
      fonts: { load: () => [woff] },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.substitutions).toEqual([]);
    const text = latin1(result.value.bytes);
    expect(text).toContain("/Subtype /Type0");
    expect(text).toContain("/Encoding /Identity-H");
    expect(text).toContain("/Subtype /CIDFontType2");
    expect(text).toContain("/CIDToGIDMap /Identity");
    expect(text).toContain("/FontFile2 ");
    expect(text).toContain("/ToUnicode ");
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+/u);
    expect(text).toMatch(/\/W \[\d+ \[/u);
    // The subset is a fraction of the source face, not the whole binary.
    expect(result.value.bytes.length).toBeLessThan(woff.length);
  });

  test("the same list and timestamp embed byte for byte alike", async () => {
    const { woff, font } = await loadFixture();
    const advances = [...TEXT].map((character) =>
      naturalAdvance(font, character.codePointAt(0) ?? 0),
    );
    const source = { load: () => [woff] };
    const first = await writePdf(listFor(advances), { fonts: source, timestamp: TIMESTAMP });
    const second = await writePdf(listFor(advances), { fonts: source, timestamp: TIMESTAMP });
    if (first.isErr() || second.isErr()) {
      throw first.isErr() ? first.error : new Error("unreachable");
    }
    expect(Buffer.from(first.value.bytes).equals(Buffer.from(second.value.bytes))).toBe(true);
  });

  test("corrects the font's advances to the ones the display list carries", async () => {
    const { woff, font } = await loadFixture();
    const natural = [...TEXT].map((character) =>
      naturalAdvance(font, character.codePointAt(0) ?? 0),
    );
    const widened = natural.map((advance) => advance + 3);
    const result = await writePdf(listFor(widened), {
      fonts: { load: () => [woff] },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    const stream = contentStreamOf(result.value.bytes);
    expect(stream).not.toBeNull();
    const [, body = ""] = /\[(.*)\] TJ/u.exec(stream ?? "") ?? [];
    const adjustments = [...body.matchAll(/>(-?[\d.]+)/gu)].map(([, value]) => Number(value));
    expect(adjustments).toHaveLength(TEXT.length);
    // Each correction is exactly the widening, in thousandths of text space.
    for (const adjustment of adjustments) {
      expect(adjustment).toBeCloseTo((-3 * 1000) / FONT_SIZE_PX, 3);
    }
  });

  test("needs no correction when the list agrees with the face", async () => {
    const { woff, font } = await loadFixture();
    const advances = [...TEXT].map((character) =>
      naturalAdvance(font, character.codePointAt(0) ?? 0),
    );
    const result = await writePdf(listFor(advances), {
      fonts: { load: () => [woff] },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    const stream = contentStreamOf(result.value.bytes);
    // One unbroken hex string: every adjustment came out zero and was
    // dropped, which is only true if the two ways of computing the advance
    // agree exactly.
    expect(stream).toMatch(/\[<[0-9A-F]+>\] TJ/u);
  });
});

/** The page's content stream, found by the text run it must contain. */
const contentStreamOf = (bytes: Uint8Array): string | null => {
  const text = latin1(bytes);
  const pattern = /<<((?:(?!<<)[^])*?)\/Length (\d+)>>\nstream\n/g;
  let match = pattern.exec(text);
  while (match !== null) {
    const [, dict = "", length = "0"] = match;
    const start = match.index + match[0].length;
    if (dict.includes("/FlateDecode") && !dict.includes("/Length1")) {
      const inflated = latin1(inflateSync(bytes.subarray(start, start + Number(length))));
      if (inflated.includes(" TJ")) {
        return inflated;
      }
    }
    match = pattern.exec(text);
  }
  return null;
};

describe.skipIf(!TEST_FONTS_INSTALLED || mutool === null)("mutool with an embedded face", () => {
  test("extracts the run's text and places every glyph where the list says", async () => {
    const { woff, font } = await loadFixture();
    const advances = [...TEXT].map(
      (character) => naturalAdvance(font, character.codePointAt(0) ?? 0) + 3,
    );
    const result = await writePdf(listFor(advances), {
      fonts: { load: () => [woff] },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    const directory = mkdtempSync(join(tmpdir(), "folio-pdf-fonts-"));
    const file = join(directory, "out.pdf");
    await Bun.write(file, result.value.bytes);
    const process = Bun.spawn({
      cmd: [mutool ?? "mutool", "draw", "-F", "stext", "-o", join(directory, "out.xml"), file],
      stdout: "pipe",
      stderr: "pipe",
    });
    expect(await process.exited).toBe(0);
    const extracted = await Bun.file(join(directory, "out.xml")).text();
    // The extractor flags a character it synthesized: widening the advances
    // past the face's own puts a space between every glyph, and those are not
    // characters this file wrote.
    const SYNTHETIC_FLAG = 4;
    const characters = [...extracted.matchAll(/<char ([^>]*)\/>/gu)]
      .map(([, attributes = ""]) => ({
        character: /c="(.)"/u.exec(attributes)?.[1] ?? "",
        leftEdge: Number(/quad="([\d.-]+) /u.exec(attributes)?.[1] ?? "0"),
        flags: Number(/flags="(\d+)"/u.exec(attributes)?.[1] ?? "0"),
      }))
      .filter(({ flags }) => (flags & SYNTHETIC_FLAG) === 0);

    expect(characters.map(({ character }) => character).join("")).toBe(TEXT);
    let expectedX = RUN_X_PX;
    for (const [index, { leftEdge }] of characters.entries()) {
      expect(leftEdge).toBeCloseTo(expectedX * POINTS_PER_PIXEL, 1);
      expectedX += advances[index] ?? 0;
    }
  });
});
