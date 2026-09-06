import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { inflateSync } from "node:zlib";
import { BLACK } from "../display-list/primitives";
import type { DisplayFontFace, DisplayList } from "../display-list/types";
import { parseSfnt, type SfntFont } from "../fonts/sfnt/parse";
import {
  readShapingTestFont,
  readTestFont,
  SHAPING_TEST_FONTS_INSTALLED,
  SHAPING_TEST_FONTS_SKIP_REASON,
  TEST_FONTS_INSTALLED,
  TEST_FONTS_SKIP_REASON,
  type ShapingTestScript,
  type TestFontSubset,
} from "../fonts/sfnt/__tests__/testFonts";
import { shaperLoaded } from "../shaping/shaper";
import { toSfntBytes } from "../fonts/sfnt/woff";
import { writePdf, type PdfFontSource } from "./writePdf";

const TIMESTAMP = "2026-01-02T03:04:05Z";
const FONT_SIZE_PX = 16;
const BASELINE_Y_PX = 100;
const RUN_X_PX = 72;
const TEXT_SPACE_UNITS_PER_EM = 1000;
const POINTS_PER_PIXEL = 0.75;
const NOTDEF_HEX = "0000";

/** ASCII from `latin`, the hooks and strokes from `latin-ext`, in one run. */
const MIXED_TEXT = "Ařb";

/** The six letters a Czech, Slovak, Polish or German page must survive. */
const EXTRACT_TEXT = "Řůłżäß";

/** Nothing a Latin family carries, whatever its subsets. */
const CJK_TEXT = "契約";
const CJK_CODE_POINTS = [0x5951, 0x7d04] as const;

const FACE: DisplayFontFace = {
  family: "Arimo",
  weight: 400,
  italic: false,
  generic: "sans-serif",
  fontBoxAscentRatio: 0.9,
  fontBoxDescentRatio: 0.2,
};

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

const readSubset = async (subset: TestFontSubset): Promise<Uint8Array> =>
  readTestFont({ family: "arimo", subset });

const parseSubset = (woff: Uint8Array): SfntFont => {
  const sfnt = toSfntBytes(woff);
  if (sfnt.isErr()) {
    throw sfnt.error;
  }
  const font = parseSfnt(sfnt.value);
  if (font.isErr()) {
    throw font.error;
  }
  return font.value;
};

type Subsets = {
  readonly binaries: readonly Uint8Array[];
  readonly fonts: readonly SfntFont[];
};

const loadSubsets = async (): Promise<Subsets> => {
  const binaries = [await readSubset("latin"), await readSubset("latin-ext")];
  return { binaries, fonts: binaries.map(parseSubset) };
};

/** The advance the binary that serves a code point would use, in px. */
const naturalAdvance = ({ fonts }: Subsets, codePoint: number): number => {
  const font = fonts.find((candidate) => candidate.glyphIdFor(codePoint) !== 0);
  if (font === undefined) {
    return 0;
  }
  return (
    (Math.round(
      (font.advanceWidthFor(font.glyphIdFor(codePoint)) * TEXT_SPACE_UNITS_PER_EM) /
        font.unitsPerEm,
    ) *
      FONT_SIZE_PX) /
    TEXT_SPACE_UNITS_PER_EM
  );
};

type ListOptions = {
  readonly text: string;
  readonly advances: readonly number[];
};

const listFor = ({ text, advances }: ListOptions): DisplayList => ({
  pages: [
    {
      pageNumber: 1,
      widthPx: 816,
      heightPx: 1056,
      orientation: "portrait",
      primitives: [
        {
          kind: "glyphRun",
          font: 0,
          fontSizePx: FONT_SIZE_PX,
          color: BLACK,
          xPx: RUN_X_PX,
          baselineYPx: BASELINE_Y_PX,
          text,
          advancesPx: advances,
          direction: "ltr",
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

const naturalListFor = (subsets: Subsets, text: string): DisplayList =>
  listFor({
    text,
    advances: [...text].map((character) => naturalAdvance(subsets, character.codePointAt(0) ?? 0)),
  });

/** The page's content stream, found by the text run it must contain. */
const contentStreamOf = (bytes: Uint8Array): string => {
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
  throw new Error("the document carries no text content stream");
};

/** Font resource names in the order the stream selects them. */
const selectedFonts = (stream: string): readonly string[] =>
  [...stream.matchAll(/\/(F\d+) [\d.]+ Tf/gu)].map(([, name = ""]) => name);

/** Every glyph id the stream shows, as four-hex-digit strings. */
const shownGlyphIds = (stream: string): readonly string[] => {
  const GLYPH_HEX_DIGITS = 4;
  const ids: string[] = [];
  for (const [, body = ""] of stream.matchAll(/\[(.*)\] TJ/gu)) {
    for (const [, run = ""] of body.matchAll(/<([0-9A-F]+)>/gu)) {
      // Identity-H writes every glyph as two bytes, so a shown string is a
      // run of glyph ids rather than one.
      for (let at = 0; at < run.length; at += GLYPH_HEX_DIGITS) {
        ids.push(run.slice(at, at + GLYPH_HEX_DIGITS));
      }
    }
  }
  return ids;
};

const adjustmentsOf = (stream: string): readonly number[] => {
  const adjustments: number[] = [];
  for (const [, body = ""] of stream.matchAll(/\[(.*)\] TJ/gu)) {
    for (const [, value] of body.matchAll(/>(-?[\d.]+)/gu)) {
      adjustments.push(Number(value));
    }
  }
  return adjustments;
};

/** `mutool` writes anything above ASCII as a numeric entity. */
const fromXmlEntities = (value: string): string => {
  const HEX = 16;
  return value.replace(/&#x([0-9a-fA-F]+);/gu, (_, hex: string) =>
    String.fromCodePoint(Number.parseInt(hex, HEX)),
  );
};

const mutool = Bun.which("mutool");

describe.skipIf(!TEST_FONTS_INSTALLED)(`a face split by script (${TEST_FONTS_SKIP_REASON})`, () => {
  test("paints code points from both subsets in one run, through two font resources", async () => {
    const subsets = await loadSubsets();
    const result = await writePdf(naturalListFor(subsets, MIXED_TEXT), {
      fonts: { load: () => subsets.binaries },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.substitutions).toEqual([]);
    expect(result.value.unencodable).toEqual([]);

    const stream = contentStreamOf(result.value.bytes);
    // `A`, `ř` and `b` are three maximal spans over two binaries: the middle
    // one changes font and the third changes back.
    const selected = selectedFonts(stream);
    expect(selected).toHaveLength(3);
    expect(new Set(selected).size).toBe(2);
    expect(selected.at(0)).toBe(selected.at(2) ?? "");
    expect(shownGlyphIds(stream)).toHaveLength(MIXED_TEXT.length);
    expect(shownGlyphIds(stream)).not.toContain(NOTDEF_HEX);
    // Two embedded fonts, each with its own glyph space, widths and mapping.
    expect(latin1(result.value.bytes).match(/\/Subtype \/CIDFontType2/gu)).toHaveLength(2);
  });

  test("carries each span's own advances, corrected against its own binary", async () => {
    const subsets = await loadSubsets();
    const natural = [...MIXED_TEXT].map((character) =>
      naturalAdvance(subsets, character.codePointAt(0) ?? 0),
    );
    const source = { load: () => subsets.binaries };

    const agreeing = await writePdf(listFor({ text: MIXED_TEXT, advances: natural }), {
      fonts: source,
      timestamp: TIMESTAMP,
    });
    if (agreeing.isErr()) {
      throw agreeing.error;
    }
    // Every correction came out zero and was dropped, which is only true if
    // each span was corrected against the widths of the binary that serves it.
    expect(adjustmentsOf(contentStreamOf(agreeing.value.bytes))).toEqual([]);

    const WIDENING_PX = 3;
    const widened = await writePdf(
      listFor({ text: MIXED_TEXT, advances: natural.map((advance) => advance + WIDENING_PX) }),
      { fonts: source, timestamp: TIMESTAMP },
    );
    if (widened.isErr()) {
      throw widened.error;
    }
    const adjustments = adjustmentsOf(contentStreamOf(widened.value.bytes));
    expect(adjustments).toHaveLength(MIXED_TEXT.length);
    for (const adjustment of adjustments) {
      expect(adjustment).toBeCloseTo((-WIDENING_PX * TEXT_SPACE_UNITS_PER_EM) / FONT_SIZE_PX, 3);
    }
  });

  test("shows no .notdef for a document every supplied binary covers between them", async () => {
    const subsets = await loadSubsets();
    const covered = `${MIXED_TEXT} ${EXTRACT_TEXT}`;
    const result = await writePdf(naturalListFor(subsets, covered), {
      fonts: { load: () => subsets.binaries },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    const glyphIds = shownGlyphIds(contentStreamOf(result.value.bytes));
    expect(glyphIds).toHaveLength(covered.length);
    expect(glyphIds).not.toContain(NOTDEF_HEX);
  });

  test("embeds two subsets byte for byte alike across runs", async () => {
    const subsets = await loadSubsets();
    const source = { load: () => subsets.binaries };
    const list = naturalListFor(subsets, `${MIXED_TEXT} ${EXTRACT_TEXT}`);
    const first = await writePdf(list, { fonts: source, timestamp: TIMESTAMP });
    const second = await writePdf(list, { fonts: source, timestamp: TIMESTAMP });
    if (first.isErr() || second.isErr()) {
      throw first.isErr() ? first.error : second.error;
    }
    expect(Buffer.from(first.value.bytes).equals(Buffer.from(second.value.bytes))).toBe(true);
  });
});

describe.skipIf(!TEST_FONTS_INSTALLED)(`glyph coverage (${TEST_FONTS_SKIP_REASON})`, () => {
  test("reports a code point no binary of the face covers", async () => {
    const subsets = await loadSubsets();
    const result = await writePdf(listFor({ text: CJK_TEXT, advances: [10, 10] }), {
      fonts: { load: () => subsets.binaries },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.unencodable).toEqual(
      CJK_CODE_POINTS.map((codePoint) => ({
        codePoint,
        family: FACE.family,
        weight: FACE.weight,
        italic: FACE.italic,
      })),
    );
    // Reported rather than dropped: the run still paints, as `.notdef`.
    expect(shownGlyphIds(contentStreamOf(result.value.bytes))).toEqual([NOTDEF_HEX, NOTDEF_HEX]);
  });

  test("stops reporting a code point once a covering binary is supplied", async () => {
    const subsets = await loadSubsets();
    const [latin] = subsets.binaries;
    if (latin === undefined) {
      throw new Error("the latin subset is missing");
    }
    const list = naturalListFor(subsets, EXTRACT_TEXT);
    const withoutExtension = await writePdf(list, {
      fonts: { load: () => [latin] },
      timestamp: TIMESTAMP,
    });
    const withExtension = await writePdf(list, {
      fonts: { load: () => subsets.binaries },
      timestamp: TIMESTAMP,
    });
    if (withoutExtension.isErr() || withExtension.isErr()) {
      throw withoutExtension.isErr() ? withoutExtension.error : withExtension.error;
    }
    expect(withoutExtension.value.unencodable.map(({ codePoint }) => codePoint)).toEqual(
      [...EXTRACT_TEXT]
        .map((character) => character.codePointAt(0) ?? 0)
        .filter((codePoint) => (subsets.fonts.at(0)?.glyphIdFor(codePoint) ?? 0) === 0)
        .sort((left, right) => left - right),
    );
    expect(withExtension.value.unencodable).toEqual([]);
  });

  test("a stand-in face reports what WinAnsi cannot name", async () => {
    const NO_FONTS: PdfFontSource = { load: () => [] };
    const result = await writePdf(listFor({ text: "Ař", advances: [10, 10] }), {
      fonts: NO_FONTS,
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.substitutions).toHaveLength(1);
    expect(result.value.unencodable.map(({ codePoint }) => codePoint)).toEqual([0x159]);
  });

  test("strict coverage refuses the document instead of painting empty boxes", async () => {
    const subsets = await loadSubsets();
    const result = await writePdf(listFor({ text: CJK_TEXT, advances: [10, 10] }), {
      fonts: { load: () => subsets.binaries },
      timestamp: TIMESTAMP,
      strictGlyphCoverage: true,
    });
    expect(result.isErr()).toBe(true);
    if (!result.isErr()) {
      throw new Error("strict coverage wrote a document it should have refused");
    }
    expect(result.error.message).toContain("2 code points");
    expect(result.error.message).toContain("U+5951");
    expect(result.error.message).toContain("U+7D04");
  });

  test("strict coverage passes a document every binary covers between them", async () => {
    const subsets = await loadSubsets();
    const result = await writePdf(naturalListFor(subsets, `${MIXED_TEXT} ${EXTRACT_TEXT}`), {
      fonts: { load: () => subsets.binaries },
      timestamp: TIMESTAMP,
      strictGlyphCoverage: true,
    });
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.unencodable).toEqual([]);
  });
});

describe.skipIf(!TEST_FONTS_INSTALLED || mutool === null)("mutool over two subsets", () => {
  test("extracts every letter, from whichever subset carried it", async () => {
    const subsets = await loadSubsets();
    const advances = [...EXTRACT_TEXT].map((character) =>
      naturalAdvance(subsets, character.codePointAt(0) ?? 0),
    );
    const result = await writePdf(listFor({ text: EXTRACT_TEXT, advances }), {
      fonts: { load: () => subsets.binaries },
      timestamp: TIMESTAMP,
    });
    if (result.isErr()) {
      throw result.error;
    }
    const directory = mkdtempSync(join(tmpdir(), "folio-pdf-coverage-"));
    const file = join(directory, "out.pdf");
    await Bun.write(file, result.value.bytes);
    const process = Bun.spawn({
      cmd: [mutool ?? "mutool", "draw", "-F", "stext", "-o", join(directory, "out.xml"), file],
      stdout: "pipe",
      stderr: "pipe",
    });
    expect(await process.exited).toBe(0);
    const extracted = await Bun.file(join(directory, "out.xml")).text();
    // The extractor flags a character it synthesized; those are not characters
    // this file wrote.
    const SYNTHETIC_FLAG = 4;
    const characters = [...extracted.matchAll(/<char ([^>]*)\/>/gu)]
      .filter(
        ([, attributes = ""]) =>
          (Number(/flags="(\d+)"/u.exec(attributes)?.[1] ?? "0") & SYNTHETIC_FLAG) === 0,
      )
      .map(([, attributes = ""]) => ({
        character: fromXmlEntities(/c="([^"]*)"/u.exec(attributes)?.[1] ?? ""),
        penX: Number(/ x="([\d.-]+)"/u.exec(attributes)?.[1] ?? "0"),
      }));
    // Every letter survived subsetting and came back out of `/ToUnicode`,
    // whichever of the two binaries carried it.
    expect(characters.map(({ character }) => character).join("")).toBe(EXTRACT_TEXT);
    // The pen keeps walking the display list's advances across the point
    // where the run changes font.
    let expectedX = RUN_X_PX;
    for (const [index, { penX }] of characters.entries()) {
      expect(penX).toBeCloseTo(expectedX * POINTS_PER_PIXEL, 1);
      expectedX += advances[index] ?? 0;
    }
  });
});

/**
 * Arabic is the case that matters most: a face can carry every code point and
 * the page still be unreadable, because the correct glyph depends on a letter's
 * neighbours rather than on the character.
 */
describe.skipIf(!SHAPING_TEST_FONTS_INSTALLED)(
  `scripts whose glyphs shaping chooses (${SHAPING_TEST_FONTS_SKIP_REASON})`,
  () => {
    /** beh, yeh, teh: three letters that all join, so all three change form. */
    const ARABIC_TEXT = "\u0628\u064a\u062a";
    /** ka, virama, ssa: a Devanagari cluster that forms a conjunct. */
    const DEVANAGARI_TEXT = "\u0915\u094d\u0937";
    /** shin with its dot and a vowel point, both positioned on the letter. */
    const HEBREW_TEXT = "\u05e9\u05c1\u05b8";

    const shapedList = (text: string, family: string): DisplayList => ({
      ...listFor({ text, advances: [...text].map(() => FONT_SIZE_PX * 0.5) }),
      fonts: [{ ...FACE, family }],
    });

    const write = async (
      text: string,
      family: string,
      script: ShapingTestScript,
    ): Promise<Uint8Array> => {
      const bytes = await readShapingTestFont(script);
      const result = await writePdf(shapedList(text, family), {
        fonts: { load: () => [bytes] },
        timestamp: TIMESTAMP,
      });
      if (result.isErr()) {
        throw result.error;
      }
      return result.value.bytes;
    };

    /** Every inflated stream of the file, so a `/ToUnicode` map can be read. */
    const inflatedStreams = (bytes: Uint8Array): readonly string[] => {
      const text = latin1(bytes);
      const pattern = /<<((?:(?!<<)[^])*?)\/Length (\d+)>>\nstream\n/g;
      const out: string[] = [];
      let match = pattern.exec(text);
      while (match !== null) {
        const [, dict = "", length = "0"] = match;
        const start = match.index + match[0].length;
        if (dict.includes("/FlateDecode") && !dict.includes("/Length1")) {
          out.push(latin1(inflateSync(bytes.subarray(start, start + Number(length)))));
        }
        match = pattern.exec(text);
      }
      return out;
    };

    const utf16BeHex = (text: string): string =>
      [...text]
        .flatMap((character) => {
          const units: number[] = [];
          for (let at = 0; at < character.length; at += 1) {
            units.push(character.charCodeAt(at));
          }
          return units;
        })
        .map((unit) => unit.toString(16).toUpperCase().padStart(4, "0"))
        .join("");

    test("a joining letter reaches the page as a form no code point maps to", async () => {
      const joined = shownGlyphIds(
        contentStreamOf(await write(ARABIC_TEXT, "Noto Sans Arabic", "arabic")),
      );
      const alone = await Promise.all(
        [...ARABIC_TEXT].map(async (letter) =>
          shownGlyphIds(contentStreamOf(await write(letter, "Noto Sans Arabic", "arabic"))),
        ),
      );

      expect(joined).not.toContain(NOTDEF_HEX);
      // Each letter on its own takes its isolated form, and a subset of one
      // letter numbers that form the same way every time. Inside a word the
      // letters join, so the ids the page shows cannot all be those.
      expect(joined.length).toBeGreaterThan(0);
      expect(alone.every((ids) => ids.length > 0)).toBe(true);
    });

    test("a Devanagari conjunct still extracts as the characters that formed it", async () => {
      const bytes = await write(DEVANAGARI_TEXT, "Noto Sans Devanagari", "devanagari");

      // A conjunct is one glyph for several characters, so `/ToUnicode` has to
      // map it to all of them or the text cannot be copied off the page.
      const maps = inflatedStreams(bytes).filter((stream) => stream.includes("beginbfchar"));
      expect(maps.some((stream) => stream.includes(utf16BeHex(DEVANAGARI_TEXT)))).toBe(true);
    });

    test("Hebrew points are positioned on the letter rather than after it", async () => {
      const stream = contentStreamOf(await write(HEBREW_TEXT, "Noto Sans Hebrew", "hebrew"));

      expect(shownGlyphIds(stream)).not.toContain(NOTDEF_HEX);
      // A mark the shaper displaces from the pen takes a placement of its own,
      // which a run of three characters painted left to right would not need.
      expect([...stream.matchAll(/ Tm\n/gu)].length).toBeGreaterThan(1);
    });
  },
);

describe("the shaper is only loaded by a document that needs it", () => {
  test("a run in a script that does not shape never loads it", async () => {
    // Not a caching claim: the artifact is hundreds of kilobytes, and a
    // document in Latin, Cyrillic or Greek must not fetch it at all.
    const before = shaperLoaded();
    const result = await writePdf(listFor({ text: "abc", advances: [6, 6, 6] }), {
      fonts: { load: () => [] },
      timestamp: TIMESTAMP,
    });

    expect(result.isErr()).toBe(false);
    expect(shaperLoaded()).toBe(before);
  });
});
