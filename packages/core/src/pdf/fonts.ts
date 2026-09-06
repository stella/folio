/**
 * Font embedding.
 *
 * Every face the document paints becomes a composite (Type 0) font with
 * `Identity-H` encoding, so a run is written as two-byte glyph ids and the
 * file carries no encoding table to disagree with the font. A face the source
 * cannot supply, or whose licence bits forbid embedding, falls back to a
 * base-14 face and is *reported*: a substitution the caller cannot see is a
 * silently different document.
 *
 * ## A face is not a binary
 *
 * A packaged family arrives cut into disjoint script subsets: one binary
 * carries ASCII, another the Central European letters, and Czech, Slovak or
 * Polish text needs both at once. Each binary is its own sfnt with its own
 * glyph space, so one face becomes one PDF font resource *per binary the
 * document actually uses*, and a code point is served by the first binary
 * whose cmap covers it.
 */

import { panic, Result, TaggedError } from "better-result";
import type { DisplayFontFace } from "../display-list/types";
import { needsShaping, placeRun, type PlacedGlyph } from "../shaping/placeRun";
import type { Shaper, ShapingDirection } from "../shaping/shaper";
import { parseSfnt, type SfntFont } from "../fonts/sfnt/parse";
import { subsetTrueType } from "../fonts/sfnt/subset";
import { toSfntBytes } from "../fonts/sfnt/woff";
import {
  type PdfDocument,
  type PdfRef,
  type PdfValue,
  pdfArray,
  pdfAsciiString,
  pdfDict,
  pdfFlateStream,
  pdfName,
  pdfNumber,
  pdfNumberArray,
} from "./objects";
import type { PdfUnencodable } from "./writePdf";

class PdfFontError extends TaggedError("PdfFontError")<{ message: string }> {}

export type PdfFontSource = {
  /**
   * Every binary that carries part of this face, in priority order. A code
   * point is served by the first binary whose cmap covers it.
   */
  readonly load: (face: DisplayFontFace) => readonly Uint8Array[];
};

export type PdfSubstitution = {
  readonly family: string;
  readonly weight: number;
  readonly italic: boolean;
  readonly reason: string;
};

/** A code point placed in the glyph space of the resource that serves it. */
export type PdfGlyph = {
  /** The PDF font resource this glyph id belongs to. */
  readonly resourceIndex: number;
  readonly glyphId: number;
  /** The width this file declares for the glyph, in 1000ths of an em. */
  readonly widthUnits: number;
};

/**
 * One face, ready to paint. The two branches differ in how a code point
 * reaches the page, so the painter switches on `kind` rather than carrying
 * optional fields that are only valid in one of them.
 */
/** One shaped glyph, in the glyph space of the resource that will paint it. */
export type PlacedPdfGlyph = PdfGlyph & {
  /** Code-point index in the run's text of the cluster this glyph came from. */
  readonly clusterIndex: number;
  readonly xAdvancePx: number;
  readonly xOffsetPx: number;
  readonly yOffsetPx: number;
};

/** A run whose glyphs shaping has to choose, as the collection pass saw it. */
export type ShapedRunRequest = {
  readonly text: string;
  readonly direction: ShapingDirection;
  readonly fontSizePx: number;
};

/** A face embedded as a subset of its own program. */
export type EmbeddedPdfFont = {
  readonly kind: "embedded";
  /** Resource and subset glyph for a code point the collection pass saw. */
  readonly glyphFor: (codePoint: number) => PdfGlyph;
  /**
   * The glyphs shaping chose for a run, in visual order, or `null` when this
   * run was not shaped: no shaper was resolved, or the run is in a script where
   * a code point selects its own glyph.
   */
  readonly placeShapedRun: (request: ShapedRunRequest) => readonly PlacedPdfGlyph[] | null;
};

/** A base-14 stand-in for a face that could not be embedded. */
export type StandardPdfFont = {
  readonly kind: "standard";
  readonly resourceIndex: number;
  /** WinAnsi byte for a code point. */
  readonly byteFor: (codePoint: number) => number;
};

export type PreparedFont = EmbeddedPdfFont | StandardPdfFont;

/** Text space is 1000 units to the em whatever the font's own head is. */
const TEXT_SPACE_UNITS_PER_EM = 1000;

/** `fsType` bit 1: the vendor forbids embedding this face outright. */
const FSTYPE_RESTRICTED_LICENSE = 0x0002;

const FONT_FLAG_FIXED_PITCH = 1;
const FONT_FLAG_SERIF = 2;
const FONT_FLAG_SYMBOLIC = 4;
const FONT_FLAG_ITALIC = 64;

/**
 * `/StemV` has no counterpart in an sfnt table folio parses. A reader only
 * uses it to synthesize a stand-in when the embedded program is unusable, so
 * an approximation from the weight is honest and a wrong number is harmless.
 */
const STEM_V_REGULAR = 80;
const STEM_V_BOLD = 160;
const BOLD_WEIGHT_THRESHOLD = 600;

const NOTDEF_GLYPH = 0;

const SUBSET_TAG_LENGTH = 6;
const ALPHABET_LENGTH = 26;
const UPPERCASE_A = 65;

const WOFF_SIGNATURES = ["wOFF", "wOF2"] as const;

const readTag = (bytes: Uint8Array): string =>
  String.fromCharCode(bytes[0] ?? 0, bytes[1] ?? 0, bytes[2] ?? 0, bytes[3] ?? 0);

/**
 * Six uppercase letters derived from the face and the exact glyph set, so the
 * tag is stable across runs and two different subsets of one family do not
 * collide under the same name.
 */
const subsetTag = (seed: string): string => {
  const FNV_OFFSET = 0x811c9dc5;
  const FNV_PRIME = 0x01000193;
  let hash = FNV_OFFSET;
  for (let index = 0; index < seed.length; index += 1) {
    hash = Math.imul(hash ^ seed.charCodeAt(index), FNV_PRIME) >>> 0;
  }
  let tag = "";
  for (let index = 0; index < SUBSET_TAG_LENGTH; index += 1) {
    tag += String.fromCharCode(UPPERCASE_A + (hash % ALPHABET_LENGTH));
    hash = Math.floor(hash / ALPHABET_LENGTH);
  }
  return tag;
};

const BASE_14_ROMAN = {
  helvetica: "Helvetica",
  times: "Times-Roman",
  courier: "Courier",
} as const;

const BASE_14_BOLD = {
  helvetica: "Helvetica-Bold",
  times: "Times-Bold",
  courier: "Courier-Bold",
} as const;

const BASE_14_ITALIC = {
  helvetica: "Helvetica-Oblique",
  times: "Times-Italic",
  courier: "Courier-Oblique",
} as const;

const BASE_14_BOLD_ITALIC = {
  helvetica: "Helvetica-BoldOblique",
  times: "Times-BoldItalic",
  courier: "Courier-BoldOblique",
} as const;

type Base14Family = keyof typeof BASE_14_ROMAN;

/**
 * Total over the display list's generic categories: a new category is a
 * compile error here rather than a face that silently lands on Helvetica.
 */
const GENERIC_TO_BASE_14 = {
  serif: "times",
  "sans-serif": "helvetica",
  monospace: "courier",
  cursive: "times",
  fantasy: "helvetica",
} as const satisfies Record<DisplayFontFace["generic"], Base14Family>;

const base14Name = (face: DisplayFontFace): string => {
  const family = GENERIC_TO_BASE_14[face.generic];
  const bold = face.weight >= BOLD_WEIGHT_THRESHOLD;
  if (bold && face.italic) {
    return BASE_14_BOLD_ITALIC[family];
  }
  if (bold) {
    return BASE_14_BOLD[family];
  }
  return face.italic ? BASE_14_ITALIC[family] : BASE_14_ROMAN[family];
};

/**
 * The code points WinAnsiEncoding places in 0x80..0x9F, where it departs from
 * Latin-1. Outside that block WinAnsi and Latin-1 agree.
 */
const WIN_ANSI_HIGH_RANGE = [
  [0x80, 0x20ac],
  [0x82, 0x201a],
  [0x83, 0x0192],
  [0x84, 0x201e],
  [0x85, 0x2026],
  [0x86, 0x2020],
  [0x87, 0x2021],
  [0x88, 0x02c6],
  [0x89, 0x2030],
  [0x8a, 0x0160],
  [0x8b, 0x2039],
  [0x8c, 0x0152],
  [0x8e, 0x017d],
  [0x91, 0x2018],
  [0x92, 0x2019],
  [0x93, 0x201c],
  [0x94, 0x201d],
  [0x95, 0x2022],
  [0x96, 0x2013],
  [0x97, 0x2014],
  [0x98, 0x02dc],
  [0x99, 0x2122],
  [0x9a, 0x0161],
  [0x9b, 0x203a],
  [0x9c, 0x0153],
  [0x9e, 0x017e],
  [0x9f, 0x0178],
] as const satisfies readonly (readonly [number, number])[];

const WIN_ANSI_QUESTION_MARK = 0x3f;

const buildWinAnsiTable = (): ReadonlyMap<number, number> => {
  const table = new Map<number, number>();
  const PRINTABLE_ASCII_START = 0x20;
  const PRINTABLE_ASCII_END = 0x7e;
  const LATIN1_HIGH_START = 0xa0;
  const LATIN1_HIGH_END = 0xff;
  for (let code = PRINTABLE_ASCII_START; code <= PRINTABLE_ASCII_END; code += 1) {
    table.set(code, code);
  }
  for (let code = LATIN1_HIGH_START; code <= LATIN1_HIGH_END; code += 1) {
    table.set(code, code);
  }
  for (const [byte, codePoint] of WIN_ANSI_HIGH_RANGE) {
    table.set(codePoint, byte);
  }
  return table;
};

const WIN_ANSI_BY_CODE_POINT = buildWinAnsiTable();

const utf16BeHexOf = (text: string): string =>
  [...text].map((character) => utf16BeHex(character.codePointAt(0) ?? 0)).join("");

const utf16BeHex = (codePoint: number): string => {
  const SUPPLEMENTARY_START = 0x10000;
  const SURROGATE_HALF_BITS = 10;
  const LOW_SURROGATE_MASK = 0x3ff;
  const HIGH_SURROGATE_BASE = 0xd800;
  const LOW_SURROGATE_BASE = 0xdc00;
  const hex = (unit: number) => unit.toString(16).padStart(4, "0").toUpperCase();
  if (codePoint < SUPPLEMENTARY_START) {
    return hex(codePoint);
  }
  const offset = codePoint - SUPPLEMENTARY_START;
  return (
    hex(HIGH_SURROGATE_BASE + (offset >> SURROGATE_HALF_BITS)) +
    hex(LOW_SURROGATE_BASE + (offset & LOW_SURROGATE_MASK))
  );
};

/** Entries per `beginbfchar` block; the CMap syntax caps this at 100. */
const BFCHAR_BLOCK_SIZE = 100;

const toUnicodeCMap = (glyphToText: ReadonlyMap<number, string>): string => {
  const entries = [...glyphToText.entries()].sort(([left], [right]) => left - right);
  let body = "";
  for (let start = 0; start < entries.length; start += BFCHAR_BLOCK_SIZE) {
    const block = entries.slice(start, start + BFCHAR_BLOCK_SIZE);
    body += `${String(block.length)} beginbfchar\n`;
    for (const [glyphId, text] of block) {
      body += `<${utf16BeHex(glyphId)}> <${utf16BeHexOf(text)}>\n`;
    }
    body += "endbfchar\n";
  }
  return `/CIDInit /ProcSet findresource begin
12 dict begin
begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /Adobe-Identity-UCS def
/CMapType 2 def
1 begincodespacerange
<0000> <FFFF>
endcodespacerange
${body}endcmap
CMapName currentdict /CMap defineresource pop
end
end
`;
};

/**
 * `/W` as consecutive runs: `[ firstGid [w w w] nextGid [w] ]`. A subset's
 * glyph ids are dense, so this is normally one run.
 */
const widthsArray = (widths: ReadonlyMap<number, number>): PdfValue => {
  const items: PdfValue[] = [];
  let runStart: number | null = null;
  let expected = 0;
  let run: number[] = [];
  const flush = () => {
    if (runStart !== null && run.length > 0) {
      items.push(pdfNumber(runStart), pdfNumberArray(run));
    }
  };
  for (const glyphId of [...widths.keys()].sort((left, right) => left - right)) {
    if (runStart === null || glyphId !== expected) {
      flush();
      runStart = glyphId;
      run = [];
    }
    run.push(widths.get(glyphId) ?? panic("width run lost a glyph"));
    expected = glyphId + 1;
  }
  flush();
  return pdfArray(items);
};

const asciiOnly = (value: string): string => value.replace(/[^\x20-\x7e]/gu, "");

type DescriptorOptions = {
  readonly font: SfntFont;
  readonly face: DisplayFontFace;
  readonly fontName: string;
  readonly fontFileKey: "FontFile2" | "FontFile3";
  readonly fontFileRef: PdfRef;
};

const fontDescriptor = ({
  font,
  face,
  fontName,
  fontFileKey,
  fontFileRef,
}: DescriptorOptions): PdfValue => {
  const scale = TEXT_SPACE_UNITS_PER_EM / font.unitsPerEm;
  const italic = face.italic || font.italicAngle !== 0;
  const flags =
    // Identity encoding means the file declares no Latin character set, which
    // is what "symbolic" states to a reader.
    FONT_FLAG_SYMBOLIC |
    (face.generic === "serif" ? FONT_FLAG_SERIF : 0) |
    (face.generic === "monospace" ? FONT_FLAG_FIXED_PITCH : 0) |
    (italic ? FONT_FLAG_ITALIC : 0);
  return pdfDict([
    ["Type", pdfName("FontDescriptor")],
    ["FontName", pdfName(fontName)],
    ["Flags", pdfNumber(flags)],
    ["FontBBox", pdfNumberArray(font.bbox.map((value) => value * scale))],
    ["ItalicAngle", pdfNumber(font.italicAngle)],
    ["Ascent", pdfNumber(font.ascender * scale)],
    ["Descent", pdfNumber(font.descender * scale)],
    ["CapHeight", pdfNumber(font.capHeight * scale)],
    ["StemV", pdfNumber(face.weight >= BOLD_WEIGHT_THRESHOLD ? STEM_V_BOLD : STEM_V_REGULAR)],
    [fontFileKey, fontFileRef],
  ]);
};

type PlanOptions = {
  readonly font: SfntFont;
  readonly face: DisplayFontFace;
  readonly codePoints: readonly number[];
  /**
   * Glyphs shaping chose that no code point maps to: a lam-alef ligature, an
   * Arabic letter's medial form, a Devanagari conjunct. They have to be in the
   * subset before the first page is written, which is why shaping happens in
   * the same pass that collects code points.
   */
  readonly shapedGlyphIds?: ReadonlySet<number> | undefined;
  /** Source text behind each shaped glyph, so extraction still reads it. */
  readonly textByShapedGlyphId?: ReadonlyMap<number, string> | undefined;
};

/**
 * Everything one binary contributes to the file, computed before a single
 * object is allocated. Planning and emitting are separate because a face that
 * arrives as several binaries must either embed all of them or none: a
 * failure discovered halfway would otherwise leave the earlier binaries in
 * the file with nothing referring to them.
 */
type FontPlan = {
  readonly font: SfntFont;
  readonly face: DisplayFontFace;
  readonly fontName: string;
  readonly programBytes: Uint8Array;
  readonly glyphIdByCodePoint: ReadonlyMap<number, number>;
  readonly widthByGlyphId: ReadonlyMap<number, number>;
  readonly glyphToText: ReadonlyMap<number, string>;
  /** Source glyph id to the id it took in the subset. */
  readonly glyphIdMap: ReadonlyMap<number, number>;
};

const planFont = ({
  font,
  face,
  codePoints,
  shapedGlyphIds,
  textByShapedGlyphId,
}: PlanOptions): Result<FontPlan, PdfFontError> => {
  const sourceGlyphByCodePoint = new Map<number, number>();
  const sourceGlyphIds = new Set<number>([NOTDEF_GLYPH]);
  for (const codePoint of codePoints) {
    const glyphId = font.glyphIdFor(codePoint);
    sourceGlyphByCodePoint.set(codePoint, glyphId);
    sourceGlyphIds.add(glyphId);
  }
  for (const glyphId of shapedGlyphIds ?? []) {
    sourceGlyphIds.add(glyphId);
  }

  // A CFF outline table is not what `subsetTrueType` rewrites, so a CFF face
  // ships whole: every glyph of the family travels with the document even
  // when it paints three of them. Subsetting CFF means rebuilding the CFF
  // INDEX structures, which is a second subsetter this backend does not own.
  let glyphIdMap: ReadonlyMap<number, number>;
  let programBytes: Uint8Array;
  if (font.isCff) {
    glyphIdMap = new Map([...sourceGlyphIds].map((glyphId) => [glyphId, glyphId]));
    programBytes = font.bytes;
  } else {
    const subset = subsetTrueType(font, sourceGlyphIds);
    if (subset.isErr()) {
      return Result.err(
        new PdfFontError({ message: `subsetting failed: ${subset.error.message}` }),
      );
    }
    glyphIdMap = subset.value.glyphIdMap;
    programBytes = subset.value.bytes;
  }

  const scale = TEXT_SPACE_UNITS_PER_EM / font.unitsPerEm;
  const glyphIdByCodePoint = new Map<number, number>();
  const widthByGlyphId = new Map<number, number>();
  const glyphToText = new Map<number, string>();
  for (const sourceGlyphId of [...sourceGlyphIds].sort((left, right) => left - right)) {
    const targetGlyphId = glyphIdMap.get(sourceGlyphId);
    if (targetGlyphId === undefined) {
      return Result.err(
        new PdfFontError({ message: `subset dropped glyph ${String(sourceGlyphId)}` }),
      );
    }
    // Rounded here and nowhere else: the painter corrects against the width
    // this file declares, not against the font's own, so the rounding is
    // absorbed rather than accumulated.
    widthByGlyphId.set(targetGlyphId, Math.round(font.advanceWidthFor(sourceGlyphId) * scale));
  }
  for (const codePoint of [...sourceGlyphByCodePoint.keys()].sort((left, right) => left - right)) {
    const sourceGlyphId = sourceGlyphByCodePoint.get(codePoint) ?? NOTDEF_GLYPH;
    const targetGlyphId = glyphIdMap.get(sourceGlyphId) ?? NOTDEF_GLYPH;
    glyphIdByCodePoint.set(codePoint, targetGlyphId);
    if (!glyphToText.has(targetGlyphId)) {
      glyphToText.set(targetGlyphId, String.fromCodePoint(codePoint));
    }
  }
  // A glyph no code point maps to still has to extract as text, or a reader
  // copies a page of Arabic and gets nothing back for every ligature on it.
  for (const [sourceGlyphId, text] of textByShapedGlyphId ?? []) {
    const targetGlyphId = glyphIdMap.get(sourceGlyphId);
    if (targetGlyphId !== undefined && !glyphToText.has(targetGlyphId)) {
      glyphToText.set(targetGlyphId, text);
    }
  }

  const baseName = asciiOnly(font.postScriptName) || asciiOnly(face.family) || "Unknown";
  const fontName = `${subsetTag(`${baseName}:${[...sourceGlyphIds].sort((left, right) => left - right).join(",")}`)}+${baseName}`;

  return Result.ok({
    font,
    face,
    fontName,
    programBytes,
    glyphIdByCodePoint,
    widthByGlyphId,
    glyphToText,
    glyphIdMap,
  });
};

const emitFont = (document: PdfDocument, plan: FontPlan): PdfRef => {
  const { font, face, fontName, programBytes } = plan;
  const fontFileRef = document.add(
    font.isCff
      ? pdfFlateStream([["Subtype", pdfName("OpenType")]], programBytes)
      : pdfFlateStream([["Length1", pdfNumber(programBytes.length)]], programBytes),
  );
  const descriptorRef = document.add(
    fontDescriptor({
      font,
      face,
      fontName,
      fontFileKey: font.isCff ? "FontFile3" : "FontFile2",
      fontFileRef,
    }),
  );
  const toUnicodeRef = document.add(
    pdfFlateStream([], new TextEncoder().encode(toUnicodeCMap(plan.glyphToText))),
  );
  // A CFF-flavoured OpenType descends from a CIDFontType0; only a `glyf`
  // program is a CIDFontType2. Both keep the two-byte Identity-H codes, which
  // is what the encoding buys; only the descendant subtype differs, and a
  // reader that is handed the wrong one rejects the font outright.
  const descendantRef = document.add(
    pdfDict([
      ["Type", pdfName("Font")],
      ["Subtype", pdfName(font.isCff ? "CIDFontType0" : "CIDFontType2")],
      ["BaseFont", pdfName(fontName)],
      [
        "CIDSystemInfo",
        pdfDict([
          ["Registry", pdfAsciiString("Adobe")],
          ["Ordering", pdfAsciiString("Identity")],
          ["Supplement", pdfNumber(0)],
        ]),
      ],
      ["FontDescriptor", descriptorRef],
      ["DW", pdfNumber(TEXT_SPACE_UNITS_PER_EM)],
      ["W", widthsArray(plan.widthByGlyphId)],
      ["CIDToGIDMap", font.isCff ? undefined : pdfName("Identity")],
    ]),
  );
  return document.add(
    pdfDict([
      ["Type", pdfName("Font")],
      ["Subtype", pdfName("Type0")],
      ["BaseFont", pdfName(fontName)],
      ["Encoding", pdfName("Identity-H")],
      ["DescendantFonts", pdfArray([descendantRef])],
      ["ToUnicode", toUnicodeRef],
    ]),
  );
};

const standardFont = (document: PdfDocument, face: DisplayFontFace): PdfRef =>
  document.add(
    pdfDict([
      ["Type", pdfName("Font")],
      ["Subtype", pdfName("Type1")],
      ["BaseFont", pdfName(base14Name(face))],
      ["Encoding", pdfName("WinAnsiEncoding")],
    ]),
  );

/** Decoded and licence-checked binaries of one face, in priority order. */
type ParsedFace =
  | { readonly kind: "usable"; readonly fonts: readonly SfntFont[] }
  | { readonly kind: "unusable"; readonly reason: string };

const parseFace = (binaries: readonly Uint8Array[]): ParsedFace => {
  if (binaries.length === 0) {
    return { kind: "unusable", reason: "the font source supplied no bytes for this face" };
  }
  const fonts: SfntFont[] = [];
  let reason = "";
  for (const bytes of binaries) {
    let sfntBytes = bytes;
    if (WOFF_SIGNATURES.includes(readTag(bytes) as (typeof WOFF_SIGNATURES)[number])) {
      const decoded = toSfntBytes(bytes);
      if (decoded.isErr()) {
        reason = `WOFF decoding failed: ${decoded.error.message}`;
        continue;
      }
      sfntBytes = decoded.value;
    }
    const parsed = parseSfnt(sfntBytes);
    if (parsed.isErr()) {
      reason = `font parsing failed: ${parsed.error.message}`;
      continue;
    }
    if ((parsed.value.fsType & FSTYPE_RESTRICTED_LICENSE) !== 0) {
      reason = "the face's fsType bits forbid embedding";
      continue;
    }
    fonts.push(parsed.value);
  }
  // One unusable binary of several is not a substitution: the face still
  // paints from the rest, and the code points it carried are reported as
  // unencodable like any other gap in the face's coverage.
  return fonts.length === 0 ? { kind: "unusable", reason } : { kind: "usable", fonts };
};

/**
 * The binary that serves each code point, by index into the face's list, and
 * the code points no binary of the face covers.
 */
type FaceCoverage = {
  readonly codePointsByBinary: ReadonlyMap<number, readonly number[]>;
  readonly uncovered: readonly number[];
};

/** Where a code point no binary covers is painted from: `.notdef` is there. */
const FALLBACK_BINARY = 0;

/**
 * Which binary of the face serves a code point, or -1 when none does. One
 * answer for the whole file: coverage and shaping must agree about which
 * binary a character belongs to, or a run is shaped against one face and
 * painted from another's glyph space.
 */
const binaryIndexFor = (fonts: readonly SfntFont[], codePoint: number): number =>
  fonts.findIndex((font) => font.glyphIdFor(codePoint) !== NOTDEF_GLYPH);

const resolveCoverage = (
  fonts: readonly SfntFont[],
  codePoints: readonly number[],
): FaceCoverage => {
  const codePointsByBinary = new Map<number, number[]>();
  const uncovered: number[] = [];
  for (const codePoint of codePoints) {
    const found = binaryIndexFor(fonts, codePoint);
    if (found === -1) {
      uncovered.push(codePoint);
    }
    const binary = found === -1 ? FALLBACK_BINARY : found;
    const bucket = codePointsByBinary.get(binary) ?? [];
    bucket.push(codePoint);
    codePointsByBinary.set(binary, bucket);
  }
  return { codePointsByBinary, uncovered };
};

/** A stretch of one run's text served by a single binary of the face. */
type RunSegment = {
  readonly binaryIndex: number;
  /** Code-point index in the run's text where this stretch begins. */
  readonly startIndex: number;
  readonly text: string;
};

const segmentRun = (fonts: readonly SfntFont[], text: string): readonly RunSegment[] => {
  const segments: { binaryIndex: number; startIndex: number; text: string }[] = [];
  let open: { binaryIndex: number; startIndex: number; text: string } | null = null;
  let index = 0;
  for (const character of text) {
    const found = binaryIndexFor(fonts, character.codePointAt(0) ?? 0);
    const binaryIndex = found === -1 ? FALLBACK_BINARY : found;
    if (open === null || open.binaryIndex !== binaryIndex) {
      open = { binaryIndex, startIndex: index, text: character };
      segments.push(open);
    } else {
      open.text += character;
    }
    index += 1;
  }
  return segments;
};

/** One shaped stretch of a run, still in the source face's glyph space. */
type PlacedSegment = {
  readonly binaryIndex: number;
  readonly startIndex: number;
  readonly glyphs: readonly PlacedGlyph[];
};

/**
 * Everything shaping contributes to one face: which glyphs the subset must
 * carry, what text each of them came from, and where every run's glyphs go.
 *
 * Computed once, before a single object is written, and read again while
 * painting. Shaping twice would be two answers to one question, and the file
 * would be subset for one of them and painted from the other.
 */
type FaceShaping = {
  readonly glyphIdsByBinary: ReadonlyMap<number, ReadonlySet<number>>;
  readonly textByGlyphByBinary: ReadonlyMap<number, ReadonlyMap<number, string>>;
  readonly placementsByRun: ReadonlyMap<string, readonly PlacedSegment[]>;
};

const EMPTY_SHAPING: FaceShaping = {
  glyphIdsByBinary: new Map(),
  textByGlyphByBinary: new Map(),
  placementsByRun: new Map(),
};

const runKeyOf = ({ text, direction, fontSizePx }: ShapedRunRequest): string =>
  `${direction}\u0000${String(fontSizePx)}\u0000${text}`;

/**
 * The characters behind one glyph: the code points of its cluster, carried by
 * the first glyph of that cluster only. Giving every mark of a cluster the same
 * characters would repeat them on extraction.
 */
const clusterText = (
  characters: readonly string[],
  glyphs: readonly PlacedGlyph[],
  position: number,
): string => {
  const glyph = glyphs[position];
  if (glyph === undefined) {
    return "";
  }
  const first = glyphs.find((other) => other.clusterIndex === glyph.clusterIndex);
  if (first !== glyph) {
    return "";
  }
  const starts = [...new Set(glyphs.map((other) => other.clusterIndex))].sort(
    (left, right) => left - right,
  );
  const next = starts.find((start) => start > glyph.clusterIndex) ?? characters.length;
  return characters.slice(glyph.clusterIndex, next).join("");
};

const shapeFace = (
  fonts: readonly SfntFont[],
  runs: readonly ShapedRunRequest[],
  shaper: Shaper,
): FaceShaping => {
  const glyphIdsByBinary = new Map<number, Set<number>>();
  const textByGlyphByBinary = new Map<number, Map<number, string>>();
  const placementsByRun = new Map<string, readonly PlacedSegment[]>();

  for (const run of runs) {
    const key = runKeyOf(run);
    if (placementsByRun.has(key)) {
      continue;
    }
    const placed: PlacedSegment[] = [];
    for (const segment of segmentRun(fonts, run.text)) {
      const font = fonts[segment.binaryIndex];
      if (font === undefined || !needsShaping(segment.text)) {
        continue;
      }
      const glyphs = placeRun({
        shaper,
        font: font.bytes,
        text: segment.text,
        fontSizePx: run.fontSizePx,
        direction: run.direction,
      });
      placed.push({ binaryIndex: segment.binaryIndex, startIndex: segment.startIndex, glyphs });

      const ids = glyphIdsByBinary.get(segment.binaryIndex) ?? new Set<number>();
      const texts = textByGlyphByBinary.get(segment.binaryIndex) ?? new Map<number, string>();
      const characters = [...segment.text];
      for (const [position, glyph] of glyphs.entries()) {
        ids.add(glyph.glyphId);
        const text = clusterText(characters, glyphs, position);
        // A glyph that is not the first of its cluster contributes no
        // characters here: they belong to the glyph that opens the cluster.
        // It may still open one elsewhere in the document, so nothing is
        // recorded for it now rather than an empty entry that would block that.
        if (text !== "" && !texts.has(glyph.glyphId)) {
          texts.set(glyph.glyphId, text);
        }
      }
      glyphIdsByBinary.set(segment.binaryIndex, ids);
      textByGlyphByBinary.set(segment.binaryIndex, texts);
    }
    placementsByRun.set(key, placed);
  }
  return { glyphIdsByBinary, textByGlyphByBinary, placementsByRun };
};

type PrepareFontsOptions = {
  readonly document: PdfDocument;
  readonly faces: readonly DisplayFontFace[];
  /** Font index to the code points the document paints from that face. */
  readonly usedCodePoints: ReadonlyMap<number, ReadonlySet<number>>;
  /** Font index to the runs whose glyphs shaping has to choose. */
  readonly shapedRuns?: ReadonlyMap<number, readonly ShapedRunRequest[]>;
  /**
   * The shaper, when the document contains a run that needs one. Absent leaves
   * such runs painted one glyph per code point, which the scripts that shape do
   * not read as; the caller resolves a shaper whenever the display list has
   * one, which is what keeps the artifact off the path of a Latin document.
   */
  readonly shaper?: Shaper | null;
  readonly source: PdfFontSource;
};

export type PreparedFonts = {
  readonly byFontIndex: ReadonlyMap<number, PreparedFont>;
  /** Every PDF font resource the pages may name, by resource index. */
  readonly fontRefByResourceIndex: ReadonlyMap<number, PdfRef>;
  readonly substitutions: readonly PdfSubstitution[];
  readonly unencodable: readonly PdfUnencodable[];
};

/**
 * Turns the display list's font table into PDF font objects. A face nothing
 * paints is skipped entirely: an unused face is not a substitution, and
 * embedding it would put bytes in the file no page reads.
 */
export const prepareFonts = ({
  document,
  faces,
  usedCodePoints,
  shapedRuns,
  shaper = null,
  source,
}: PrepareFontsOptions): PreparedFonts => {
  const byFontIndex = new Map<number, PreparedFont>();
  const fontRefByResourceIndex = new Map<number, PdfRef>();
  const substitutions: PdfSubstitution[] = [];
  const unencodable: PdfUnencodable[] = [];
  // Resource indices run in face order, then in binary order within a face,
  // so the names a page assigns from its sorted keys never depend on the
  // order the painter happened to reach a run.
  let nextResourceIndex = 0;
  const claimResource = (ref: PdfRef): number => {
    const resourceIndex = nextResourceIndex;
    nextResourceIndex += 1;
    fontRefByResourceIndex.set(resourceIndex, ref);
    return resourceIndex;
  };

  // Two faces of one family can differ only in a field the report does not
  // carry, so the same gap must not be reported twice.
  const reported = new Set<string>();
  const report = (face: DisplayFontFace, codePoints: readonly number[]) => {
    for (const codePoint of codePoints) {
      const key = `${face.family} ${String(face.weight)} ${String(face.italic)} ${String(codePoint)}`;
      if (reported.has(key)) {
        continue;
      }
      reported.add(key);
      unencodable.push({
        codePoint,
        family: face.family,
        weight: face.weight,
        italic: face.italic,
      });
    }
  };

  const substitute = (
    fontIndex: number,
    face: DisplayFontFace,
    codePoints: readonly number[],
    reason: string,
  ) => {
    substitutions.push({
      family: face.family,
      weight: face.weight,
      italic: face.italic,
      reason,
    });
    // A base-14 stand-in carries no cmap this process can consult, so its
    // coverage is the encoding's: anything WinAnsi cannot name paints as `?`.
    report(
      face,
      codePoints.filter((codePoint) => !WIN_ANSI_BY_CODE_POINT.has(codePoint)),
    );
    byFontIndex.set(fontIndex, {
      kind: "standard",
      resourceIndex: claimResource(standardFont(document, face)),
      byteFor: (codePoint) => WIN_ANSI_BY_CODE_POINT.get(codePoint) ?? WIN_ANSI_QUESTION_MARK,
    });
  };

  for (const [fontIndex, face] of faces.entries()) {
    const used = usedCodePoints.get(fontIndex);
    if (used === undefined || used.size === 0) {
      continue;
    }
    const codePoints = [...used].sort((left, right) => left - right);
    // A face embedded in the source package wins over anything the host can
    // supply: its bytes are the ones the measurer took its advances from.
    const binaries = face.embedded === undefined ? source.load(face) : [face.embedded.bytes];
    const parsed = parseFace(binaries);
    if (parsed.kind === "unusable") {
      substitute(fontIndex, face, codePoints, parsed.reason);
      continue;
    }

    // Shaping first: the subset has to carry the glyphs it chooses, and those
    // are not reachable from any code point.
    const shaping =
      shaper === null
        ? EMPTY_SHAPING
        : shapeFace(parsed.fonts, shapedRuns?.get(fontIndex) ?? [], shaper);

    const coverage = resolveCoverage(parsed.fonts, codePoints);
    const binaryIndices = [
      ...new Set([...coverage.codePointsByBinary.keys(), ...shaping.glyphIdsByBinary.keys()]),
    ].sort((left, right) => left - right);
    const plans: FontPlan[] = [];
    let failure: PdfFontError | null = null;
    for (const binaryIndex of binaryIndices) {
      const plan = planFont({
        font: parsed.fonts[binaryIndex] ?? panic(`face lost binary ${String(binaryIndex)}`),
        face,
        codePoints: coverage.codePointsByBinary.get(binaryIndex) ?? [],
        shapedGlyphIds: shaping.glyphIdsByBinary.get(binaryIndex),
        textByShapedGlyphId: shaping.textByGlyphByBinary.get(binaryIndex),
      });
      if (plan.isErr()) {
        failure = plan.error;
        break;
      }
      plans.push(plan.value);
    }
    if (failure !== null) {
      substitute(fontIndex, face, codePoints, failure.message);
      continue;
    }

    const glyphByCodePoint = new Map<number, PdfGlyph>();
    const emitted = new Map<number, { plan: FontPlan; resourceIndex: number }>();
    for (const [position, plan] of plans.entries()) {
      const resourceIndex = claimResource(emitFont(document, plan));
      emitted.set(binaryIndices[position] ?? FALLBACK_BINARY, { plan, resourceIndex });
      for (const [codePoint, glyphId] of plan.glyphIdByCodePoint) {
        glyphByCodePoint.set(codePoint, {
          resourceIndex,
          glyphId,
          widthUnits:
            plan.widthByGlyphId.get(glyphId) ?? panic(`no width for glyph ${String(glyphId)}`),
        });
      }
    }
    report(face, coverage.uncovered);
    byFontIndex.set(fontIndex, {
      kind: "embedded",
      // A code point the collection pass did not see is a painter that walked
      // the display list differently from the collector: a defect, not a
      // missing glyph.
      glyphFor: (codePoint) =>
        glyphByCodePoint.get(codePoint) ??
        panic(`code point ${String(codePoint)} was painted but never collected`),
      placeShapedRun: (request) => {
        const segments = shaping.placementsByRun.get(runKeyOf(request));
        if (segments === undefined || segments.length === 0) {
          return null;
        }
        return segments.flatMap(({ binaryIndex, startIndex, glyphs }) => {
          const target =
            emitted.get(binaryIndex) ??
            panic(`shaped run used binary ${String(binaryIndex)}, which was never emitted`);
          return glyphs.map(({ glyphId, clusterIndex, xAdvancePx, xOffsetPx, yOffsetPx }) => {
            const subsetGlyphId =
              target.plan.glyphIdMap.get(glyphId) ??
              panic(`subset dropped shaped glyph ${String(glyphId)}`);
            return {
              resourceIndex: target.resourceIndex,
              glyphId: subsetGlyphId,
              widthUnits:
                target.plan.widthByGlyphId.get(subsetGlyphId) ??
                panic(`no width for glyph ${String(subsetGlyphId)}`),
              clusterIndex: startIndex + clusterIndex,
              xAdvancePx,
              xOffsetPx,
              yOffsetPx,
            };
          });
        });
      },
    });
  }

  // Code-unit order, not locale order: a report a caller diffs across two
  // machines must not depend on either machine's collation.
  const byCodeUnit = (left: string, right: string): number => {
    if (left === right) {
      return 0;
    }
    return left < right ? -1 : 1;
  };
  unencodable.sort(
    (left, right) =>
      byCodeUnit(left.family, right.family) ||
      left.weight - right.weight ||
      Number(left.italic) - Number(right.italic) ||
      left.codePoint - right.codePoint,
  );
  return { byFontIndex, fontRefByResourceIndex, substitutions, unencodable };
};
