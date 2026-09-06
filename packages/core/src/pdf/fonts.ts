/**
 * Font embedding.
 *
 * Every face the document paints becomes a composite (Type 0) font with
 * `Identity-H` encoding, so a run is written as two-byte glyph ids and the
 * file carries no encoding table to disagree with the font. A face the source
 * cannot supply, or whose licence bits forbid embedding, falls back to a
 * base-14 face and is *reported*: a substitution the caller cannot see is a
 * silently different document.
 */

import { panic, Result, TaggedError } from "better-result";
import type { DisplayFontFace } from "../display-list/types";
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

class PdfFontError extends TaggedError("PdfFontError")<{ message: string }> {}

export type PdfFontSource = {
  /**
   * Font bytes for one face of the display list's font table. Returning null
   * is a reported substitution, never a silent one.
   */
  readonly load: (face: DisplayFontFace) => Uint8Array | null;
};

export type PdfSubstitution = {
  readonly family: string;
  readonly weight: number;
  readonly italic: boolean;
  readonly reason: string;
};

/**
 * One face, ready to paint. The two branches differ in how a code point
 * reaches the page, so the painter switches on `kind` rather than carrying
 * optional fields that are only valid in one of them.
 */
export type PreparedFont =
  | {
      readonly kind: "embedded";
      readonly ref: PdfRef;
      /** Subset glyph id for a code point the collection pass saw. */
      readonly glyphIdFor: (codePoint: number) => number;
      /** The width this file declares for the glyph, in 1000ths of an em. */
      readonly widthFor: (glyphId: number) => number;
    }
  | {
      readonly kind: "standard";
      readonly ref: PdfRef;
      /** WinAnsi byte for a code point. */
      readonly byteFor: (codePoint: number) => number;
    };

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

const toUnicodeCMap = (glyphToCodePoint: ReadonlyMap<number, number>): string => {
  const entries = [...glyphToCodePoint.entries()].sort(([left], [right]) => left - right);
  let body = "";
  for (let start = 0; start < entries.length; start += BFCHAR_BLOCK_SIZE) {
    const block = entries.slice(start, start + BFCHAR_BLOCK_SIZE);
    body += `${String(block.length)} beginbfchar\n`;
    for (const [glyphId, codePoint] of block) {
      body += `<${utf16BeHex(glyphId)}> <${utf16BeHex(codePoint)}>\n`;
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

type EmbedOptions = {
  readonly document: PdfDocument;
  readonly font: SfntFont;
  readonly face: DisplayFontFace;
  readonly codePoints: readonly number[];
};

type EmbedResult = {
  readonly ref: PdfRef;
  readonly glyphIdByCodePoint: ReadonlyMap<number, number>;
  readonly widthByGlyphId: ReadonlyMap<number, number>;
};

const embedFont = ({
  document,
  font,
  face,
  codePoints,
}: EmbedOptions): Result<EmbedResult, PdfFontError> => {
  const NOTDEF_GLYPH = 0;
  const sourceGlyphByCodePoint = new Map<number, number>();
  const sourceGlyphIds = new Set<number>([NOTDEF_GLYPH]);
  for (const codePoint of codePoints) {
    const glyphId = font.glyphIdFor(codePoint);
    sourceGlyphByCodePoint.set(codePoint, glyphId);
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
  const glyphToCodePoint = new Map<number, number>();
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
    if (!glyphToCodePoint.has(targetGlyphId)) {
      glyphToCodePoint.set(targetGlyphId, codePoint);
    }
  }

  const baseName = asciiOnly(font.postScriptName) || asciiOnly(face.family) || "Unknown";
  const fontName = `${subsetTag(`${baseName}:${[...sourceGlyphIds].sort((left, right) => left - right).join(",")}`)}+${baseName}`;

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
    pdfFlateStream([], new TextEncoder().encode(toUnicodeCMap(glyphToCodePoint))),
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
      ["W", widthsArray(widthByGlyphId)],
      ["CIDToGIDMap", font.isCff ? undefined : pdfName("Identity")],
    ]),
  );
  const ref = document.add(
    pdfDict([
      ["Type", pdfName("Font")],
      ["Subtype", pdfName("Type0")],
      ["BaseFont", pdfName(fontName)],
      ["Encoding", pdfName("Identity-H")],
      ["DescendantFonts", pdfArray([descendantRef])],
      ["ToUnicode", toUnicodeRef],
    ]),
  );
  return Result.ok({ ref, glyphIdByCodePoint, widthByGlyphId });
};

const standardFont = (document: PdfDocument, face: DisplayFontFace): PreparedFont => {
  const ref = document.add(
    pdfDict([
      ["Type", pdfName("Font")],
      ["Subtype", pdfName("Type1")],
      ["BaseFont", pdfName(base14Name(face))],
      ["Encoding", pdfName("WinAnsiEncoding")],
    ]),
  );
  return {
    kind: "standard",
    ref,
    byteFor: (codePoint) => WIN_ANSI_BY_CODE_POINT.get(codePoint) ?? WIN_ANSI_QUESTION_MARK,
  };
};

type PrepareFontsOptions = {
  readonly document: PdfDocument;
  readonly faces: readonly DisplayFontFace[];
  /** Font index to the code points the document paints from that face. */
  readonly usedCodePoints: ReadonlyMap<number, ReadonlySet<number>>;
  readonly source: PdfFontSource;
};

export type PreparedFonts = {
  readonly byFontIndex: ReadonlyMap<number, PreparedFont>;
  readonly substitutions: readonly PdfSubstitution[];
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
  source,
}: PrepareFontsOptions): PreparedFonts => {
  const byFontIndex = new Map<number, PreparedFont>();
  const substitutions: PdfSubstitution[] = [];

  const substitute = (fontIndex: number, face: DisplayFontFace, reason: string) => {
    substitutions.push({
      family: face.family,
      weight: face.weight,
      italic: face.italic,
      reason,
    });
    byFontIndex.set(fontIndex, standardFont(document, face));
  };

  for (const [fontIndex, face] of faces.entries()) {
    const codePoints = usedCodePoints.get(fontIndex);
    if (codePoints === undefined || codePoints.size === 0) {
      continue;
    }
    // A face embedded in the source package wins over anything the host can
    // supply: its bytes are the ones the measurer took its advances from.
    const bytes = face.embedded?.bytes ?? source.load(face);
    if (bytes === null || bytes === undefined) {
      substitute(fontIndex, face, "the font source supplied no bytes for this face");
      continue;
    }

    let sfntBytes = bytes;
    if (WOFF_SIGNATURES.includes(readTag(bytes) as (typeof WOFF_SIGNATURES)[number])) {
      const decoded = toSfntBytes(bytes);
      if (decoded.isErr()) {
        substitute(fontIndex, face, `WOFF decoding failed: ${decoded.error.message}`);
        continue;
      }
      sfntBytes = decoded.value;
    }

    const parsed = parseSfnt(sfntBytes);
    if (parsed.isErr()) {
      substitute(fontIndex, face, `font parsing failed: ${parsed.error.message}`);
      continue;
    }
    const font = parsed.value;
    if ((font.fsType & FSTYPE_RESTRICTED_LICENSE) !== 0) {
      substitute(fontIndex, face, "the face's fsType bits forbid embedding");
      continue;
    }

    const embedded = embedFont({
      document,
      font,
      face,
      codePoints: [...codePoints].sort((left, right) => left - right),
    });
    if (embedded.isErr()) {
      substitute(fontIndex, face, embedded.error.message);
      continue;
    }
    const { ref, glyphIdByCodePoint, widthByGlyphId } = embedded.value;
    byFontIndex.set(fontIndex, {
      kind: "embedded",
      ref,
      // A code point the collection pass did not see is a painter that walked
      // the display list differently from the collector: a defect, not a
      // missing glyph.
      glyphIdFor: (codePoint) =>
        glyphIdByCodePoint.get(codePoint) ??
        panic(`code point ${String(codePoint)} was painted but never collected`),
      widthFor: (glyphId) =>
        widthByGlyphId.get(glyphId) ?? panic(`no width for glyph ${String(glyphId)}`),
    });
  }

  return { byFontIndex, substitutions };
};
