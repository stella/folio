/**
 * A self-contained TrueType/OpenType (sfnt) reader.
 *
 * It reads exactly what text measurement and PDF embedding need: the metrics
 * head/hhea/OS-2 carry, a character-to-glyph lookup, and advance widths. It
 * does not decode outlines (the subsetter copies `glyf` bytes verbatim) and it
 * does not hint, shape or rasterize.
 *
 * Robustness contract: every offset and length is checked against the buffer
 * before it is read, so a truncated or lying font yields an `Err` rather than
 * a throw, and no field is ever `NaN`.
 */

import { Result, TaggedError } from "better-result";

import {
  fitsInView,
  readLocaOffsets,
  readSfntDirectory,
  readSfntVersion,
  SFNT_TABLE,
  SFNT_VERSION,
  type SfntDirectory,
  type SfntTableRecord,
} from "./tables";

export type SfntFont = {
  readonly unitsPerEm: number;
  /** hhea ascender/descender/lineGap, in font units; descender is negative. */
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
  /** OS/2 sCapHeight and sxHeight when version >= 2, else derived from glyph bounds. */
  readonly capHeight: number;
  readonly xHeight: number;
  /** post table italicAngle, degrees, negative for forward slant. */
  readonly italicAngle: number;
  /** head glyph bounding box: [xMin, yMin, xMax, yMax] in font units. */
  readonly bbox: readonly [number, number, number, number];
  readonly postScriptName: string;
  /** true when outlines live in a `CFF ` table rather than `glyf`. */
  readonly isCff: boolean;
  readonly numGlyphs: number;
  /** OS/2 fsType embedding permission bits, for the caller to police. */
  readonly fsType: number;
  /** The decoded sfnt bytes this font was parsed from. */
  readonly bytes: Uint8Array;
  /** Best-available cmap lookup. Returns 0 (.notdef) when unmapped. */
  readonly glyphIdFor: (codePoint: number) => number;
  /** hmtx advance in font units. Clamped to the last entry, per spec. */
  readonly advanceWidthFor: (glyphId: number) => number;
  /**
   * Ink bounding box of one glyph in font units, or null when the glyph is
   * empty (a space) or has no outline this reader can measure.
   *
   * This is what a canvas backend reports as `actualBoundingBox*`, and it is
   * not the hhea metrics: a headless measure provider needs the ink extent to
   * place baselines the way canvas does.
   */
  readonly glyphBoundsFor: (glyphId: number) => SfntGlyphBounds | null;
};

/** Ink extent of a single glyph, in font units. */
export type SfntGlyphBounds = {
  readonly xMin: number;
  readonly yMin: number;
  readonly xMax: number;
  readonly yMax: number;
};

export class SfntParseError extends TaggedError("SfntParseError")<{ message: string }> {}

const fail = (message: string) => Result.err(new SfntParseError({ message }));

/** Byte offsets inside `head`. */
const HEAD = {
  checkSumAdjustment: 8,
  magicNumber: 12,
  unitsPerEm: 18,
  xMin: 36,
  yMin: 38,
  xMax: 40,
  yMax: 42,
  indexToLocFormat: 50,
  size: 54,
} as const;

/** `head.magicNumber`, the one field that proves the table is really `head`. */
const HEAD_MAGIC = 0x5f0f3cf5;
const MIN_UNITS_PER_EM = 16;
const MAX_UNITS_PER_EM = 16384;

/** Byte offsets inside `hhea`. */
const HHEA = {
  ascender: 4,
  descender: 6,
  lineGap: 8,
  numberOfHMetrics: 34,
  size: 36,
} as const;

/** Byte offsets inside `maxp` (only the half both versions share). */
const MAXP = { numGlyphs: 4, size: 6 } as const;

/** Byte offsets inside `OS/2`. */
const OS2 = {
  version: 0,
  fsType: 8,
  sxHeight: 86,
  sCapHeight: 88,
  /** Version 2 is the first that has sxHeight/sCapHeight at all. */
  minimumVersionWithHeights: 2,
  sizeWithHeights: 96,
} as const;

/** Byte offsets inside `post`. */
const POST = { italicAngle: 4, size: 32 } as const;

/** 16.16 fixed-point, as `post.italicAngle` and table versions are stored. */
const FIXED_POINT_DIVISOR = 65536;

/** One hmtx `longHorMetric`: advanceWidth (uint16) then lsb (int16). */
const HMTX_METRIC_SIZE = 4;

/** Byte offsets inside a glyph description header. */
const GLYPH_HEADER = { numberOfContours: 0, xMin: 2, yMin: 4, xMax: 6, yMax: 8, size: 10 } as const;

/**
 * Cap and x-height fallbacks for fonts without OS/2 version 2 whose 'H' and
 * 'x' cannot be measured (no cmap, no outline, or CFF outlines this reader
 * does not decode). Ratios of the em, rounded, so measurement stays sane.
 */
const CAP_HEIGHT_EM_RATIO = 0.7;
const X_HEIGHT_EM_RATIO = 0.5;
const CAP_HEIGHT_SAMPLE = 0x48; // 'H'
const X_HEIGHT_SAMPLE = 0x78; // 'x'

const NAME = {
  count: 2,
  stringOffset: 4,
  recordsStart: 6,
  recordSize: 12,
  size: 6,
} as const;

const NAME_RECORD = {
  platformId: 0,
  encodingId: 2,
  languageId: 4,
  nameId: 6,
  length: 8,
  offset: 10,
} as const;

/** name id 6 is the PostScript name. */
const POSTSCRIPT_NAME_ID = 6;

const PLATFORM = { unicode: 0, macintosh: 1, windows: 3 } as const;

const WINDOWS_ENCODING = { symbol: 0, bmp: 1, full: 10 } as const;

const CMAP = { numTables: 2, recordsStart: 4, recordSize: 8 } as const;

const CMAP_RECORD = { platformId: 0, encodingId: 2, offset: 4 } as const;

const CMAP_FORMAT = {
  /** Segment mapping to delta values: the BMP workhorse. */
  segmentMapping: 4,
  /** Trimmed table mapping: a single contiguous run. */
  trimmedTable: 6,
  /** Segmented coverage: the full Unicode range, uint32 code points. */
  segmentedCoverage: 12,
} as const;

/** Byte offsets inside a format 4 subtable. */
const FORMAT_4 = { segCountX2: 6, endCodes: 14, headerSize: 14 } as const;

/** Byte offsets inside a format 6 subtable. */
const FORMAT_6 = { firstCode: 6, entryCount: 8, glyphIds: 10, headerSize: 10 } as const;

/** Byte offsets inside a format 12 subtable. */
const FORMAT_12 = { numGroups: 12, groups: 16, headerSize: 16, groupSize: 12 } as const;

/**
 * Symbol fonts (platform 3, encoding 0) map their glyphs into the private use
 * area, so a lookup for 'A' has to be retried as `0xF000 | 'A'`.
 */
const SYMBOL_CODEPOINT_BASE = 0xf000;

/**
 * Ceiling on entries materialized from a format 6/12 subtable. Glyph ids are
 * uint16, so an honest font cannot map more code points than this; a font that
 * claims to is lying about its group ranges and must not be allowed to
 * allocate against it.
 */
const MAX_CMAP_MAPPINGS = 1 << 18;

const MAX_UNICODE_CODE_POINT = 0x10ffff;

type CmapLookup = (codePoint: number) => number;

/** A cmap subtable located in the buffer, with the format already read. */
type CmapSubtable = {
  readonly platformId: number;
  readonly encodingId: number;
  readonly format: number;
  readonly start: number;
  /** End of the enclosing cmap table: the hard limit for any read. */
  readonly end: number;
  readonly view: DataView;
};

/** Reads `count` big-endian uint16s into a typed array. Bounds pre-checked. */
const readUint16Array = (view: DataView, offset: number, count: number): Uint16Array => {
  const values = new Uint16Array(count);
  for (let index = 0; index < count; index++) {
    values[index] = view.getUint16(offset + index * 2);
  }
  return values;
};

const readInt16Array = (view: DataView, offset: number, count: number): Int16Array => {
  const values = new Int16Array(count);
  for (let index = 0; index < count; index++) {
    values[index] = view.getInt16(offset + index * 2);
  }
  return values;
};

/**
 * Format 4: four parallel segment arrays plus a trailing glyph id array that
 * segments address by a byte offset from their own `idRangeOffset` slot. The
 * arrays are read once here; lookup is a binary search over `endCodes`.
 */
const buildFormat4Lookup = ({ view, start, end }: CmapSubtable): CmapLookup | undefined => {
  if (!fitsInView(view, start, FORMAT_4.headerSize) || start + FORMAT_4.headerSize > end) {
    return undefined;
  }

  const segCountX2 = view.getUint16(start + FORMAT_4.segCountX2);
  const segCount = segCountX2 >> 1;
  if (segCount === 0) {
    return undefined;
  }

  const endCodesBase = start + FORMAT_4.endCodes;
  // The spec inserts a reserved uint16 between endCode[] and startCode[].
  const startCodesBase = endCodesBase + segCountX2 + 2;
  const idDeltasBase = startCodesBase + segCountX2;
  const idRangeOffsetsBase = idDeltasBase + segCountX2;
  const arraysSize = segCountX2 * 4 + 2;
  if (!fitsInView(view, endCodesBase, arraysSize) || endCodesBase + arraysSize > end) {
    return undefined;
  }

  const endCodes = readUint16Array(view, endCodesBase, segCount);
  const startCodes = readUint16Array(view, startCodesBase, segCount);
  const idDeltas = readInt16Array(view, idDeltasBase, segCount);
  const idRangeOffsets = readUint16Array(view, idRangeOffsetsBase, segCount);

  return (codePoint) => {
    if (codePoint < 0 || codePoint > 0xffff) {
      return 0;
    }

    let low = 0;
    let high = segCount - 1;
    let segment = -1;
    while (low <= high) {
      const middle = (low + high) >> 1;
      if ((endCodes[middle] ?? 0) < codePoint) {
        low = middle + 1;
      } else {
        segment = middle;
        high = middle - 1;
      }
    }
    if (segment < 0) {
      return 0;
    }

    const segmentStart = startCodes[segment] ?? 0;
    if (segmentStart > codePoint) {
      return 0;
    }

    const delta = idDeltas[segment] ?? 0;
    const rangeOffset = idRangeOffsets[segment] ?? 0;
    if (rangeOffset === 0) {
      return (codePoint + delta) & 0xffff;
    }

    const address = idRangeOffsetsBase + segment * 2 + rangeOffset + (codePoint - segmentStart) * 2;
    if (!fitsInView(view, address, 2) || address + 2 > end) {
      return 0;
    }
    const glyphId = view.getUint16(address);
    return glyphId === 0 ? 0 : (glyphId + delta) & 0xffff;
  };
};

/** Format 6: one contiguous run of glyph ids, materialized eagerly. */
const buildFormat6Lookup = ({ view, start, end }: CmapSubtable): CmapLookup | undefined => {
  if (!fitsInView(view, start, FORMAT_6.headerSize) || start + FORMAT_6.headerSize > end) {
    return undefined;
  }

  const firstCode = view.getUint16(start + FORMAT_6.firstCode);
  const entryCount = view.getUint16(start + FORMAT_6.entryCount);
  const glyphIdsBase = start + FORMAT_6.glyphIds;
  if (!fitsInView(view, glyphIdsBase, entryCount * 2) || glyphIdsBase + entryCount * 2 > end) {
    return undefined;
  }

  const mappings = new Map<number, number>();
  for (let index = 0; index < entryCount; index++) {
    const glyphId = view.getUint16(glyphIdsBase + index * 2);
    if (glyphId !== 0) {
      mappings.set(firstCode + index, glyphId);
    }
  }

  return (codePoint) => mappings.get(codePoint) ?? 0;
};

/**
 * Format 12: sorted `(startCharCode, endCharCode, startGlyphID)` groups. Every
 * group is expanded into the map up to {@link MAX_CMAP_MAPPINGS}, so lookup is
 * a hash hit rather than a per-character binary search.
 */
const buildFormat12Lookup = ({ view, start, end }: CmapSubtable): CmapLookup | undefined => {
  if (!fitsInView(view, start, FORMAT_12.headerSize) || start + FORMAT_12.headerSize > end) {
    return undefined;
  }

  const numGroups = view.getUint32(start + FORMAT_12.numGroups);
  const groupsBase = start + FORMAT_12.groups;
  const groupsSize = numGroups * FORMAT_12.groupSize;
  if (!fitsInView(view, groupsBase, groupsSize) || groupsBase + groupsSize > end) {
    return undefined;
  }

  const mappings = new Map<number, number>();
  for (let index = 0; index < numGroups; index++) {
    const group = groupsBase + index * FORMAT_12.groupSize;
    const startCharCode = view.getUint32(group);
    const endCharCode = view.getUint32(group + 4);
    const startGlyphId = view.getUint32(group + 8);
    if (endCharCode < startCharCode || endCharCode > MAX_UNICODE_CODE_POINT) {
      return undefined;
    }
    if (mappings.size + (endCharCode - startCharCode + 1) > MAX_CMAP_MAPPINGS) {
      return undefined;
    }

    for (let codePoint = startCharCode; codePoint <= endCharCode; codePoint++) {
      const glyphId = startGlyphId + (codePoint - startCharCode);
      if (glyphId !== 0) {
        mappings.set(codePoint, glyphId);
      }
    }
  }

  return (codePoint) => mappings.get(codePoint) ?? 0;
};

const buildLookup = (subtable: CmapSubtable): CmapLookup | undefined => {
  switch (subtable.format) {
    case CMAP_FORMAT.segmentMapping:
      return buildFormat4Lookup(subtable);
    case CMAP_FORMAT.trimmedTable:
      return buildFormat6Lookup(subtable);
    case CMAP_FORMAT.segmentedCoverage:
      return buildFormat12Lookup(subtable);
    default:
      return undefined;
  }
};

/**
 * Subtable preference, most specific first. Order is the whole contract here:
 * a full-coverage Windows table beats the BMP one, a Unicode-platform table
 * beats the symbol encoding, and the symbol encoding is the only one whose
 * lookup retries in the private use area.
 */
const CMAP_PREFERENCES = [
  {
    matches: (subtable: CmapSubtable) =>
      subtable.platformId === PLATFORM.windows &&
      subtable.encodingId === WINDOWS_ENCODING.full &&
      subtable.format === CMAP_FORMAT.segmentedCoverage,
    symbol: false,
  },
  {
    matches: (subtable: CmapSubtable) =>
      subtable.platformId === PLATFORM.windows &&
      subtable.encodingId === WINDOWS_ENCODING.bmp &&
      subtable.format === CMAP_FORMAT.segmentMapping,
    symbol: false,
  },
  {
    matches: (subtable: CmapSubtable) =>
      subtable.platformId === PLATFORM.unicode &&
      (subtable.format === CMAP_FORMAT.segmentedCoverage ||
        subtable.format === CMAP_FORMAT.segmentMapping),
    symbol: false,
  },
  {
    matches: (subtable: CmapSubtable) =>
      subtable.platformId === PLATFORM.windows &&
      subtable.encodingId === WINDOWS_ENCODING.symbol &&
      subtable.format === CMAP_FORMAT.segmentMapping,
    symbol: true,
  },
  {
    // Last resort: any Macintosh or Unicode subtable in a format we read, so
    // an old font with only a (1,0) format 6 table still measures.
    matches: (subtable: CmapSubtable) =>
      subtable.platformId === PLATFORM.macintosh || subtable.platformId === PLATFORM.unicode,
    symbol: false,
  },
] as const;

const listCmapSubtables = (view: DataView, cmap: SfntTableRecord): readonly CmapSubtable[] => {
  const end = cmap.offset + cmap.length;
  if (!fitsInView(view, cmap.offset, CMAP.recordsStart)) {
    return [];
  }

  const numTables = view.getUint16(cmap.offset + CMAP.numTables);
  const subtables: CmapSubtable[] = [];
  for (let index = 0; index < numTables; index++) {
    const record = cmap.offset + CMAP.recordsStart + index * CMAP.recordSize;
    if (record + CMAP.recordSize > end || !fitsInView(view, record, CMAP.recordSize)) {
      break;
    }

    const start = cmap.offset + view.getUint32(record + CMAP_RECORD.offset);
    if (!fitsInView(view, start, 2) || start + 2 > end) {
      continue;
    }

    subtables.push({
      platformId: view.getUint16(record + CMAP_RECORD.platformId),
      encodingId: view.getUint16(record + CMAP_RECORD.encodingId),
      format: view.getUint16(start),
      start,
      end,
      view,
    });
  }
  return subtables;
};

/**
 * Picks the best subtable the preference order admits and builds its lookup
 * once. A `cmap` that exists but yields no usable subtable is an error: the
 * alternative is a font that silently maps every character to .notdef.
 */
const buildCmapLookup = (
  view: DataView,
  cmap: SfntTableRecord,
): Result<CmapLookup, SfntParseError> => {
  const subtables = listCmapSubtables(view, cmap);
  for (const preference of CMAP_PREFERENCES) {
    for (const subtable of subtables) {
      if (!preference.matches(subtable)) {
        continue;
      }

      const lookup = buildLookup(subtable);
      if (!lookup) {
        continue;
      }

      if (!preference.symbol) {
        return Result.ok(lookup);
      }
      return Result.ok((codePoint) => {
        const glyphId = lookup(codePoint);
        return glyphId === 0 ? lookup(SYMBOL_CODEPOINT_BASE | codePoint) : glyphId;
      });
    }
  }

  const formats = subtables.map((subtable) => subtable.format).join(", ");
  return fail(
    `cmap has no subtable this reader can use (found ${subtables.length} subtable(s), format(s): ${formats || "none"})`,
  );
};

/** Strips the delimiters a PDF name cannot carry and caps the spec's length. */
const MAX_POSTSCRIPT_NAME_LENGTH = 63;
const POSTSCRIPT_NAME_DELIMITERS = new Set("()<>[]{}/% ");

const sanitizePostScriptName = (raw: string): string => {
  let name = "";
  for (const character of raw) {
    const code = character.codePointAt(0) ?? 0;
    if (code > 0x20 && code < 0x7f && !POSTSCRIPT_NAME_DELIMITERS.has(character)) {
      name += character;
    }
  }
  return name.slice(0, MAX_POSTSCRIPT_NAME_LENGTH);
};

const decodeUtf16Be = (view: DataView, offset: number, length: number): string => {
  let text = "";
  for (let index = 0; index + 1 < length; index += 2) {
    text += String.fromCharCode(view.getUint16(offset + index));
  }
  return text;
};

const decodeMacRomanAscii = (view: DataView, offset: number, length: number): string => {
  let text = "";
  for (let index = 0; index < length; index++) {
    const code = view.getUint8(offset + index);
    // Only the ASCII subset is unambiguous without a MacRoman table; the
    // sanitizer would drop the high half anyway.
    text += code < 0x80 ? String.fromCharCode(code) : "";
  }
  return text;
};

/**
 * name id 6, preferring the Windows UTF-16BE record over the Macintosh one.
 * Returns `""` when the table is absent, unreadable, or carries no usable
 * PostScript name; the caller substitutes a synthetic name.
 */
const readPostScriptName = (view: DataView, name: SfntTableRecord): string => {
  const end = name.offset + name.length;
  if (!fitsInView(view, name.offset, NAME.size)) {
    return "";
  }

  const count = view.getUint16(name.offset + NAME.count);
  const stringsBase = name.offset + view.getUint16(name.offset + NAME.stringOffset);
  let windows = "";
  let macintosh = "";
  for (let index = 0; index < count; index++) {
    const record = name.offset + NAME.recordsStart + index * NAME.recordSize;
    if (record + NAME.recordSize > end || !fitsInView(view, record, NAME.recordSize)) {
      break;
    }
    if (view.getUint16(record + NAME_RECORD.nameId) !== POSTSCRIPT_NAME_ID) {
      continue;
    }

    const platformId = view.getUint16(record + NAME_RECORD.platformId);
    const length = view.getUint16(record + NAME_RECORD.length);
    const offset = stringsBase + view.getUint16(record + NAME_RECORD.offset);
    if (!fitsInView(view, offset, length) || offset + length > end) {
      continue;
    }

    if (platformId === PLATFORM.windows && windows === "") {
      windows = decodeUtf16Be(view, offset, length);
    } else if (platformId === PLATFORM.macintosh && macintosh === "") {
      macintosh = decodeMacRomanAscii(view, offset, length);
    }
  }

  return sanitizePostScriptName(windows) || sanitizePostScriptName(macintosh);
};

/**
 * A stable stand-in when the font carries no PostScript name. Derived from
 * `head.checkSumAdjustment`, which is a function of the whole file, so the
 * same bytes always produce the same name and two different faces do not
 * collide in a PDF resource dictionary.
 */
const syntheticPostScriptName = (checkSumAdjustment: number): string =>
  `SfntFont-${checkSumAdjustment.toString(16).toUpperCase().padStart(8, "0")}`;

type GlyphBoundsReaderOptions = {
  readonly directory: SfntDirectory;
  readonly numGlyphs: number;
  readonly longLoca: boolean;
};

/**
 * Builds the `glyphBoundsFor` accessor.
 *
 * A composite glyph carries its own bounds in the same header fields as a
 * simple one, so no component recursion is needed. A CFF font returns null
 * for every glyph: charstring interpretation is out of scope, and the caller
 * has to fall back on the hhea metrics.
 */
const buildGlyphBoundsReader = ({
  directory,
  numGlyphs,
  longLoca,
}: GlyphBoundsReaderOptions): ((glyphId: number) => SfntGlyphBounds | null) => {
  const loca = directory.tables.get(SFNT_TABLE.loca);
  const glyf = directory.tables.get(SFNT_TABLE.glyf);
  const offsets =
    loca && glyf
      ? readLocaOffsets({ view: directory.view, loca, numGlyphs, longFormat: longLoca })
      : undefined;
  if (!glyf || !offsets) {
    return () => null;
  }

  const { view } = directory;
  return (glyphId) => {
    if (!Number.isInteger(glyphId) || glyphId < 0 || glyphId >= numGlyphs) {
      return null;
    }

    const start = offsets[glyphId];
    const end = offsets[glyphId + 1];
    // Equal offsets are how a font spells "this glyph has no outline".
    if (start === undefined || end === undefined || end - start < GLYPH_HEADER.size) {
      return null;
    }

    const at = glyf.offset + start;
    if (!fitsInView(view, at, GLYPH_HEADER.size) || start + GLYPH_HEADER.size > glyf.length) {
      return null;
    }

    return {
      xMin: view.getInt16(at + GLYPH_HEADER.xMin),
      yMin: view.getInt16(at + GLYPH_HEADER.yMin),
      xMax: view.getInt16(at + GLYPH_HEADER.xMax),
      yMax: view.getInt16(at + GLYPH_HEADER.yMax),
    };
  };
};

/** Requires a table to exist and to be at least `size` bytes long. */
const requireTable = (
  directory: SfntDirectory,
  tag: string,
  size: number,
): Result<SfntTableRecord, SfntParseError> => {
  const record = directory.tables.get(tag);
  if (!record) {
    return fail(`font has no '${tag}' table`);
  }
  if (record.length < size || !fitsInView(directory.view, record.offset, size)) {
    return fail(`'${tag}' table is ${record.length} bytes, needs at least ${size}`);
  }
  return Result.ok(record);
};

const rejectContainer = (version: number): Result<void, SfntParseError> => {
  switch (version) {
    case SFNT_VERSION.collection:
      return fail("font collections ('ttcf') are not supported; extract a single face first");
    case SFNT_VERSION.woff:
      return fail("this is a WOFF container; decode it with decodeWoff before parsing");
    case SFNT_VERSION.woff2:
      return fail("this is a WOFF2 container; it must be decoded before parsing");
    case SFNT_VERSION.trueType:
    case SFNT_VERSION.trueMac:
    case SFNT_VERSION.openTypeCff:
      return Result.ok();
    default:
      return fail(`unrecognized sfnt version 0x${(version >>> 0).toString(16).padStart(8, "0")}`);
  }
};

export const parseSfnt = (bytes: Uint8Array): Result<SfntFont, SfntParseError> => {
  const version = readSfntVersion(bytes);
  if (version === undefined) {
    return fail(`buffer is ${bytes.byteLength} bytes, too short to be a font`);
  }

  const container = rejectContainer(version);
  if (container.isErr()) {
    return Result.err(container.error);
  }

  const directoryResult = readSfntDirectory(bytes).mapError(
    (error) => new SfntParseError({ message: error.message }),
  );
  if (directoryResult.isErr()) {
    return Result.err(directoryResult.error);
  }
  const directory = directoryResult.value;
  const { view } = directory;

  const headResult = requireTable(directory, SFNT_TABLE.head, HEAD.size);
  if (headResult.isErr()) {
    return Result.err(headResult.error);
  }
  const head = headResult.value;
  if (view.getUint32(head.offset + HEAD.magicNumber) !== HEAD_MAGIC) {
    return fail("'head' table has the wrong magic number");
  }

  const unitsPerEm = view.getUint16(head.offset + HEAD.unitsPerEm);
  if (unitsPerEm < MIN_UNITS_PER_EM || unitsPerEm > MAX_UNITS_PER_EM) {
    return fail(`head.unitsPerEm is ${unitsPerEm}, outside the permitted 16..16384`);
  }

  const hheaResult = requireTable(directory, SFNT_TABLE.hhea, HHEA.size);
  if (hheaResult.isErr()) {
    return Result.err(hheaResult.error);
  }
  const hhea = hheaResult.value;

  const maxpResult = requireTable(directory, SFNT_TABLE.maxp, MAXP.size);
  if (maxpResult.isErr()) {
    return Result.err(maxpResult.error);
  }
  const numGlyphs = view.getUint16(maxpResult.value.offset + MAXP.numGlyphs);
  if (numGlyphs === 0) {
    return fail("maxp.numGlyphs is 0");
  }

  const numberOfHMetrics = view.getUint16(hhea.offset + HHEA.numberOfHMetrics);
  if (numberOfHMetrics === 0 || numberOfHMetrics > numGlyphs) {
    return fail(
      `hhea.numberOfHMetrics is ${numberOfHMetrics}, which is not in 1..${numGlyphs} (maxp.numGlyphs)`,
    );
  }

  const hmtxResult = requireTable(directory, SFNT_TABLE.hmtx, numberOfHMetrics * HMTX_METRIC_SIZE);
  if (hmtxResult.isErr()) {
    return Result.err(hmtxResult.error);
  }
  const hmtx = hmtxResult.value;

  const cmap = directory.tables.get(SFNT_TABLE.cmap);
  // A subset font legitimately has no cmap: a PDF CIDFontType2 addresses
  // glyphs directly, so the tables that map characters are dropped.
  const unmappedLookup: CmapLookup = () => 0;
  const lookupResult = cmap ? buildCmapLookup(view, cmap) : Result.ok(unmappedLookup);
  if (lookupResult.isErr()) {
    return Result.err(lookupResult.error);
  }
  const lookup = lookupResult.value;
  const glyphIdFor = (codePoint: number) => {
    if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > MAX_UNICODE_CODE_POINT) {
      return 0;
    }
    const glyphId = lookup(codePoint);
    return glyphId < numGlyphs ? glyphId : 0;
  };

  const advanceWidthFor = (glyphId: number) => {
    if (!Number.isInteger(glyphId) || glyphId < 0) {
      return 0;
    }
    // Past the last longHorMetric every glyph shares that entry's advance.
    const index = Math.min(glyphId, numberOfHMetrics - 1);
    return view.getUint16(hmtx.offset + index * HMTX_METRIC_SIZE);
  };

  const longLoca = view.getInt16(head.offset + HEAD.indexToLocFormat) === 1;
  const glyphBoundsFor = buildGlyphBoundsReader({ directory, numGlyphs, longLoca });

  const os2 = directory.tables.get(SFNT_TABLE.os2);
  const os2HasHeights =
    os2 !== undefined &&
    os2.length >= OS2.sizeWithHeights &&
    fitsInView(view, os2.offset, OS2.sizeWithHeights) &&
    view.getUint16(os2.offset + OS2.version) >= OS2.minimumVersionWithHeights;

  const capHeight = os2HasHeights
    ? view.getInt16(os2.offset + OS2.sCapHeight)
    : (glyphBoundsFor(glyphIdFor(CAP_HEIGHT_SAMPLE))?.yMax ??
      Math.round(unitsPerEm * CAP_HEIGHT_EM_RATIO));
  const xHeight = os2HasHeights
    ? view.getInt16(os2.offset + OS2.sxHeight)
    : (glyphBoundsFor(glyphIdFor(X_HEIGHT_SAMPLE))?.yMax ??
      Math.round(unitsPerEm * X_HEIGHT_EM_RATIO));

  const fsType =
    os2 !== undefined &&
    os2.length >= OS2.fsType + 2 &&
    fitsInView(view, os2.offset, OS2.fsType + 2)
      ? view.getUint16(os2.offset + OS2.fsType)
      : 0;

  const post = directory.tables.get(SFNT_TABLE.post);
  const italicAngle =
    post !== undefined && post.length >= POST.size && fitsInView(view, post.offset, POST.size)
      ? view.getInt32(post.offset + POST.italicAngle) / FIXED_POINT_DIVISOR
      : 0;

  const name = directory.tables.get(SFNT_TABLE.name);
  const postScriptName =
    (name ? readPostScriptName(view, name) : "") ||
    syntheticPostScriptName(view.getUint32(head.offset + HEAD.checkSumAdjustment));

  return Result.ok({
    unitsPerEm,
    ascender: view.getInt16(hhea.offset + HHEA.ascender),
    descender: view.getInt16(hhea.offset + HHEA.descender),
    lineGap: view.getInt16(hhea.offset + HHEA.lineGap),
    capHeight,
    xHeight,
    italicAngle,
    bbox: [
      view.getInt16(head.offset + HEAD.xMin),
      view.getInt16(head.offset + HEAD.yMin),
      view.getInt16(head.offset + HEAD.xMax),
      view.getInt16(head.offset + HEAD.yMax),
    ],
    postScriptName,
    isCff: directory.tables.has(SFNT_TABLE.cff),
    numGlyphs,
    fsType,
    bytes,
    glyphIdFor,
    advanceWidthFor,
    glyphBoundsFor,
  });
};
