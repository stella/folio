/**
 * Shared sfnt container primitives: table tags, the container magics, the
 * table directory, and the `loca` offset array.
 *
 * `parse.ts`, `woff.ts` and `subset.ts` all address tables through the same
 * directory, and both the parser (deriving cap height from glyph bounds) and
 * the subsetter (copying glyph data) read `loca`. Keeping one reader here
 * means the two cannot disagree about what a truncated or non-monotonic
 * table means.
 */

import { Result, TaggedError } from "better-result";

/** Four-character table tags this package addresses by name. */
export const SFNT_TABLE = {
  cff: "CFF ",
  cmap: "cmap",
  cvt: "cvt ",
  fpgm: "fpgm",
  glyf: "glyf",
  head: "head",
  hhea: "hhea",
  hmtx: "hmtx",
  loca: "loca",
  maxp: "maxp",
  name: "name",
  os2: "OS/2",
  post: "post",
  prep: "prep",
} as const;

/**
 * sfnt `version` words (the same field WOFF calls `flavor`), plus the
 * container magics that are *not* a bare sfnt and must be reported as such.
 */
export const SFNT_VERSION = {
  /** `0x00010000`: TrueType outlines. */
  trueType: 0x00010000,
  /** `true`: the legacy Apple spelling of the same thing. */
  trueMac: 0x74727565,
  /** `OTTO`: OpenType with CFF outlines. */
  openTypeCff: 0x4f54544f,
  /** `ttcf`: a font collection, not a single font. */
  collection: 0x74746366,
  /** `wOFF`: WOFF 1.0 container. */
  woff: 0x774f4646,
  /** `wOF2`: WOFF 2.0 container. */
  woff2: 0x774f4632,
} as const;

/** sfnt header: version, numTables, searchRange, entrySelector, rangeShift. */
export const SFNT_HEADER_SIZE = 12;
/** Directory entry: tag, checksum, offset, length. */
export const SFNT_DIRECTORY_ENTRY_SIZE = 16;
/** Every table starts on a four-byte boundary and is zero-padded to one. */
export const SFNT_TABLE_ALIGNMENT = 4;

const DIRECTORY_ENTRY = {
  tag: 0,
  checkSum: 4,
  offset: 8,
  length: 12,
} as const;

const TAG_LENGTH = 4;

/** A structural problem in the container: bad header, or a table out of bounds. */
export class SfntDirectoryError extends TaggedError("SfntDirectoryError")<{ message: string }> {}

/** One entry of the table directory, already bounds-checked against the buffer. */
export type SfntTableRecord = {
  readonly tag: string;
  readonly offset: number;
  readonly length: number;
};

export type SfntDirectory = {
  readonly version: number;
  readonly view: DataView;
  readonly tables: ReadonlyMap<string, SfntTableRecord>;
};

/** Whether `[offset, offset + length)` lies wholly inside `view`. */
export const fitsInView = (view: DataView, offset: number, length: number): boolean =>
  Number.isInteger(offset) &&
  Number.isInteger(length) &&
  offset >= 0 &&
  length >= 0 &&
  offset + length <= view.byteLength;

/** The four-byte tag at `offset`, or `""` when it does not fit. */
export const readTag = (view: DataView, offset: number): string => {
  if (!fitsInView(view, offset, TAG_LENGTH)) {
    return "";
  }

  let tag = "";
  for (let index = 0; index < TAG_LENGTH; index++) {
    tag += String.fromCharCode(view.getUint8(offset + index));
  }
  return tag;
};

/** The leading version word, or `undefined` when the buffer is shorter than it. */
export const readSfntVersion = (bytes: Uint8Array): number | undefined => {
  if (bytes.byteLength < TAG_LENGTH) {
    return undefined;
  }
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0);
};

/**
 * Reads the table directory, validating every entry against the buffer.
 *
 * The caller is expected to have rejected non-sfnt magics first: a collection
 * or WOFF header parsed as a directory yields nonsense offsets and a much
 * worse error message than the container check gives.
 */
export const readSfntDirectory = (bytes: Uint8Array): Result<SfntDirectory, SfntDirectoryError> => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (!fitsInView(view, 0, SFNT_HEADER_SIZE)) {
    return Result.err(
      new SfntDirectoryError({
        message: `sfnt header needs ${SFNT_HEADER_SIZE} bytes, buffer has ${view.byteLength}`,
      }),
    );
  }

  const version = view.getUint32(0);
  const numTables = view.getUint16(4);
  const directorySize = numTables * SFNT_DIRECTORY_ENTRY_SIZE;
  if (!fitsInView(view, SFNT_HEADER_SIZE, directorySize)) {
    return Result.err(
      new SfntDirectoryError({
        message: `table directory declares ${numTables} tables, which does not fit in ${view.byteLength} bytes`,
      }),
    );
  }

  const tables = new Map<string, SfntTableRecord>();
  for (let index = 0; index < numTables; index++) {
    const entry = SFNT_HEADER_SIZE + index * SFNT_DIRECTORY_ENTRY_SIZE;
    const tag = readTag(view, entry + DIRECTORY_ENTRY.tag);
    const offset = view.getUint32(entry + DIRECTORY_ENTRY.offset);
    const length = view.getUint32(entry + DIRECTORY_ENTRY.length);
    if (!fitsInView(view, offset, length)) {
      return Result.err(
        new SfntDirectoryError({
          message: `table '${tag}' spans [${offset}, ${offset + length}) outside the ${view.byteLength}-byte font`,
        }),
      );
    }

    // A duplicate tag is malformed; the first entry wins so the result does
    // not depend on directory order.
    if (!tables.has(tag)) {
      tables.set(tag, { tag, offset, length });
    }
  }

  return Result.ok({ version, view, tables });
};

type LocaOffsetsOptions = {
  readonly view: DataView;
  readonly loca: SfntTableRecord;
  readonly numGlyphs: number;
  /** `head.indexToLocFormat === 1`: offsets are uint32, not uint16 halves. */
  readonly longFormat: boolean;
};

/**
 * The `numGlyphs + 1` glyph offsets from `loca`, or `undefined` when the table
 * is too short or its offsets run backwards (either makes glyph extraction
 * meaningless).
 */
export const readLocaOffsets = ({
  view,
  loca,
  numGlyphs,
  longFormat,
}: LocaOffsetsOptions): readonly number[] | undefined => {
  const entrySize = longFormat ? 4 : 2;
  const count = numGlyphs + 1;
  if (loca.length < count * entrySize || !fitsInView(view, loca.offset, count * entrySize)) {
    return undefined;
  }

  const offsets: number[] = [];
  let previous = -1;
  for (let index = 0; index < count; index++) {
    const at = loca.offset + index * entrySize;
    const offset = longFormat ? view.getUint32(at) : view.getUint16(at) * 2;
    if (offset < previous) {
      return undefined;
    }
    previous = offset;
    offsets.push(offset);
  }
  return offsets;
};
