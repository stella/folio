/**
 * WOFF 1.0 container decoding.
 *
 * WOFF 1.0 is a repackaging, not a transform: each table is stored either raw
 * or zlib-deflated, and the sfnt is recovered by inflating the tables and
 * writing a fresh directory around them. WOFF 2.0 is a different problem (its
 * `glyf`/`loca` transform has to be reversed and Brotli decoded), so it is
 * rejected here rather than half-supported.
 */

import { inflateSync } from "node:zlib";

import { Result, TaggedError } from "better-result";

import {
  readSfntVersion,
  readTag,
  SFNT_DIRECTORY_ENTRY_SIZE,
  SFNT_HEADER_SIZE,
  SFNT_TABLE_ALIGNMENT,
  SFNT_VERSION,
} from "./tables";

export class WoffDecodeError extends TaggedError("WoffDecodeError")<{ message: string }> {}

const fail = (message: string) => Result.err(new WoffDecodeError({ message }));

/** Byte offsets inside the WOFF header. */
const WOFF_HEADER = {
  signature: 0,
  flavor: 4,
  length: 8,
  numTables: 12,
  size: 44,
} as const;

/** Byte offsets inside a WOFF table directory entry. */
const WOFF_ENTRY = {
  tag: 0,
  offset: 4,
  compLength: 8,
  origLength: 12,
  origChecksum: 16,
  size: 20,
} as const;

/**
 * Ceiling on a single decompressed table. Inflate is a decompression bomb
 * vector, and a font table larger than this is not a font we can use anyway.
 */
const MAX_TABLE_LENGTH = 1 << 27;

/**
 * Ceiling on the whole decoded font. Nothing stops two directory entries from
 * naming the same compressed span, so a per-table limit bounds one inflate but
 * not what 65535 of them add up to. The largest faces anyone ships as web
 * fonts, full CJK OpenType, decompress to some tens of megabytes; 256 MB
 * clears that with room to spare and still refuses the terabytes an adversarial
 * directory could otherwise claim.
 */
const MAX_DECODED_FONT_LENGTH = 1 << 28;

/** One WOFF directory entry, read before any table is inflated. */
type WoffEntry = {
  readonly tag: string;
  readonly offset: number;
  readonly compLength: number;
  readonly origLength: number;
  readonly checksum: number;
};

type DecodedTable = {
  readonly tag: string;
  readonly checksum: number;
  readonly data: Uint8Array;
};

const alignUp = (value: number): number =>
  Math.ceil(value / SFNT_TABLE_ALIGNMENT) * SFNT_TABLE_ALIGNMENT;

/**
 * `searchRange`, `entrySelector` and `rangeShift` from the sfnt header: a
 * binary-search hint derived from the table count. Wrong values do not break
 * every consumer, but they do break the strict ones.
 */
const searchParametersFor = (numTables: number) => {
  let entrySelector = 0;
  while (1 << (entrySelector + 1) <= numTables) {
    entrySelector += 1;
  }
  const searchRange = (1 << entrySelector) * SFNT_DIRECTORY_ENTRY_SIZE;
  return {
    searchRange,
    entrySelector,
    rangeShift: numTables * SFNT_DIRECTORY_ENTRY_SIZE - searchRange,
  };
};

/** Assembles decoded tables into a standard sfnt, directory sorted by tag. */
const buildSfnt = (flavor: number, tables: readonly DecodedTable[]): Uint8Array => {
  const sorted = [...tables].sort((left, right) => (left.tag < right.tag ? -1 : 1));
  const directorySize = sorted.length * SFNT_DIRECTORY_ENTRY_SIZE;
  const bodySize = sorted.reduce((total, table) => total + alignUp(table.data.byteLength), 0);
  const output = new Uint8Array(SFNT_HEADER_SIZE + directorySize + bodySize);
  const view = new DataView(output.buffer);

  const { searchRange, entrySelector, rangeShift } = searchParametersFor(sorted.length);
  view.setUint32(0, flavor);
  view.setUint16(4, sorted.length);
  view.setUint16(6, searchRange);
  view.setUint16(8, entrySelector);
  view.setUint16(10, rangeShift);

  let entry = SFNT_HEADER_SIZE;
  let dataOffset = SFNT_HEADER_SIZE + directorySize;
  for (const table of sorted) {
    for (let index = 0; index < table.tag.length; index++) {
      view.setUint8(entry + index, table.tag.charCodeAt(index));
    }
    view.setUint32(entry + 4, table.checksum);
    view.setUint32(entry + 8, dataOffset);
    view.setUint32(entry + 12, table.data.byteLength);
    output.set(table.data, dataOffset);
    // The gap up to the next four-byte boundary is already zero.
    dataOffset += alignUp(table.data.byteLength);
    entry += SFNT_DIRECTORY_ENTRY_SIZE;
  }

  return output;
};

export const decodeWoff = (bytes: Uint8Array): Result<Uint8Array, WoffDecodeError> => {
  if (bytes.byteLength < WOFF_HEADER.size) {
    return fail(`WOFF header needs ${WOFF_HEADER.size} bytes, buffer has ${bytes.byteLength}`);
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const signature = view.getUint32(WOFF_HEADER.signature);
  if (signature !== SFNT_VERSION.woff) {
    return fail(
      signature === SFNT_VERSION.woff2
        ? "this is a WOFF2 container; decodeWoff only handles WOFF 1.0"
        : `not a WOFF container: signature 0x${signature.toString(16).padStart(8, "0")}`,
    );
  }

  const flavor = view.getUint32(WOFF_HEADER.flavor);
  const numTables = view.getUint16(WOFF_HEADER.numTables);
  if (numTables === 0) {
    return fail("WOFF declares 0 tables");
  }

  const directorySize = numTables * WOFF_ENTRY.size;
  if (WOFF_HEADER.size + directorySize > bytes.byteLength) {
    return fail(
      `WOFF declares ${numTables} tables, whose directory does not fit in ${bytes.byteLength} bytes`,
    );
  }

  const entries: WoffEntry[] = [];
  let decodedLength = 0;
  for (let index = 0; index < numTables; index++) {
    const entry = WOFF_HEADER.size + index * WOFF_ENTRY.size;
    const tag = readTag(view, entry + WOFF_ENTRY.tag);
    const offset = view.getUint32(entry + WOFF_ENTRY.offset);
    const compLength = view.getUint32(entry + WOFF_ENTRY.compLength);
    const origLength = view.getUint32(entry + WOFF_ENTRY.origLength);
    const checksum = view.getUint32(entry + WOFF_ENTRY.origChecksum);

    if (offset + compLength > bytes.byteLength) {
      return fail(
        `WOFF table '${tag}' spans [${offset}, ${offset + compLength}) outside the ${bytes.byteLength}-byte container`,
      );
    }
    if (origLength > MAX_TABLE_LENGTH) {
      return fail(`WOFF table '${tag}' claims ${origLength} bytes, above the decode ceiling`);
    }
    decodedLength += origLength;
    if (decodedLength > MAX_DECODED_FONT_LENGTH) {
      return fail(
        `WOFF directory claims ${decodedLength} decoded bytes in total, above the ${MAX_DECODED_FONT_LENGTH}-byte container ceiling`,
      );
    }

    entries.push({ tag, offset, compLength, origLength, checksum });
  }

  const tables: DecodedTable[] = [];
  for (const { tag, offset, compLength, origLength, checksum } of entries) {
    const stored = bytes.subarray(offset, offset + compLength);
    if (compLength === origLength) {
      tables.push({ tag, checksum, data: stored });
      continue;
    }

    // The entry declares the decompressed size, so the inflate is bounded by
    // that rather than by the ceiling: a table cannot spend the whole per-table
    // budget only to be rejected afterwards for not matching its own header.
    const inflated = Result.try({
      try: () => inflateSync(stored, { maxOutputLength: origLength }),
      catch: (cause) =>
        new WoffDecodeError({
          message: `WOFF table '${tag}' failed to inflate: ${cause instanceof Error ? cause.message : String(cause)}`,
        }),
    });
    if (inflated.isErr()) {
      return Result.err(inflated.error);
    }
    if (inflated.value.byteLength !== origLength) {
      return fail(
        `WOFF table '${tag}' inflated to ${inflated.value.byteLength} bytes, not the declared ${origLength}`,
      );
    }

    tables.push({
      tag,
      checksum,
      data: new Uint8Array(
        inflated.value.buffer,
        inflated.value.byteOffset,
        inflated.value.byteLength,
      ),
    });
  }

  return Result.ok(buildSfnt(flavor, tables));
};

/**
 * The entry point a caller with bytes of unknown provenance wants: plain sfnt
 * passes through untouched, WOFF is decoded, WOFF2 is refused by name.
 */
export const toSfntBytes = (bytes: Uint8Array): Result<Uint8Array, WoffDecodeError> => {
  const version = readSfntVersion(bytes);
  if (version === undefined) {
    return fail(`buffer is ${bytes.byteLength} bytes, too short to identify`);
  }

  switch (version) {
    case SFNT_VERSION.woff:
      return decodeWoff(bytes);
    case SFNT_VERSION.woff2:
      return fail(
        "WOFF2 is not supported: its glyf/loca transform has to be reversed, which is a separate decoder",
      );
    case SFNT_VERSION.trueType:
    case SFNT_VERSION.trueMac:
    case SFNT_VERSION.openTypeCff:
    case SFNT_VERSION.collection:
      // A collection is not decodable here either, but it is already sfnt
      // bytes: let the parser be the one that explains it.
      return Result.ok(bytes);
    default:
      return fail(`unrecognized font container 0x${version.toString(16).padStart(8, "0")}`);
  }
};
