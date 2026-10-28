/**
 * Central-directory reading with every size an archive states about itself
 * checked before anything is inflated.
 *
 * An archive is a set of claims written by whoever produced it: the entry
 * count, each entry's sizes and offsets, the names. The reader accepts the
 * archive only when the end record, the central directory and every local
 * header agree, the entries occupy disjoint byte ranges inside the archive,
 * and the declared sizes fit the limits. Inflation later stops at the
 * declared size, so a size that passed here bounds every read.
 */
import { Result } from "better-result";

import { ZipError, type ZipErrorReason } from "./errors";
import {
  CENTRAL_DIRECTORY_HEADER_SIZE,
  CENTRAL_DIRECTORY_SIGNATURE,
  type DosDateTime,
  END_OF_CENTRAL_DIRECTORY_SIGNATURE,
  END_OF_CENTRAL_DIRECTORY_SIZE,
  FLAG_DATA_DESCRIPTOR,
  FLAG_ENCRYPTED,
  FLAG_STRONG_ENCRYPTION,
  isSafeZipPath,
  LOCAL_FILE_HEADER_SIGNATURE,
  LOCAL_FILE_HEADER_SIZE,
  METHOD_DEFLATE,
  METHOD_STORE,
  UINT16_SENTINEL,
  UINT32_SENTINEL,
  ZIP64_END_LOCATOR_SIGNATURE,
  ZIP64_END_LOCATOR_SIZE,
  ZIP64_END_OF_CENTRAL_DIRECTORY_SIGNATURE,
  ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE,
  ZIP64_EXTRA_FIELD_TAG,
} from "./format";

/** Bounds a read enforces. Every byte limit counts inflated bytes. */
export type ZipReadLimits = {
  /** Records the central directory may hold, directories included. */
  readonly maxEntries: number;
  /** Inflated bytes one entry may declare. */
  readonly maxEntryBytes: number;
  /** Inflated bytes all entries together may declare. */
  readonly maxTotalBytes: number;
  /**
   * Inflated bytes per compressed byte, for each entry and for the archive
   * as a whole. Sizes under `compressionRatioGraceBytes` are exempt: a small
   * part can compress extremely well without costing anything to hold.
   */
  readonly maxCompressionRatio: number;
  readonly compressionRatioGraceBytes: number;
};

const MEBIBYTE = 1024 * 1024;

/**
 * Defaults sized for office packages. A DEFLATE stream tops out near 1032:1;
 * markup in real documents stays under 75:1.
 */
export const ZIP_DEFAULT_READ_LIMITS: ZipReadLimits = {
  maxEntries: 10_000,
  maxEntryBytes: 256 * MEBIBYTE,
  maxTotalBytes: 512 * MEBIBYTE,
  maxCompressionRatio: 200,
  compressionRatioGraceBytes: 4 * MEBIBYTE,
};

/** One file entry as the central directory and its local header describe it. */
export type SourceEntry = {
  readonly path: string;
  readonly name: Uint8Array;
  readonly flags: number;
  readonly method: typeof METHOD_STORE | typeof METHOD_DEFLATE;
  readonly dateTime: DosDateTime;
  readonly crc32: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  /** Offset of the entry's compressed bytes in the archive. */
  readonly dataOffset: number;
};

type ReadOptions = {
  readonly limits: ZipReadLimits;
  /** Entries the per-entry ratio cap skips; the archive ratio still counts them. */
  readonly isCompressionRatioExempt: (path: string) => boolean;
};

const fail = (
  reason: ZipErrorReason,
  message: string,
  path?: string,
): Result<never, ZipError> =>
  Result.err(
    path === undefined
      ? new ZipError({ message, reason })
      : new ZipError({ message, reason, path }),
  );

const malformed = (message: string): Result<never, ZipError> =>
  fail("malformed", message);

type Directory = {
  readonly count: number;
  readonly offset: number;
  readonly end: number;
};

const toSafeNumber = (value: bigint): number | null =>
  value <= BigInt(Number.MAX_SAFE_INTEGER) ? Number(value) : null;

const findEndRecord = (bytes: Uint8Array, view: DataView): number => {
  const earliest = Math.max(
    0,
    bytes.byteLength - END_OF_CENTRAL_DIRECTORY_SIZE - UINT16_SENTINEL,
  );
  for (
    let offset = bytes.byteLength - END_OF_CENTRAL_DIRECTORY_SIZE;
    offset >= earliest;
    offset--
  ) {
    if (
      view.getUint32(offset, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE &&
      offset +
        END_OF_CENTRAL_DIRECTORY_SIZE +
        view.getUint16(offset + 20, true) ===
        bytes.byteLength
    ) {
      return offset;
    }
  }
  return -1;
};

const readZip64Directory = (
  view: DataView,
  end: number,
): Result<Directory, ZipError> => {
  const locator = end - ZIP64_END_LOCATOR_SIZE;
  if (
    locator < ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE ||
    view.getUint32(locator, true) !== ZIP64_END_LOCATOR_SIGNATURE
  ) {
    return malformed("Zip64 end locator is missing");
  }
  if (
    view.getUint32(locator + 4, true) !== 0 ||
    view.getUint32(locator + 16, true) !== 1
  ) {
    return fail("unsupported", "Multi-disk archives are not supported");
  }
  const record = toSafeNumber(view.getBigUint64(locator + 8, true));
  // Only the fixed-size record is accepted: no extensible data sector, so the
  // record ends exactly at its locator.
  if (
    record === null ||
    record + ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE !== locator ||
    view.getUint32(record, true) !== ZIP64_END_OF_CENTRAL_DIRECTORY_SIGNATURE ||
    view.getBigUint64(record + 4, true) !==
      BigInt(ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE - 12)
  ) {
    return malformed("Zip64 end record is malformed");
  }
  if (
    view.getUint32(record + 16, true) !== 0 ||
    view.getUint32(record + 20, true) !== 0
  ) {
    return fail("unsupported", "Multi-disk archives are not supported");
  }
  const diskCount = view.getBigUint64(record + 24, true);
  const count = view.getBigUint64(record + 32, true);
  const size = toSafeNumber(view.getBigUint64(record + 40, true));
  const offset = toSafeNumber(view.getBigUint64(record + 48, true));
  const entryCount = toSafeNumber(count);
  if (
    diskCount !== count ||
    size === null ||
    offset === null ||
    entryCount === null ||
    offset + size !== record
  ) {
    return malformed("Zip64 end record disagrees with the central directory");
  }
  return Result.ok({ count: entryCount, offset, end: record });
};

const readDirectory = (
  bytes: Uint8Array,
  view: DataView,
): Result<Directory, ZipError> => {
  const end = findEndRecord(bytes, view);
  if (end === -1) return malformed("End of central directory is missing");
  if (
    view.getUint16(end + 4, true) !== 0 ||
    view.getUint16(end + 6, true) !== 0
  ) {
    return fail("unsupported", "Multi-disk archives are not supported");
  }
  const diskCount = view.getUint16(end + 8, true);
  const count = view.getUint16(end + 10, true);
  const size = view.getUint32(end + 12, true);
  const offset = view.getUint32(end + 16, true);
  const deferred =
    count === UINT16_SENTINEL ||
    diskCount === UINT16_SENTINEL ||
    size === UINT32_SENTINEL ||
    offset === UINT32_SENTINEL;
  const hasLocator =
    end >= ZIP64_END_LOCATOR_SIZE &&
    view.getUint32(end - ZIP64_END_LOCATOR_SIZE, true) ===
      ZIP64_END_LOCATOR_SIGNATURE;
  if (deferred || hasLocator) {
    const zip64 = readZip64Directory(view, end);
    if (Result.isError(zip64)) return zip64;
    const directory = zip64.value;
    // A classic field that is not the sentinel must agree with Zip64.
    if (
      (count !== UINT16_SENTINEL && count !== directory.count) ||
      (offset !== UINT32_SENTINEL && offset !== directory.offset) ||
      (size !== UINT32_SENTINEL && size !== directory.end - directory.offset)
    ) {
      return malformed("End records disagree");
    }
    return Result.ok(directory);
  }
  if (diskCount !== count || offset + size !== end) {
    return malformed("End record disagrees with the central directory");
  }
  return Result.ok({ count, offset, end });
};

type Zip64Fields = {
  uncompressedSize: number;
  compressedSize: number;
  localOffset: number;
};

const readZip64Extra = (
  view: DataView,
  start: number,
  end: number,
  fields: Zip64Fields,
): Result<Zip64Fields, ZipError> => {
  let cursor = start;
  let seen = false;
  const resolved = { ...fields };
  while (cursor < end) {
    if (cursor + 4 > end) return malformed("Extra field is truncated");
    const tag = view.getUint16(cursor, true);
    const length = view.getUint16(cursor + 2, true);
    const payload = cursor + 4;
    if (payload + length > end) return malformed("Extra field is truncated");
    if (tag === ZIP64_EXTRA_FIELD_TAG) {
      if (seen) return malformed("Duplicate Zip64 extra field");
      seen = true;
      let field = payload;
      for (const key of [
        "uncompressedSize",
        "compressedSize",
        "localOffset",
      ] as const) {
        if (fields[key] !== UINT32_SENTINEL) continue;
        if (field + 8 > payload + length) {
          return malformed("Zip64 extra field is truncated");
        }
        const value = toSafeNumber(view.getBigUint64(field, true));
        if (value === null) return malformed("Zip64 size is out of range");
        resolved[key] = value;
        field += 8;
      }
    }
    cursor = payload + length;
  }
  if (
    !seen &&
    (fields.uncompressedSize === UINT32_SENTINEL ||
      fields.compressedSize === UINT32_SENTINEL ||
      fields.localOffset === UINT32_SENTINEL)
  ) {
    return malformed("Zip64 extra field is missing");
  }
  return Result.ok(resolved);
};

const utf8Decoder = new TextDecoder("utf-8", { fatal: true });

// A name that is not UTF-8 has no unambiguous reading, so it is refused
// rather than guessed at.
const decodeName = (name: Uint8Array): string | null => {
  const decoded = Result.try(() => utf8Decoder.decode(name));
  return Result.isOk(decoded) ? decoded.value : null;
};

const LOCAL_FLAGS_THAT_MUST_AGREE = FLAG_ENCRYPTED | FLAG_DATA_DESCRIPTOR;

const bytesEqual = (left: Uint8Array, right: Uint8Array): boolean =>
  left.byteLength === right.byteLength &&
  left.every((byte, index) => byte === right[index]);

const exceedsRatio = (
  inflated: number,
  compressed: number,
  limits: ZipReadLimits,
): boolean =>
  inflated > limits.compressionRatioGraceBytes &&
  inflated > compressed * limits.maxCompressionRatio;

type Span = { readonly start: number; readonly end: number };

/**
 * Read and validate the archive's directory. Nothing is inflated; the
 * returned entries carry the sizes every later inflation is held to.
 */
export const readEntries = (
  bytes: Uint8Array,
  { limits, isCompressionRatioExempt }: ReadOptions,
): Result<SourceEntry[], ZipError> => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const directory = readDirectory(bytes, view);
  if (Result.isError(directory)) return directory;
  const { count, offset: directoryStart, end: directoryEnd } = directory.value;
  if (count > limits.maxEntries) {
    return fail("too-many-entries", "Archive declares too many entries");
  }

  const entries: SourceEntry[] = [];
  const names = new Set<string>();
  const spans: Span[] = [];
  let totalBytes = 0;
  let cursor = directoryStart;
  for (let index = 0; index < count; index++) {
    if (
      cursor + CENTRAL_DIRECTORY_HEADER_SIZE > directoryEnd ||
      view.getUint32(cursor, true) !== CENTRAL_DIRECTORY_SIGNATURE
    ) {
      return malformed("Central directory record is malformed");
    }
    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const nameStart = cursor + CENTRAL_DIRECTORY_HEADER_SIZE;
    const extraStart = nameStart + nameLength;
    const next = extraStart + extraLength + commentLength;
    if (next > directoryEnd || nameLength === 0) {
      return malformed("Central directory record is malformed");
    }
    if (view.getUint16(cursor + 34, true) !== 0) {
      return fail("unsupported", "Multi-disk archives are not supported");
    }
    const name = bytes.subarray(nameStart, extraStart);
    const path = decodeName(name);
    if (path === null || !isSafeZipPath(path)) {
      return fail(
        "invalid-entry-name",
        "Archive contains an unsafe entry name",
      );
    }
    if (names.has(path)) {
      return fail("duplicate-entry", "Archive names an entry twice", path);
    }
    names.add(path);
    if ((flags & (FLAG_ENCRYPTED | FLAG_STRONG_ENCRYPTION)) !== 0) {
      return fail("unsupported", "Encrypted entries are not supported", path);
    }
    if (method !== METHOD_STORE && method !== METHOD_DEFLATE) {
      return fail("unsupported", "Unsupported compression method", path);
    }
    const sizes = readZip64Extra(view, extraStart, extraStart + extraLength, {
      uncompressedSize: view.getUint32(cursor + 24, true),
      compressedSize: view.getUint32(cursor + 20, true),
      localOffset: view.getUint32(cursor + 42, true),
    });
    if (Result.isError(sizes)) return sizes;
    const { uncompressedSize, compressedSize, localOffset } = sizes.value;
    if (method === METHOD_STORE && compressedSize !== uncompressedSize) {
      return malformed("Stored entry sizes disagree");
    }

    // The local header must exist where the directory says, carry the same
    // name, method and flags, and leave the data inside the archive's file
    // section.
    if (
      localOffset + LOCAL_FILE_HEADER_SIZE > directoryStart ||
      view.getUint32(localOffset, true) !== LOCAL_FILE_HEADER_SIGNATURE
    ) {
      return malformed("Local header is missing");
    }
    const localFlags = view.getUint16(localOffset + 6, true);
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataOffset =
      localOffset + LOCAL_FILE_HEADER_SIZE + localNameLength + localExtraLength;
    if (
      view.getUint16(localOffset + 8, true) !== method ||
      (localFlags & LOCAL_FLAGS_THAT_MUST_AGREE) !==
        (flags & LOCAL_FLAGS_THAT_MUST_AGREE) ||
      dataOffset + compressedSize > directoryStart ||
      !bytesEqual(
        bytes.subarray(
          localOffset + LOCAL_FILE_HEADER_SIZE,
          localOffset + LOCAL_FILE_HEADER_SIZE + localNameLength,
        ),
        name,
      )
    ) {
      return malformed("Local header disagrees with the central directory");
    }
    if ((flags & FLAG_DATA_DESCRIPTOR) === 0) {
      const localCrc = view.getUint32(localOffset + 14, true);
      const localCompressed = view.getUint32(localOffset + 18, true);
      const localUncompressed = view.getUint32(localOffset + 22, true);
      if (
        localCrc !== view.getUint32(cursor + 16, true) ||
        (localCompressed !== UINT32_SENTINEL &&
          localCompressed !== compressedSize) ||
        (localUncompressed !== UINT32_SENTINEL &&
          localUncompressed !== uncompressedSize)
      ) {
        return malformed("Local header disagrees with the central directory");
      }
    }
    spans.push({ start: localOffset, end: dataOffset + compressedSize });

    const directoryEntry = path.endsWith("/");
    if (directoryEntry) {
      if (uncompressedSize !== 0) return malformed("Directory entry has data");
      cursor = next;
      continue;
    }
    if (uncompressedSize > limits.maxEntryBytes) {
      return fail("entry-too-large", "Archive entry is too large", path);
    }
    totalBytes += uncompressedSize;
    if (totalBytes > limits.maxTotalBytes) {
      return fail("total-too-large", "Archive expands beyond its limit");
    }
    if (
      method === METHOD_DEFLATE &&
      !isCompressionRatioExempt(path) &&
      exceedsRatio(uncompressedSize, compressedSize, limits)
    ) {
      return fail(
        "compression-ratio",
        "Archive entry exceeds the compression ratio limit",
        path,
      );
    }
    entries.push({
      path,
      name,
      flags,
      method,
      dateTime: {
        time: view.getUint16(cursor + 12, true),
        date: view.getUint16(cursor + 14, true),
      },
      crc32: view.getUint32(cursor + 16, true),
      compressedSize,
      uncompressedSize,
      dataOffset,
    });
    cursor = next;
  }
  if (cursor !== directoryEnd) {
    return malformed("Central directory size disagrees with its records");
  }
  if (exceedsRatio(totalBytes, bytes.byteLength, limits)) {
    return fail(
      "compression-ratio",
      "Archive exceeds the compression ratio limit",
    );
  }
  // Entries that share bytes let one compressed stream count many times.
  spans.sort((left, right) => left.start - right.start);
  for (let index = 1; index < spans.length; index++) {
    // SAFETY: index and index - 1 are within the array bounds.
    if (spans[index]!.start < spans[index - 1]!.end) {
      return malformed("Archive entries overlap");
    }
  }
  return Result.ok(entries);
};
