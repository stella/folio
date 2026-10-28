/** ZIP record layouts (APPNOTE 6.3.10) shared by the readers and writers. */

export const LOCAL_FILE_HEADER_SIGNATURE = 0x04_03_4b_50;
export const CENTRAL_DIRECTORY_SIGNATURE = 0x02_01_4b_50;
export const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06_05_4b_50;
export const ZIP64_END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06_06_4b_50;
export const ZIP64_END_LOCATOR_SIGNATURE = 0x07_06_4b_50;
export const DATA_DESCRIPTOR_SIGNATURE = 0x08_07_4b_50;

export const LOCAL_FILE_HEADER_SIZE = 30;
export const CENTRAL_DIRECTORY_HEADER_SIZE = 46;
export const END_OF_CENTRAL_DIRECTORY_SIZE = 22;
export const ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE = 56;
export const ZIP64_END_LOCATOR_SIZE = 20;
export const ZIP64_EXTRA_FIELD_TAG = 0x00_01;

export const METHOD_STORE = 0;
export const METHOD_DEFLATE = 8;

export const FLAG_ENCRYPTED = 0x00_01;
export const FLAG_DATA_DESCRIPTOR = 0x00_08;
export const FLAG_STRONG_ENCRYPTION = 0x00_40;
export const FLAG_UTF8_NAME = 0x08_00;

/** A 32-bit field at this value defers to the Zip64 extra field. */
export const UINT32_SENTINEL = 0xff_ff_ff_ff;
/** A 16-bit count at this value defers to the Zip64 end record. */
export const UINT16_SENTINEL = 0xff_ff;

const VERSION_DEFAULT = 20;
const VERSION_ZIP64 = 45;
const DOS_DIRECTORY_ATTRIBUTE = 0x10;

/** The compression an entry is written with. */
export type ZipCompression = "store" | "deflate";

/** DOS date and time fields, two seconds of resolution, 1980 to 2107. */
export type DosDateTime = { readonly date: number; readonly time: number };

const DOS_EPOCH_YEAR = 1980;
const DOS_LAST_YEAR = 2107;

/** `date` as DOS fields in UTC, clamped to the range DOS dates can hold. */
export const toDosDateTime = (date: Date): DosDateTime => {
  const year = date.getUTCFullYear();
  if (Number.isNaN(year) || year < DOS_EPOCH_YEAR) {
    return { date: (0 << 9) | (1 << 5) | 1, time: 0 };
  }
  if (year > DOS_LAST_YEAR) {
    return {
      date: ((DOS_LAST_YEAR - DOS_EPOCH_YEAR) << 9) | (12 << 5) | 31,
      time: (23 << 11) | (59 << 5) | 29,
    };
  }
  return {
    date:
      ((year - DOS_EPOCH_YEAR) << 9) |
      ((date.getUTCMonth() + 1) << 5) |
      date.getUTCDate(),
    time:
      (date.getUTCHours() << 11) |
      (date.getUTCMinutes() << 5) |
      (date.getUTCSeconds() >> 1),
  };
};

/** DOS fields as a UTC `Date`; out-of-range fields roll over like `Date.UTC`. */
export const fromDosDateTime = ({ date, time }: DosDateTime): Date =>
  new Date(
    Date.UTC(
      DOS_EPOCH_YEAR + (date >> 9),
      ((date >> 5) & 0x0f) - 1,
      date & 0x1f,
      time >> 11,
      (time >> 5) & 0x3f,
      (time & 0x1f) * 2,
    ),
  );

/** The reproducible timestamp written for new entries: 2000-01-01T00:00Z. */
export const FIXED_DOS_DATE_TIME: DosDateTime = toDosDateTime(
  new Date(Date.UTC(2000, 0, 1)),
);

const utf8Encoder = new TextEncoder();

/** Entry name bytes and the general-purpose flag they need. */
export const encodeEntryName = (
  path: string,
): { readonly bytes: Uint8Array<ArrayBuffer>; readonly flags: number } => {
  const bytes = utf8Encoder.encode(path);
  // ASCII names need no flag; any other byte is UTF-8 and says so.
  const ascii = bytes.every((byte) => byte < 0x80);
  return { bytes, flags: ascii ? 0 : FLAG_UTF8_NAME };
};

/** Whether `path` is a relative, forward-slash name with no `..` segment. */
export const isSafeZipPath = (path: string): boolean => {
  if (path.length === 0 || path.startsWith("/")) return false;
  if (path.includes("\\") || path.includes("\0")) return false;
  if (/^[A-Za-z]:/u.test(path)) return false;
  return !path.split("/").some((segment) => segment === "..");
};

/** What one record needs to describe an entry. */
export type EntryRecord = {
  readonly name: Uint8Array;
  readonly flags: number;
  readonly method: number;
  readonly dateTime: DosDateTime;
  readonly crc32: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly directory: boolean;
};

const writeCommonFields = (
  view: DataView,
  offset: number,
  record: EntryRecord,
): void => {
  view.setUint16(offset, record.flags, true);
  view.setUint16(offset + 2, record.method, true);
  view.setUint16(offset + 4, record.dateTime.time, true);
  view.setUint16(offset + 6, record.dateTime.date, true);
  view.setUint32(offset + 8, record.crc32, true);
};

const needsZip64 = (value: number): boolean => value >= UINT32_SENTINEL;

/**
 * A local file header. Sizes go in the header unless the entry streams them
 * in a data descriptor; a header that carries sizes of 4 GiB or more puts
 * both in a Zip64 extra field, as APPNOTE 4.5.3 requires.
 */
export const encodeLocalHeader = (
  record: EntryRecord,
): Uint8Array<ArrayBuffer> => {
  const streamed = (record.flags & FLAG_DATA_DESCRIPTOR) !== 0;
  const zip64 =
    !streamed &&
    (needsZip64(record.compressedSize) || needsZip64(record.uncompressedSize));
  const extraLength = zip64 ? 20 : 0;
  const bytes = new Uint8Array(
    LOCAL_FILE_HEADER_SIZE + record.name.byteLength + extraLength,
  );
  const view = new DataView(bytes.buffer);
  view.setUint32(0, LOCAL_FILE_HEADER_SIGNATURE, true);
  view.setUint16(4, zip64 ? VERSION_ZIP64 : VERSION_DEFAULT, true);
  writeCommonFields(view, 6, record);
  if (streamed) {
    view.setUint32(14, 0, true);
  } else {
    view.setUint32(18, zip64 ? UINT32_SENTINEL : record.compressedSize, true);
    view.setUint32(22, zip64 ? UINT32_SENTINEL : record.uncompressedSize, true);
  }
  view.setUint16(26, record.name.byteLength, true);
  view.setUint16(28, extraLength, true);
  bytes.set(record.name, LOCAL_FILE_HEADER_SIZE);
  if (zip64) {
    const extra = LOCAL_FILE_HEADER_SIZE + record.name.byteLength;
    view.setUint16(extra, ZIP64_EXTRA_FIELD_TAG, true);
    view.setUint16(extra + 2, 16, true);
    view.setBigUint64(extra + 4, BigInt(record.uncompressedSize), true);
    view.setBigUint64(extra + 12, BigInt(record.compressedSize), true);
  }
  return bytes;
};

/**
 * The data descriptor that follows a streamed entry. Sizes of 4 GiB or more
 * are written as 8-byte fields.
 */
export const encodeDataDescriptor = (
  record: EntryRecord,
): Uint8Array<ArrayBuffer> => {
  const zip64 =
    needsZip64(record.compressedSize) || needsZip64(record.uncompressedSize);
  const bytes = new Uint8Array(zip64 ? 24 : 16);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, DATA_DESCRIPTOR_SIGNATURE, true);
  view.setUint32(4, record.crc32, true);
  if (zip64) {
    view.setBigUint64(8, BigInt(record.compressedSize), true);
    view.setBigUint64(16, BigInt(record.uncompressedSize), true);
  } else {
    view.setUint32(8, record.compressedSize, true);
    view.setUint32(12, record.uncompressedSize, true);
  }
  return bytes;
};

/**
 * A central directory header. Each field of 4 GiB or more is written as the
 * sentinel and carried, in APPNOTE order, by a Zip64 extra field.
 */
export const encodeCentralHeader = (
  record: EntryRecord,
  localOffset: number,
): Uint8Array<ArrayBuffer> => {
  const overflow = [
    record.uncompressedSize,
    record.compressedSize,
    localOffset,
  ].filter((value) => needsZip64(value));
  const extraLength = overflow.length === 0 ? 0 : 4 + overflow.length * 8;
  const bytes = new Uint8Array(
    CENTRAL_DIRECTORY_HEADER_SIZE + record.name.byteLength + extraLength,
  );
  const view = new DataView(bytes.buffer);
  const version = extraLength === 0 ? VERSION_DEFAULT : VERSION_ZIP64;
  view.setUint32(0, CENTRAL_DIRECTORY_SIGNATURE, true);
  view.setUint16(4, version, true);
  view.setUint16(6, version, true);
  writeCommonFields(view, 8, record);
  view.setUint32(
    20,
    needsZip64(record.compressedSize) ? UINT32_SENTINEL : record.compressedSize,
    true,
  );
  view.setUint32(
    24,
    needsZip64(record.uncompressedSize)
      ? UINT32_SENTINEL
      : record.uncompressedSize,
    true,
  );
  view.setUint16(28, record.name.byteLength, true);
  view.setUint16(30, extraLength, true);
  view.setUint32(38, record.directory ? DOS_DIRECTORY_ATTRIBUTE : 0, true);
  view.setUint32(
    42,
    needsZip64(localOffset) ? UINT32_SENTINEL : localOffset,
    true,
  );
  bytes.set(record.name, CENTRAL_DIRECTORY_HEADER_SIZE);
  if (extraLength !== 0) {
    let field = CENTRAL_DIRECTORY_HEADER_SIZE + record.name.byteLength;
    view.setUint16(field, ZIP64_EXTRA_FIELD_TAG, true);
    view.setUint16(field + 2, overflow.length * 8, true);
    field += 4;
    for (const value of overflow) {
      view.setBigUint64(field, BigInt(value), true);
      field += 8;
    }
  }
  return bytes;
};

/** Where the central directory sits and how many records it holds. */
export type DirectoryLocation = {
  readonly entryCount: number;
  readonly offset: number;
  readonly size: number;
};

/**
 * The end-of-central-directory record, preceded by the Zip64 end record and
 * its locator when a count, size or offset does not fit the classic fields.
 */
export const encodeEndOfCentralDirectory = ({
  entryCount,
  offset,
  size,
}: DirectoryLocation): Uint8Array<ArrayBuffer> => {
  const zip64 =
    entryCount >= UINT16_SENTINEL || needsZip64(offset) || needsZip64(size);
  const zip64Length = zip64
    ? ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE + ZIP64_END_LOCATOR_SIZE
    : 0;
  const bytes = new Uint8Array(zip64Length + END_OF_CENTRAL_DIRECTORY_SIZE);
  const view = new DataView(bytes.buffer);
  if (zip64) {
    const recordOffset = offset + size;
    view.setUint32(0, ZIP64_END_OF_CENTRAL_DIRECTORY_SIGNATURE, true);
    view.setBigUint64(
      4,
      BigInt(ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE - 12),
      true,
    );
    view.setUint16(12, VERSION_ZIP64, true);
    view.setUint16(14, VERSION_ZIP64, true);
    view.setBigUint64(24, BigInt(entryCount), true);
    view.setBigUint64(32, BigInt(entryCount), true);
    view.setBigUint64(40, BigInt(size), true);
    view.setBigUint64(48, BigInt(offset), true);
    const locator = ZIP64_END_OF_CENTRAL_DIRECTORY_SIZE;
    view.setUint32(locator, ZIP64_END_LOCATOR_SIGNATURE, true);
    view.setBigUint64(locator + 8, BigInt(recordOffset), true);
    view.setUint32(locator + 16, 1, true);
  }
  const end = zip64Length;
  view.setUint32(end, END_OF_CENTRAL_DIRECTORY_SIGNATURE, true);
  const count = Math.min(entryCount, UINT16_SENTINEL);
  view.setUint16(end + 8, count, true);
  view.setUint16(end + 10, count, true);
  view.setUint32(end + 12, needsZip64(size) ? UINT32_SENTINEL : size, true);
  view.setUint32(end + 16, needsZip64(offset) ? UINT32_SENTINEL : offset, true);
  return bytes;
};
