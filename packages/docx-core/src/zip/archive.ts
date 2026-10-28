/**
 * An editable ZIP archive held in memory.
 *
 * Reading parses and validates the central directory without inflating
 * anything; each entry inflates on demand and stops at its declared size, so
 * the limits checked at read bound every later read. Writing passes entries
 * that were read and not replaced through as their original compressed
 * bytes, and deflates only what changed.
 */
import { Result } from "better-result";

import { decodeText, encodeUtf8 } from "@stll/utf8-boundaries";

import { platformZipCodec, type ZipCodec } from "./codec";
import { ZipError } from "./errors";
import {
  type DosDateTime,
  encodeCentralHeader,
  encodeEndOfCentralDirectory,
  encodeEntryName,
  encodeLocalHeader,
  type EntryRecord,
  FIXED_DOS_DATE_TIME,
  fromDosDateTime,
  isSafeZipPath,
  METHOD_DEFLATE,
  METHOD_STORE,
  toDosDateTime,
  type ZipCompression,
} from "./format";
import {
  readEntries,
  type SourceEntry,
  ZIP_DEFAULT_READ_LIMITS,
  type ZipReadLimits,
} from "./read";

export type { InflateExpectation, InflateOutcome, ZipCodec } from "./codec";
export { ZipError, type ZipErrorReason } from "./errors";
export { isSafeZipPath, type ZipCompression } from "./format";
export { ZIP_DEFAULT_READ_LIMITS, type ZipReadLimits } from "./read";

/** Entry content. Text is written as UTF-8. */
export type ZipEntryData = string | Uint8Array | ArrayBuffer;

/** How {@link ZipArchive.set} writes one entry. */
export type ZipEntryOptions = {
  /** Overrides the archive-wide compression passed to `generate`. */
  readonly compression?: ZipCompression;
  /** Overrides the archive-wide DEFLATE level passed to `generate`. */
  readonly level?: number;
  /** Entry timestamp; defaults to 2000-01-01T00:00Z for reproducible output. */
  readonly modified?: Date;
};

/** How {@link ZipArchive.read} bounds an archive. */
export type ZipReadOptions = {
  /** Limits to apply; omitted fields take {@link ZIP_DEFAULT_READ_LIMITS}. */
  readonly limits?: Partial<ZipReadLimits>;
  /**
   * Entries the per-entry compression-ratio cap skips, such as binary parts
   * that legitimately compress like a run of one byte. They still count
   * toward the archive-wide ratio and every byte limit.
   */
  readonly isCompressionRatioExempt?: (path: string) => boolean;
  /** DEFLATE implementation; defaults to the platform's. */
  readonly codec?: ZipCodec;
};

/** How {@link ZipArchive.generate} writes new and replaced entries. */
export type ZipGenerateOptions = {
  /** Defaults to `"deflate"`. Entries passed through keep their compression. */
  readonly compression?: ZipCompression;
  /** DEFLATE level 0-9; defaults to 6. */
  readonly level?: number;
  /** Timestamp for every entry, replacing the ones read or set. */
  readonly modified?: Date;
};

const DEFAULT_DEFLATE_LEVEL = 6;

type StoredEntry =
  | { readonly type: "source"; readonly entry: SourceEntry }
  | {
      readonly type: "data";
      readonly bytes: Uint8Array<ArrayBuffer>;
      readonly compression: ZipCompression | undefined;
      readonly level: number | undefined;
      readonly dateTime: DosDateTime;
    };

const toEntryBytes = (data: ZipEntryData): Uint8Array<ArrayBuffer> => {
  if (typeof data === "string") return encodeUtf8(data);
  if (data instanceof ArrayBuffer) return new Uint8Array(data.slice(0));
  return data.slice();
};

const toSourceBytes = (input: Uint8Array | ArrayBuffer): Uint8Array =>
  input instanceof ArrayBuffer ? new Uint8Array(input) : input;

/** One entry of a {@link ZipArchive}. */
export type ZipEntry = {
  /** Entry name, a relative forward-slash path. */
  readonly path: string;
  /** Inflated size in bytes. */
  readonly size: number;
  /** Entry timestamp, read as UTC with two seconds of resolution. */
  readonly modified: Date;
  /**
   * The entry's bytes. A read entry inflates at most its declared size and
   * must match its CRC-32; anything else rejects with a {@link ZipError}.
   */
  readonly bytes: () => Promise<Uint8Array<ArrayBuffer>>;
  /** The entry decoded as UTF-8, keeping any byte-order mark. */
  readonly text: () => Promise<string>;
  /** The entry's bytes in an `ArrayBuffer` of exactly their length. */
  readonly arrayBuffer: () => Promise<ArrayBuffer>;
};

class ArchiveEntry implements ZipEntry {
  /** Entry name, a relative forward-slash path. */
  readonly path: string;
  /** Inflated size in bytes. */
  readonly size: number;
  /** Entry timestamp, read as UTC with two seconds of resolution. */
  readonly modified: Date;
  readonly #stored: StoredEntry;
  readonly #source: Uint8Array;
  readonly #codec: ZipCodec;

  constructor(
    path: string,
    stored: StoredEntry,
    archive: { source: Uint8Array; codec: ZipCodec },
  ) {
    this.path = path;
    this.size =
      stored.type === "source"
        ? stored.entry.uncompressedSize
        : stored.bytes.byteLength;
    this.modified = fromDosDateTime(
      stored.type === "source" ? stored.entry.dateTime : stored.dateTime,
    );
    this.#stored = stored;
    this.#source = archive.source;
    this.#codec = archive.codec;
  }

  /**
   * The entry's bytes. A read entry inflates at most its declared size and
   * must match its CRC-32; anything else rejects with a {@link ZipError}.
   */
  async bytes(): Promise<Uint8Array<ArrayBuffer>> {
    const stored = this.#stored;
    if (stored.type === "data") return stored.bytes.slice();
    const { entry } = stored;
    const raw = this.#source.subarray(
      entry.dataOffset,
      entry.dataOffset + entry.compressedSize,
    );
    let bytes: Uint8Array<ArrayBuffer>;
    if (entry.method === METHOD_STORE) {
      bytes = raw.slice();
    } else {
      const outcome = await this.#codec.inflate(raw, {
        size: entry.uncompressedSize,
        crc32: entry.crc32,
      });
      if (outcome.type === "overflow") {
        throw new ZipError({
          message: "Archive entry inflates past its declared size",
          reason: "declared-size-exceeded",
          path: entry.path,
        });
      }
      if (outcome.type === "corrupt") {
        throw new ZipError({
          message: "Archive entry data is corrupt",
          reason: "corrupt-entry",
          path: entry.path,
          cause: outcome.cause,
        });
      }
      ({ bytes } = outcome);
    }
    if (this.#codec.crc32(bytes) !== entry.crc32) {
      throw new ZipError({
        message: "Archive entry fails its CRC-32 check",
        reason: "corrupt-entry",
        path: entry.path,
      });
    }
    return bytes;
  }

  /** The entry decoded as UTF-8, keeping any byte-order mark. */
  async text(): Promise<string> {
    // Keep a leading byte-order mark: callers that parse XML decide what it means.
    return decodeText(await this.bytes(), { ignoreBOM: true });
  }

  /** The entry's bytes in an `ArrayBuffer` of exactly their length. */
  async arrayBuffer(): Promise<ArrayBuffer> {
    return (await this.bytes()).buffer;
  }
}

type PreparedEntry = {
  readonly record: EntryRecord;
  readonly payload: Uint8Array;
};

/** An editable, bounded, in-memory ZIP archive. */
export class ZipArchive {
  readonly #entries = new Map<string, StoredEntry>();
  readonly #codec: ZipCodec;
  #source: Uint8Array = new Uint8Array(0);

  constructor(options: { readonly codec?: ZipCodec } = {}) {
    this.#codec = options.codec ?? platformZipCodec();
  }

  /**
   * Read an archive. Rejects with a {@link ZipError} when the archive is
   * malformed or exceeds a limit; entries inflate later, on demand. The
   * archive keeps a reference to `input`, which must not change afterwards.
   */
  static read(
    input: Uint8Array | ArrayBuffer,
    options: ZipReadOptions = {},
  ): Promise<ZipArchive> {
    const bytes = toSourceBytes(input);
    const entries = readEntries(bytes, {
      limits: { ...ZIP_DEFAULT_READ_LIMITS, ...options.limits },
      isCompressionRatioExempt:
        options.isCompressionRatioExempt ?? (() => false),
    });
    if (Result.isError(entries)) {
      return Promise.reject(entries.error);
    }
    const archive = new ZipArchive(
      options.codec === undefined ? {} : { codec: options.codec },
    );
    archive.#source = bytes;
    for (const entry of entries.value) {
      archive.#entries.set(entry.path, { type: "source", entry });
    }
    return Promise.resolve(archive);
  }

  /** Entry names in archive order. */
  paths(): string[] {
    return [...this.#entries.keys()];
  }

  /** Entries in archive order. */
  entries(): ZipEntry[] {
    return [...this.#entries].map(([path, stored]) => this.#wrap(path, stored));
  }

  /** Whether an entry named `path` exists. */
  has(path: string): boolean {
    return this.#entries.has(path);
  }

  /** The entry named `path`, or `undefined`. */
  entry(path: string): ZipEntry | undefined {
    const stored = this.#entries.get(path);
    return stored === undefined ? undefined : this.#wrap(path, stored);
  }

  /**
   * Add or replace one entry. A replaced entry keeps its position. No parent
   * directory entries are created.
   */
  set(path: string, data: ZipEntryData, options: ZipEntryOptions = {}): void {
    if (!isSafeZipPath(path) || path.endsWith("/")) {
      throw new ZipError({
        message: "Archive entry name must be a relative file path",
        reason: "invalid-entry-name",
        path,
      });
    }
    this.#entries.set(path, {
      type: "data",
      bytes: toEntryBytes(data),
      compression: options.compression,
      level: options.level,
      dateTime:
        options.modified === undefined
          ? FIXED_DOS_DATE_TIME
          : toDosDateTime(options.modified),
    });
  }

  /**
   * An independent copy: later edits to either archive leave the other
   * unchanged. Entry content is shared, never copied, since it is immutable.
   */
  clone(): ZipArchive {
    const copy = new ZipArchive({ codec: this.#codec });
    copy.#source = this.#source;
    for (const [path, stored] of this.#entries) {
      copy.#entries.set(path, stored);
    }
    return copy;
  }

  /** Remove an entry; returns whether it existed. */
  delete(path: string): boolean {
    return this.#entries.delete(path);
  }

  /** Write the archive. Entries read and not replaced keep their bytes. */
  async generate({
    compression = "deflate",
    level = DEFAULT_DEFLATE_LEVEL,
    modified,
  }: ZipGenerateOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
    const dateTime =
      modified === undefined ? undefined : toDosDateTime(modified);
    const prepared = await Promise.all(
      [...this.#entries].map(
        async ([path, stored]) =>
          await this.#prepare({ path, stored, compression, level, dateTime }),
      ),
    );
    const locals = prepared.map(({ record }) => encodeLocalHeader(record));
    let offset = 0;
    const offsets: number[] = [];
    for (const [index, { payload }] of prepared.entries()) {
      offsets.push(offset);
      // SAFETY: `locals` maps `prepared` index for index.
      offset += locals[index]!.byteLength + payload.byteLength;
    }
    const centrals = prepared.map(({ record }, index) =>
      // SAFETY: `offsets` holds one offset per prepared entry.
      encodeCentralHeader(record, offsets[index]!),
    );
    const directorySize = centrals.reduce(
      (total, header) => total + header.byteLength,
      0,
    );
    const end = encodeEndOfCentralDirectory({
      entryCount: prepared.length,
      offset,
      size: directorySize,
    });
    const output = new Uint8Array(offset + directorySize + end.byteLength);
    let cursor = 0;
    for (const [index, { payload }] of prepared.entries()) {
      // SAFETY: `locals` maps `prepared` index for index.
      const local = locals[index]!;
      output.set(local, cursor);
      cursor += local.byteLength;
      output.set(payload, cursor);
      cursor += payload.byteLength;
    }
    for (const header of centrals) {
      output.set(header, cursor);
      cursor += header.byteLength;
    }
    output.set(end, cursor);
    return output;
  }

  #wrap(path: string, stored: StoredEntry): ZipEntry {
    return new ArchiveEntry(path, stored, {
      source: this.#source,
      codec: this.#codec,
    });
  }

  async #prepare({
    path,
    stored,
    compression,
    level,
    dateTime,
  }: {
    path: string;
    stored: StoredEntry;
    compression: ZipCompression;
    level: number;
    dateTime: DosDateTime | undefined;
  }): Promise<PreparedEntry> {
    const name = encodeEntryName(path);
    if (stored.type === "source") {
      const { entry } = stored;
      return {
        record: {
          name: name.bytes,
          flags: name.flags,
          method: entry.method,
          dateTime: dateTime ?? entry.dateTime,
          crc32: entry.crc32,
          compressedSize: entry.compressedSize,
          uncompressedSize: entry.uncompressedSize,
          directory: false,
        },
        payload: this.#source.subarray(
          entry.dataOffset,
          entry.dataOffset + entry.compressedSize,
        ),
      };
    }
    const { bytes } = stored;
    const deflate = (stored.compression ?? compression) === "deflate";
    const payload = deflate
      ? await this.#codec.deflate(bytes, stored.level ?? level)
      : bytes;
    return {
      record: {
        name: name.bytes,
        flags: name.flags,
        method: deflate ? METHOD_DEFLATE : METHOD_STORE,
        dateTime: dateTime ?? stored.dateTime,
        crc32: this.#codec.crc32(bytes),
        compressedSize: payload.byteLength,
        uncompressedSize: bytes.byteLength,
        directory: false,
      },
      payload,
    };
  }
}
