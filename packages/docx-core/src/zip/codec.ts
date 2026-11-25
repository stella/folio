/**
 * The platform's DEFLATE and CRC-32.
 *
 * Bun and Node.js expose native zlib through `process.getBuiltinModule`, which
 * needs no import a browser bundler would have to resolve. Everywhere else
 * (browsers, web workers, and runtimes too old for `getBuiltinModule`) uses
 * `CompressionStream` / `DecompressionStream` and a table CRC-32. Runtimes
 * whose streams lack `"deflate-raw"` get the same raw DEFLATE data framed as
 * gzip: a fixed ten-byte header before it and the entry's CRC-32 and size,
 * known from the central directory, after it.
 */
import { Result } from "better-result";

/** How an inflation ended. `overflow` means the data inflates past `size`. */
export type InflateOutcome =
  | { readonly type: "ok"; readonly bytes: Uint8Array<ArrayBuffer> }
  | { readonly type: "overflow" }
  | { readonly type: "corrupt"; readonly cause: unknown };

/** What an entry declares about its inflated content. */
export type InflateExpectation = {
  /** Inflated size in bytes. */
  readonly size: number;
  /** CRC-32 of the inflated content. */
  readonly crc32: number;
};

/** DEFLATE and CRC-32 for one platform. */
export type ZipCodec = {
  /**
   * Inflate raw DEFLATE `data` that must produce exactly `expected.size`
   * bytes. Output stops at that size, so memory never exceeds it.
   */
  readonly inflate: (
    data: Uint8Array,
    expected: InflateExpectation,
  ) => Promise<InflateOutcome>;
  /** Raw DEFLATE at `level` (0-9). */
  readonly deflate: (
    data: Uint8Array,
    level: number,
  ) => Promise<Uint8Array<ArrayBuffer>>;
  /** CRC-32 of `data`, continuing from `previous` when given. */
  readonly crc32: (data: Uint8Array, previous?: number) => number;
};

type Crc32 = ZipCodec["crc32"];

type NativeZlib = {
  inflateRawSync: (
    data: Uint8Array,
    options: { maxOutputLength: number },
  ) => Uint8Array;
  deflateRawSync: (data: Uint8Array, options: { level: number }) => Uint8Array;
};

const isNativeZlib = (value: unknown): value is NativeZlib =>
  typeof value === "object" &&
  value !== null &&
  "inflateRawSync" in value &&
  typeof value.inflateRawSync === "function" &&
  "deflateRawSync" in value &&
  typeof value.deflateRawSync === "function";

const nativeZlibCrc32 = (zlib: object): Crc32 | null => {
  if (!("crc32" in zlib) || typeof zlib.crc32 !== "function") return null;
  const { crc32 } = zlib;
  return (data, previous) => crc32(data, previous);
};

const loadNativeZlib = (): NativeZlib | null => {
  const runtimeProcess: unknown = Reflect.get(globalThis, "process");
  if (
    typeof runtimeProcess !== "object" ||
    runtimeProcess === null ||
    !("getBuiltinModule" in runtimeProcess) ||
    typeof runtimeProcess.getBuiltinModule !== "function"
  ) {
    return null;
  }
  const zlib: unknown = runtimeProcess.getBuiltinModule("node:zlib");
  return isNativeZlib(zlib) ? zlib : null;
};

const loadBunCrc32 = (): Crc32 | null => {
  const bun: unknown = Reflect.get(globalThis, "Bun");
  if (typeof bun !== "object" || bun === null || !("hash" in bun)) return null;
  const { hash } = bun;
  if (
    typeof hash !== "function" ||
    !("crc32" in hash) ||
    typeof hash.crc32 !== "function"
  ) {
    return null;
  }
  const { crc32 } = hash;
  return (data, previous) => crc32(data, previous);
};

/**
 * A copy of `bytes` in a new `ArrayBuffer` of exactly its length. Node.js
 * `Buffer#slice` returns a view into a shared pool, so copies never go
 * through `slice`.
 */
export const copyBytes = (bytes: Uint8Array): Uint8Array<ArrayBuffer> =>
  new Uint8Array(bytes);

/** `bytes` backed by an `ArrayBuffer` of exactly its length. */
export const exactBytes = (bytes: Uint8Array): Uint8Array<ArrayBuffer> => {
  const { buffer } = bytes;
  return buffer instanceof ArrayBuffer &&
    bytes.byteOffset === 0 &&
    bytes.byteLength === buffer.byteLength
    ? new Uint8Array(buffer)
    : copyBytes(bytes);
};

/**
 * Eight 256-entry CRC-32 tables, one after another: table 0 is the classic
 * byte table, and table `k` advances a byte's contribution by `k` more bytes.
 */
let crcTables: Uint32Array | null = null;

const buildCrcTables = (): Uint32Array => {
  const tables = new Uint32Array(256 * 8);
  for (let index = 0; index < 256; index++) {
    let value = index;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? 0xed_b8_83_20 ^ (value >>> 1) : value >>> 1;
    }
    tables[index] = value >>> 0;
  }
  for (let index = 0; index < 256; index++) {
    // SAFETY: every index below is in 0..2047, inside the 2048 entries.
    let value = tables[index]!;
    for (let table = 1; table < 8; table++) {
      value = tables[value & 0xff]! ^ (value >>> 8);
      tables[table * 256 + index] = value >>> 0;
    }
  }
  return tables;
};

/**
 * CRC-32 where the runtime offers none (browsers): eight bytes per step over
 * eight tables, then byte by byte. The loops index the array; iterating a
 * typed array with `for…of` runs an order of magnitude slower, and every
 * entry an archive inflates is checked through here.
 */
const tableCrc32: Crc32 = (data, previous = 0) => {
  crcTables ??= buildCrcTables();
  const t = crcTables;
  let crc = ~previous >>> 0;
  const length = data.length;
  const whole = length - (length % 8);
  let index = 0;
  // SAFETY: table indices are masked to 0..255 and offset by k * 256 for
  // k in 0..7, inside the 2048 entries; data indices stay below `length`.
  for (; index < whole; index += 8) {
    const low =
      (crc ^
        (data[index]! |
          (data[index + 1]! << 8) |
          (data[index + 2]! << 16) |
          (data[index + 3]! << 24))) >>>
      0;
    crc =
      t[1792 + (low & 0xff)]! ^
      t[1536 + ((low >>> 8) & 0xff)]! ^
      t[1280 + ((low >>> 16) & 0xff)]! ^
      t[1024 + (low >>> 24)]! ^
      t[768 + data[index + 4]!]! ^
      t[512 + data[index + 5]!]! ^
      t[256 + data[index + 6]!]! ^
      t[data[index + 7]!]!;
  }
  for (; index < length; index++) {
    crc = t[(crc ^ data[index]!) & 0xff]! ^ (crc >>> 8);
  }
  return ~crc >>> 0;
};

const BUFFER_TOO_LARGE = "ERR_BUFFER_TOO_LARGE";

const isBufferTooLarge = (error: unknown): boolean =>
  typeof error === "object" &&
  error !== null &&
  "code" in error &&
  error.code === BUFFER_TOO_LARGE;

const shortOf = (size: number): InflateOutcome => ({
  type: "corrupt",
  cause: new RangeError(`Entry inflates short of its declared ${size} bytes`),
});

const nativeCodec = (zlib: NativeZlib): ZipCodec => ({
  inflate: (data, { size }) => {
    // `maxOutputLength` must be at least 1; an empty entry is checked by
    // length below like any other.
    const inflated = Result.try(() =>
      zlib.inflateRawSync(data, { maxOutputLength: Math.max(1, size) }),
    );
    if (Result.isError(inflated)) {
      const cause = inflated.error.cause;
      return Promise.resolve(
        isBufferTooLarge(cause)
          ? { type: "overflow" }
          : { type: "corrupt", cause },
      );
    }
    const bytes = inflated.value;
    if (bytes.byteLength > size) return Promise.resolve({ type: "overflow" });
    if (bytes.byteLength < size) return Promise.resolve(shortOf(size));
    return Promise.resolve({ type: "ok", bytes: exactBytes(bytes) });
  },
  deflate: (data, level) =>
    Promise.resolve(exactBytes(zlib.deflateRawSync(data, { level }))),
  crc32: loadBunCrc32() ?? nativeZlibCrc32(zlib) ?? tableCrc32,
});

const concatChunks = (
  chunks: readonly Uint8Array[],
  length: number,
): Uint8Array<ArrayBuffer> => {
  const merged = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return merged;
};

/**
 * Feed `data` through a transform without awaiting the write, so the reader
 * drains the output while the input is still being written. The write
 * settles with the readable side; its rejection carries the same failure the
 * read loop reports, so it is observed here rather than left unhandled.
 */
const pipeBytes = (
  transform: ReadableWritablePair<Uint8Array, BufferSource>,
  chunks: readonly Uint8Array[],
): ReadableStreamDefaultReader<Uint8Array> => {
  const writer = transform.writable.getWriter();
  const write = async () => {
    for (const chunk of chunks) {
      // oxlint-disable-next-line no-await-in-loop -- chunks are written in order
      await writer.write(exactBytes(chunk));
    }
    await writer.close();
  };
  write().catch(() => {
    // Reported by the reader of `transform.readable`.
  });
  return transform.readable.getReader();
};

/** The streams' raw DEFLATE, or gzip framing where they lack it. */
export type StreamFraming = "deflate-raw" | "gzip";

/** Whether this runtime's streams accept `"deflate-raw"`. */
const detectFraming = (): StreamFraming =>
  Result.isOk(Result.try(() => new DecompressionStream("deflate-raw")))
    ? "deflate-raw"
    : "gzip";

// ID1 ID2, CM = deflate, no flags, no mtime, no extra flags, OS unknown.
const GZIP_HEADER = Uint8Array.of(0x1f, 0x8b, 8, 0, 0, 0, 0, 0, 0, 0xff);
const GZIP_HEADER_SIZE = GZIP_HEADER.byteLength;
const GZIP_TRAILER_SIZE = 8;

const gzipTrailer = ({ size, crc32 }: InflateExpectation): Uint8Array => {
  const trailer = new Uint8Array(GZIP_TRAILER_SIZE);
  const view = new DataView(trailer.buffer);
  view.setUint32(0, crc32, true);
  view.setUint32(4, size % 0x1_00_00_00_00, true);
  return trailer;
};

const streamInflate = async (
  framing: StreamFraming,
  data: Uint8Array,
  expected: InflateExpectation,
): Promise<InflateOutcome> => {
  const { size } = expected;
  const reader = pipeBytes(
    new DecompressionStream(framing),
    framing === "gzip" ? [GZIP_HEADER, data, gzipTrailer(expected)] : [data],
  );
  const bytes = new Uint8Array(size);
  let offset = 0;
  const read = await Result.tryPromise(async () => {
    for (;;) {
      // oxlint-disable-next-line no-await-in-loop -- a stream is read one chunk at a time
      const { done, value } = await reader.read();
      if (done) return "ended" as const;
      if (offset + value.byteLength > size) {
        await reader.cancel();
        return "overflow" as const;
      }
      bytes.set(value, offset);
      offset += value.byteLength;
    }
  });
  if (Result.isError(read)) return { type: "corrupt", cause: read.error.cause };
  if (read.value === "overflow") return { type: "overflow" };
  if (offset < size) return shortOf(size);
  return { type: "ok", bytes };
};

const streamDeflate = async (
  framing: StreamFraming,
  data: Uint8Array,
): Promise<Uint8Array<ArrayBuffer>> => {
  const reader = pipeBytes(new CompressionStream(framing), [data]);
  const chunks: Uint8Array[] = [];
  let length = 0;
  for (;;) {
    // oxlint-disable-next-line no-await-in-loop -- a stream is read one chunk at a time
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    length += value.byteLength;
  }
  const output = concatChunks(chunks, length);
  // CompressionStream writes the fixed ten-byte gzip header (no optional
  // fields) and an eight-byte trailer; the raw DEFLATE data is in between.
  return framing === "gzip"
    ? output.slice(GZIP_HEADER_SIZE, output.byteLength - GZIP_TRAILER_SIZE)
    : output;
};

/**
 * A codec over the web compression streams with the given framing.
 * `CompressionStream` has no level parameter; it writes the platform's
 * default level.
 */
export const streamZipCodec = (framing: StreamFraming): ZipCodec => ({
  inflate: async (data, expected) =>
    await streamInflate(framing, data, expected),
  deflate: async (data) => await streamDeflate(framing, data),
  crc32: tableCrc32,
});

let webCodec: ZipCodec | null = null;

/** The web compression streams, in the framing this runtime supports. */
export const webZipCodec = (): ZipCodec => {
  webCodec ??= streamZipCodec(detectFraming());
  return webCodec;
};

let platformCodec: ZipCodec | null = null;

/** Native zlib where the runtime has it, the web streams otherwise. */
export const platformZipCodec = (): ZipCodec => {
  if (platformCodec === null) {
    const zlib = loadNativeZlib();
    platformCodec = zlib === null ? webZipCodec() : nativeCodec(zlib);
  }
  return platformCodec;
};
