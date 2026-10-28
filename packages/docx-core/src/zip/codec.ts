/**
 * The platform's DEFLATE and CRC-32.
 *
 * Bun and Node.js expose native zlib through `process.getBuiltinModule`, which
 * needs no import a browser bundler would have to resolve. Browsers and web
 * workers use `CompressionStream` / `DecompressionStream` with
 * `"deflate-raw"` and a table CRC-32.
 */
import { Result } from "better-result";

/** How an inflation ended. `overflow` means the data inflates past `size`. */
export type InflateOutcome =
  | { readonly type: "ok"; readonly bytes: Uint8Array<ArrayBuffer> }
  | { readonly type: "overflow" }
  | { readonly type: "corrupt"; readonly cause: unknown };

/** DEFLATE and CRC-32 for one platform. */
export type ZipCodec = {
  /**
   * Inflate raw DEFLATE `data` that must produce exactly `size` bytes. Output
   * stops at `size + 1` bytes, so memory never exceeds the declared size.
   */
  readonly inflate: (data: Uint8Array, size: number) => Promise<InflateOutcome>;
  /** Raw DEFLATE at `level` (0-9). */
  readonly deflate: (
    data: Uint8Array,
    level: number,
  ) => Promise<Uint8Array<ArrayBuffer>>;
  /** CRC-32 of `data`, continuing from `previous` when given. */
  readonly crc32: (data: Uint8Array, previous?: number) => number;
};

type NativeZlib = {
  inflateRawSync: (
    data: Uint8Array,
    options: { maxOutputLength: number },
  ) => Uint8Array;
  deflateRawSync: (data: Uint8Array, options: { level: number }) => Uint8Array;
  crc32: (data: Uint8Array, previous?: number) => number;
};

const isNativeZlib = (value: unknown): value is NativeZlib =>
  typeof value === "object" &&
  value !== null &&
  "inflateRawSync" in value &&
  typeof value.inflateRawSync === "function" &&
  "deflateRawSync" in value &&
  typeof value.deflateRawSync === "function" &&
  "crc32" in value &&
  typeof value.crc32 === "function";

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

type BunHash = { crc32: (data: Uint8Array, previous?: number) => number };

const loadBunCrc32 = (): BunHash["crc32"] | null => {
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

/** `bytes` backed by an `ArrayBuffer` of exactly its length. */
export const exactBytes = (bytes: Uint8Array): Uint8Array<ArrayBuffer> => {
  const { buffer } = bytes;
  return buffer instanceof ArrayBuffer &&
    bytes.byteOffset === 0 &&
    bytes.byteLength === buffer.byteLength
    ? new Uint8Array(buffer)
    : bytes.slice();
};

const BUFFER_TOO_LARGE = "ERR_BUFFER_TOO_LARGE";

const isBufferTooLarge = (error: unknown): boolean =>
  typeof error === "object" &&
  error !== null &&
  "code" in error &&
  error.code === BUFFER_TOO_LARGE;

const nativeCodec = (zlib: NativeZlib): ZipCodec => {
  const crc32 =
    loadBunCrc32() ?? ((data, previous) => zlib.crc32(data, previous));
  return {
    inflate: (data, size) => {
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
      if (bytes.byteLength < size) {
        return Promise.resolve({
          type: "corrupt",
          cause: new RangeError("Entry inflates short of its declared size"),
        });
      }
      return Promise.resolve({ type: "ok", bytes: exactBytes(bytes) });
    },
    deflate: (data, level) =>
      Promise.resolve(exactBytes(zlib.deflateRawSync(data, { level }))),
    crc32,
  };
};

let crcTable: Uint32Array | null = null;

const tableCrc32 = (data: Uint8Array, previous = 0): number => {
  if (crcTable === null) {
    crcTable = new Uint32Array(256);
    for (let index = 0; index < 256; index++) {
      let value = index;
      for (let bit = 0; bit < 8; bit++) {
        value = value & 1 ? 0xed_b8_83_20 ^ (value >>> 1) : value >>> 1;
      }
      crcTable[index] = value >>> 0;
    }
  }
  const table = crcTable;
  let crc = ~previous >>> 0;
  for (const byte of data) {
    // SAFETY: the index is masked to 0..255 and the table has 256 entries.
    crc = table[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  }
  return ~crc >>> 0;
};

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
  data: Uint8Array,
): ReadableStreamDefaultReader<Uint8Array> => {
  const writer = transform.writable.getWriter();
  writer
    .write(exactBytes(data))
    .then(async () => await writer.close())
    .catch(() => {
      // Reported by the reader of `transform.readable`.
    });
  return transform.readable.getReader();
};

const webInflate = async (
  data: Uint8Array,
  size: number,
): Promise<InflateOutcome> => {
  const reader = pipeBytes(new DecompressionStream("deflate-raw"), data);
  const bytes = new Uint8Array(size);
  let offset = 0;
  try {
    for (;;) {
      // oxlint-disable-next-line no-await-in-loop -- a stream is read one chunk at a time
      const { done, value } = await reader.read();
      if (done) break;
      if (offset + value.byteLength > size) {
        await reader.cancel();
        return { type: "overflow" };
      }
      bytes.set(value, offset);
      offset += value.byteLength;
    }
  } catch (cause) {
    return { type: "corrupt", cause };
  }
  if (offset < size) {
    return {
      type: "corrupt",
      cause: new RangeError("Entry inflates short of its declared size"),
    };
  }
  return { type: "ok", bytes };
};

const webDeflate = async (
  data: Uint8Array,
): Promise<Uint8Array<ArrayBuffer>> => {
  const reader = pipeBytes(new CompressionStream("deflate-raw"), data);
  const chunks: Uint8Array[] = [];
  let length = 0;
  for (;;) {
    // oxlint-disable-next-line no-await-in-loop -- a stream is read one chunk at a time
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    length += value.byteLength;
  }
  return concatChunks(chunks, length);
};

/**
 * The browser codec. `CompressionStream` has no level parameter; it always
 * writes the platform's default level.
 */
export const webZipCodec: ZipCodec = {
  inflate: webInflate,
  deflate: async (data) => await webDeflate(data),
  crc32: tableCrc32,
};

let platformCodec: ZipCodec | null = null;

/** Native zlib where the runtime has it, the web streams otherwise. */
export const platformZipCodec = (): ZipCodec => {
  if (platformCodec === null) {
    const zlib = loadNativeZlib();
    platformCodec = zlib === null ? webZipCodec : nativeCodec(zlib);
  }
  return platformCodec;
};
