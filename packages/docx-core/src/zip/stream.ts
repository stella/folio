/**
 * A streaming ZIP writer for archives too large to hold in memory.
 *
 * Entries are stored, not compressed, and their bytes flow straight from the
 * input to the output: each entry's CRC-32 and size are counted as it passes
 * and written in a data descriptor after it, and Zip64 records are used when
 * a size, an offset or the entry count outgrows the classic fields. Memory
 * holds one chunk at a time plus one small central-directory record per
 * entry.
 */
import { platformZipCodec, type ZipCodec } from "./codec";
import { ZipError } from "./errors";
import {
  encodeCentralHeader,
  encodeDataDescriptor,
  encodeEndOfCentralDirectory,
  encodeEntryName,
  encodeLocalHeader,
  type EntryRecord,
  FIXED_DOS_DATE_TIME,
  FLAG_DATA_DESCRIPTOR,
  isSafeZipPath,
  METHOD_STORE,
  toDosDateTime,
} from "./format";

export { ZipError, type ZipErrorReason } from "./errors";

/** Content of one streamed entry. */
export type ZipStreamData =
  string | Uint8Array | AsyncIterable<Uint8Array> | Iterable<Uint8Array>;

/**
 * One entry of a streamed archive. A path ending in `/` is a directory entry
 * and takes no data.
 */
export type ZipStreamEntry = {
  readonly path: string;
  readonly data?: ZipStreamData;
  /** Entry timestamp; defaults to 2000-01-01T00:00Z. */
  readonly modified?: Date;
};

const utf8Encoder = new TextEncoder();

async function* entryChunks(
  data: ZipStreamData | undefined,
): AsyncGenerator<Uint8Array> {
  if (data === undefined) return;
  if (typeof data === "string") {
    yield utf8Encoder.encode(data);
    return;
  }
  if (data instanceof Uint8Array) {
    yield data;
    return;
  }
  yield* data;
}

const invalidName = (path: string): ZipError =>
  new ZipError({
    message: "Archive entry name must be a relative path",
    reason: "invalid-entry-name",
    path,
  });

/**
 * The archive bytes for `entries`, produced as they are consumed. Entries
 * are pulled one at a time, so a source that opens a download per entry
 * opens it only when the archive reaches it.
 */
export async function* zipStream(
  entries: AsyncIterable<ZipStreamEntry> | Iterable<ZipStreamEntry>,
  { codec = platformZipCodec() }: { readonly codec?: ZipCodec } = {},
): AsyncGenerator<Uint8Array> {
  const centralHeaders: Uint8Array[] = [];
  const names = new Set<string>();
  let offset = 0;
  for await (const { path, data, modified } of entries) {
    if (!isSafeZipPath(path)) throw invalidName(path);
    if (names.has(path)) {
      throw new ZipError({
        message: "Archive names an entry twice",
        reason: "duplicate-entry",
        path,
      });
    }
    names.add(path);
    const directory = path.endsWith("/");
    if (directory && data !== undefined) throw invalidName(path);
    const name = encodeEntryName(path);
    const header: EntryRecord = {
      name: name.bytes,
      flags: name.flags | FLAG_DATA_DESCRIPTOR,
      method: METHOD_STORE,
      dateTime:
        modified === undefined ? FIXED_DOS_DATE_TIME : toDosDateTime(modified),
      crc32: 0,
      compressedSize: 0,
      uncompressedSize: 0,
      directory,
    };
    const localOffset = offset;
    const local = encodeLocalHeader(header);
    yield local;
    offset += local.byteLength;

    let crc32 = 0;
    let size = 0;
    for await (const chunk of entryChunks(data)) {
      if (chunk.byteLength === 0) continue;
      crc32 = codec.crc32(chunk, crc32);
      size += chunk.byteLength;
      yield chunk;
    }
    offset += size;

    const record: EntryRecord = {
      ...header,
      crc32,
      compressedSize: size,
      uncompressedSize: size,
    };
    const descriptor = encodeDataDescriptor(record);
    yield descriptor;
    offset += descriptor.byteLength;
    centralHeaders.push(encodeCentralHeader(record, localOffset));
  }

  let directorySize = 0;
  for (const header of centralHeaders) {
    yield header;
    directorySize += header.byteLength;
  }
  yield encodeEndOfCentralDirectory({
    entryCount: centralHeaders.length,
    offset,
    size: directorySize,
  });
}

/**
 * {@link zipStream} as a `ReadableStream`. The stream pulls the next chunk
 * only when its consumer asks, so a slow consumer slows the sources rather
 * than buffering them.
 */
export const createZipStream = (
  entries: AsyncIterable<ZipStreamEntry> | Iterable<ZipStreamEntry>,
  options: { readonly codec?: ZipCodec } = {},
): ReadableStream<Uint8Array> => {
  const chunks = zipStream(entries, options);
  return new ReadableStream<Uint8Array>({
    pull: async (controller) => {
      const { done, value } = await chunks.next();
      if (done) {
        controller.close();
        return;
      }
      controller.enqueue(value);
    },
    // Ending the generator closes the entry source it is reading.
    cancel: async () => {
      await chunks.return(undefined);
    },
  });
};
