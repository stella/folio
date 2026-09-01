import { TaggedError } from "better-result";
import JSZip from "jszip";

declare module "jszip" {
  // oxlint-disable-next-line typescript/consistent-type-definitions -- declaration merging requires an interface
  interface JSZipObject {
    /**
     * Chunked read of the entry content, missing from the published typings.
     * `nodeStream` is this stream wrapped in a Node.js `Readable`, which
     * browsers and web workers cannot provide; the stream itself is
     * platform-neutral.
     */
    internalStream(type: "uint8array"): JSZip.JSZipStreamHelper<Uint8Array>;
  }
}

export const DOCX_MAX_ENTRY_BYTES = 128 * 1024 * 1024;
export const DOCX_MAX_TOTAL_BYTES = 256 * 1024 * 1024;
export const DOCX_MAX_ENTRIES = 4096;
export const DOCX_MAX_INPUT_BYTES = 50 * 1024 * 1024;

/** Error raised when a DOCX archive cannot be loaded within configured limits. */
export class DocxArchiveError extends TaggedError("DocxArchiveError")<{
  message: string;
  reason:
    | "load-failed"
    | "input-too-large"
    | "too-many-entries"
    | "entry-too-large"
    | "total-too-large"
    | "invalid-options";
  cause?: unknown;
}> {}

export type DocxArchiveOptions = {
  maxInputBytes?: number;
  maxEntryBytes?: number;
  maxTotalBytes?: number;
  maxEntries?: number;
};

export type DocxArchiveEntry = {
  readonly path: string;
  readonly directory: boolean;
  readonly declaredUncompressedBytes: number | null;
};

export type DocxArchiveReadOptions = {
  maxBytes?: number;
};

export type DocxArchive = {
  entries: readonly string[];
  entryMetadata: readonly DocxArchiveEntry[];
  readEntryString: (path: string) => Promise<string | null>;
  readEntryUint8: (path: string, options?: DocxArchiveReadOptions) => Promise<Uint8Array | null>;
};

const concatChunks = (chunks: readonly Uint8Array[]): Uint8Array => {
  let totalBytes = 0;
  for (const chunk of chunks) {
    totalBytes += chunk.length;
  }
  const merged = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return merged;
};

type CollectStreamOptions = {
  stream: JSZip.JSZipStreamHelper<Uint8Array>;
  maxEntryBytes: number;
  remainingBytes: number;
  maxTotalBytes: number;
  path: string;
};

/**
 * Accumulate an entry chunk by chunk, checking both caps before each chunk is
 * retained. Pausing the stream abandons a decompression bomb mid-inflate, so
 * the caps bound memory instead of merely reporting the overrun afterwards.
 */
const collectStream = async ({
  stream,
  maxEntryBytes,
  remainingBytes,
  maxTotalBytes,
  path,
}: CollectStreamOptions): Promise<Uint8Array> =>
  await new Promise<Uint8Array>((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let entryBytes = 0;

    const fail = (reason: "entry-too-large" | "total-too-large", message: string) => {
      stream.pause();
      reject(new DocxArchiveError({ message, reason }));
    };

    stream
      .on("data", (chunk) => {
        entryBytes += chunk.length;

        if (entryBytes > maxEntryBytes) {
          fail("entry-too-large", `DOCX entry "${path}" exceeded the ${maxEntryBytes}-byte limit`);
          return;
        }
        if (entryBytes > remainingBytes) {
          fail(
            "total-too-large",
            `DOCX archive exceeded the ${maxTotalBytes}-byte cumulative limit while reading "${path}"`,
          );
          return;
        }
        chunks.push(chunk);
      })
      .on("end", () => resolve(concatChunks(chunks)))
      .on("error", reject)
      .resume();
  });

const getDeclaredUncompressedBytes = (entry: JSZip.JSZipObject): number | null => {
  const data = "_data" in entry ? entry._data : undefined;
  const declaredBytes =
    typeof data === "object" && data !== null && "uncompressedSize" in data
      ? data.uncompressedSize
      : undefined;
  return typeof declaredBytes === "number" && Number.isFinite(declaredBytes) ? declaredBytes : null;
};

type ResolveByteLimitOptions = {
  value: number | undefined;
  fallback: number;
  name: string;
};

const resolveByteLimit = ({ value, fallback, name }: ResolveByteLimitOptions): number => {
  const limit = value ?? fallback;
  if (!Number.isSafeInteger(limit) || limit < 0) {
    throw new DocxArchiveError({
      message: `${name} must be a non-negative safe integer`,
      reason: "invalid-options",
    });
  }
  return limit;
};

export const loadDocxArchive = async (
  bytes: ArrayBuffer | Uint8Array,
  options: DocxArchiveOptions = {},
): Promise<DocxArchive> => {
  const maxInputBytes = resolveByteLimit({
    value: options.maxInputBytes,
    fallback: DOCX_MAX_INPUT_BYTES,
    name: "DOCX input byte limit",
  });
  const maxEntryBytes = resolveByteLimit({
    value: options.maxEntryBytes,
    fallback: DOCX_MAX_ENTRY_BYTES,
    name: "DOCX entry byte limit",
  });
  const maxTotalBytes = resolveByteLimit({
    value: options.maxTotalBytes,
    fallback: DOCX_MAX_TOTAL_BYTES,
    name: "DOCX cumulative byte limit",
  });
  const maxEntries = resolveByteLimit({
    value: options.maxEntries,
    fallback: DOCX_MAX_ENTRIES,
    name: "DOCX entry limit",
  });

  if (bytes.byteLength > maxInputBytes) {
    throw new DocxArchiveError({
      message: `DOCX input contains ${bytes.byteLength} bytes (max ${maxInputBytes})`,
      reason: "input-too-large",
    });
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (cause) {
    throw new DocxArchiveError({
      message: "Failed to parse DOCX archive",
      reason: "load-failed",
      cause,
    });
  }

  const archiveEntries = Object.values(zip.files);
  if (archiveEntries.length > maxEntries) {
    throw new DocxArchiveError({
      message: `DOCX archive declares ${archiveEntries.length} entries (max ${maxEntries})`,
      reason: "too-many-entries",
    });
  }

  let declaredTotalBytes = 0;
  for (const entry of archiveEntries) {
    if (entry.dir) {
      continue;
    }
    const declaredBytes = getDeclaredUncompressedBytes(entry);
    if (declaredBytes === null) {
      declaredTotalBytes = Number.NaN;
      break;
    }
    if (declaredBytes > maxEntryBytes) {
      throw new DocxArchiveError({
        message: `DOCX entry "${entry.name}" declares ${declaredBytes} bytes (max ${maxEntryBytes})`,
        reason: "entry-too-large",
      });
    }
    declaredTotalBytes += declaredBytes;
  }

  if (Number.isFinite(declaredTotalBytes) && declaredTotalBytes > maxTotalBytes) {
    throw new DocxArchiveError({
      message: `DOCX archive declares ${declaredTotalBytes} cumulative bytes (max ${maxTotalBytes})`,
      reason: "total-too-large",
    });
  }

  let totalBytesRead = 0;
  let readChain: Promise<unknown> = Promise.resolve();

  const readEntry = async (
    path: string,
    readOptions: DocxArchiveReadOptions = {},
  ): Promise<Uint8Array | null> => {
    const requestedMaxBytes = resolveByteLimit({
      value: readOptions.maxBytes,
      fallback: maxEntryBytes,
      name: "DOCX entry read byte limit",
    });
    const work = async (): Promise<Uint8Array | null> => {
      const entry = zip.file(path);
      if (!entry) {
        return null;
      }
      const content = await collectStream({
        stream: entry.internalStream("uint8array"),
        maxEntryBytes: Math.min(requestedMaxBytes, maxEntryBytes),
        remainingBytes: maxTotalBytes - totalBytesRead,
        maxTotalBytes,
        path,
      });
      totalBytesRead += content.length;
      return content;
    };

    const next = readChain.then(work, work);
    readChain = next.then(
      () => undefined,
      () => undefined,
    );
    return await next;
  };

  return {
    entries: Object.freeze(archiveEntries.map(({ name }) => name)),
    entryMetadata: Object.freeze(
      archiveEntries.map((entry) => ({
        path: entry.name,
        directory: entry.dir,
        declaredUncompressedBytes: getDeclaredUncompressedBytes(entry),
      })),
    ),
    async readEntryString(path) {
      const content = await readEntry(path);
      // `ignoreBOM` keeps a leading U+FEFF in the string: OOXML parts written
      // by Word carry a UTF-8 BOM, and callers that splice a part and write it
      // back must not silently drop it.
      return content === null
        ? null
        : new TextDecoder("utf-8", { ignoreBOM: true }).decode(content);
    },
    readEntryUint8: readEntry,
  };
};
