import { TaggedError } from "better-result";
import JSZip from "jszip";

export const DOCX_MAX_ENTRY_BYTES = 128 * 1024 * 1024;
export const DOCX_MAX_TOTAL_BYTES = 256 * 1024 * 1024;
export const DOCX_MAX_ENTRIES = 4096;

/** Error raised when a DOCX archive cannot be loaded within configured limits. */
export class DocxArchiveError extends TaggedError("DocxArchiveError")<{
  message: string;
  reason: "load-failed" | "too-many-entries" | "entry-too-large" | "total-too-large";
  cause?: unknown;
}>() {}

export type DocxArchiveOptions = {
  maxEntryBytes?: number;
  maxTotalBytes?: number;
  maxEntries?: number;
};

export type DocxArchive = {
  entries: readonly string[];
  readEntryString: (path: string) => Promise<string | null>;
  readEntryUint8: (path: string) => Promise<Uint8Array | null>;
};

type CollectStreamOptions = {
  stream: NodeJS.ReadableStream;
  maxEntryBytes: number;
  remainingBytes: number;
  maxTotalBytes: number;
  path: string;
};

const collectStream = async ({
  stream,
  maxEntryBytes,
  remainingBytes,
  maxTotalBytes,
  path,
}: CollectStreamOptions): Promise<Buffer> =>
  await new Promise<Buffer>((resolve, reject) => {
    const chunks: Buffer[] = [];
    let entryBytes = 0;

    const fail = (reason: "entry-too-large" | "total-too-large", message: string) => {
      const destroy: unknown = Reflect.get(stream, "destroy");
      if (typeof destroy === "function") {
        Reflect.apply(destroy, stream, []);
      }
      reject(new DocxArchiveError({ message, reason }));
    };

    stream.on("data", (chunk: Buffer | string) => {
      const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
      entryBytes += bytes.length;

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
      chunks.push(bytes);
    });
    stream.on("end", () => resolve(Buffer.concat(chunks)));
    stream.on("error", reject);
  });

export const loadDocxArchive = async (
  bytes: ArrayBuffer | Uint8Array,
  options: DocxArchiveOptions = {},
): Promise<DocxArchive> => {
  const maxEntryBytes = options.maxEntryBytes ?? DOCX_MAX_ENTRY_BYTES;
  const maxTotalBytes = options.maxTotalBytes ?? DOCX_MAX_TOTAL_BYTES;
  const maxEntries = options.maxEntries ?? DOCX_MAX_ENTRIES;

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
    const data = "_data" in entry ? entry._data : undefined;
    const declaredBytes =
      typeof data === "object" && data !== null && "uncompressedSize" in data
        ? data.uncompressedSize
        : undefined;
    if (typeof declaredBytes !== "number" || !Number.isFinite(declaredBytes)) {
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

  const readEntry = async (path: string): Promise<Buffer | null> => {
    const work = async (): Promise<Buffer | null> => {
      const entry = zip.file(path);
      if (!entry) {
        return null;
      }
      const buffer = await collectStream({
        stream: entry.nodeStream("nodebuffer"),
        maxEntryBytes,
        remainingBytes: maxTotalBytes - totalBytesRead,
        maxTotalBytes,
        path,
      });
      totalBytesRead += buffer.length;
      return buffer;
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
    async readEntryString(path) {
      const buffer = await readEntry(path);
      return buffer === null ? null : buffer.toString("utf-8");
    },
    async readEntryUint8(path) {
      const buffer = await readEntry(path);
      return buffer === null ? null : new Uint8Array(buffer);
    },
  };
};
