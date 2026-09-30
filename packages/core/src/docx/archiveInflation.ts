/**
 * Bounded inflation of ZIP package entries.
 *
 * Every size a ZIP archive states about itself is a claim: the central
 * directory's entry sizes and record count are written by whoever produced the
 * file, and JSZip only compares an entry's inflated length with its declared
 * one after the whole entry has been inflated into memory. The helpers here
 * turn those claims into bounds that hold while bytes are produced:
 *
 * - {@link countCentralDirectoryRecords} bounds how many entries an archive
 *   can hold before JSZip builds an object per entry.
 * - {@link exceedsCompressionRatio} refuses an entry, or a package, whose
 *   declared expansion is out of proportion to the bytes that carry it.
 * - {@link inflateEntryWithinLimits} streams one entry and stops the moment it
 *   passes its declared size, its own cap, the ratio cap or the package
 *   budget, so no limit is checked only after the allocation it guards.
 *
 * This module is platform-neutral: it uses JSZip's chunked stream, not the
 * Node.js stream wrapper, so browser and worker callers share it.
 */
import type JSZip from "jszip";

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

/**
 * Default ceiling on inflated bytes per compressed byte.
 *
 * Document parts in the repository's fixtures stay under 75:1, and parts over a
 * mebibyte under 15:1. A DEFLATE stream tops out near 1032:1, which is the
 * shape of an entry made of one repeated byte.
 */
export const DOCX_MAX_COMPRESSION_RATIO = 200;

/**
 * Inflated size below which the ratio cap does not apply.
 *
 * A small part can compress extremely well without costing anything to hold,
 * so the ratio only has to bound entries (and packages) large enough to
 * matter. The byte caps still apply below this size.
 */
export const DOCX_COMPRESSION_RATIO_GRACE_BYTES = 4 * 1024 * 1024;

const CENTRAL_DIRECTORY_RECORD_SIGNATURE = [0x50, 0x4b, 0x01, 0x02] as const;

/**
 * Upper bound on the central-directory records in `bytes`, counted without
 * parsing the archive and stopping once the count passes `stopAfter`.
 *
 * Every record JSZip reads starts with the record signature, so the number of
 * signature occurrences cannot be lower than the number of entries JSZip would
 * build, whatever the end-of-central-directory record claims.
 */
export const countCentralDirectoryRecords = (bytes: Uint8Array, stopAfter: number): number => {
  const [first, second, third, fourth] = CENTRAL_DIRECTORY_RECORD_SIGNATURE;
  let count = 0;
  let offset = bytes.indexOf(first);
  while (offset !== -1 && offset + 3 < bytes.length) {
    if (
      bytes[offset + 1] === second &&
      bytes[offset + 2] === third &&
      bytes[offset + 3] === fourth
    ) {
      count += 1;
      if (count > stopAfter) {
        return count;
      }
    }
    offset = bytes.indexOf(first, offset + 1);
  }
  return count;
};

/** Sizes a loaded entry declares in the central directory. */
export type ZipEntrySizes = {
  readonly compressedBytes: number | null;
  readonly uncompressedBytes: number | null;
};

const finiteSize = (value: unknown): number | null =>
  typeof value === "number" && Number.isFinite(value) && value >= 0 ? value : null;

/** Declared sizes of an entry JSZip loaded from an archive. */
export const getZipEntrySizes = (entry: JSZip.JSZipObject): ZipEntrySizes => {
  const data: unknown = "_data" in entry ? entry._data : undefined;
  if (typeof data !== "object" || data === null) {
    return { compressedBytes: null, uncompressedBytes: null };
  }
  return {
    compressedBytes: "compressedSize" in data ? finiteSize(data.compressedSize) : null,
    uncompressedBytes: "uncompressedSize" in data ? finiteSize(data.uncompressedSize) : null,
  };
};

const STORED_COMPRESSION_MAGIC = "\x00\x00";

/**
 * Whether a loaded entry is stored rather than compressed. A stored entry
 * reads back exactly the bytes the archive carries, so it cannot expand.
 */
export const isStoredZipEntry = (entry: JSZip.JSZipObject): boolean => {
  const data: unknown = "_data" in entry ? entry._data : undefined;
  if (typeof data !== "object" || data === null || !("compression" in data)) {
    return false;
  }
  const { compression } = data;
  return (
    typeof compression === "object" &&
    compression !== null &&
    "magic" in compression &&
    compression.magic === STORED_COMPRESSION_MAGIC
  );
};

const RATIO_BOUNDED_EXTENSIONS: ReadonlySet<string> = new Set([
  "xml",
  "rels",
  "vml",
  "txt",
  "htm",
  "html",
  "mht",
  "mhtml",
  "rtf",
]);

/**
 * Whether the per-entry ratio cap applies to a part. Markup and text parts
 * never compress anywhere near the cap. Binary parts can: an uncompressed
 * bitmap of one colour compresses like a run of one byte. Those stay bounded
 * by the byte caps and the package-wide ratio instead.
 */
export const isRatioBoundedPart = (path: string): boolean =>
  RATIO_BOUNDED_EXTENSIONS.has(path.slice(path.lastIndexOf(".") + 1).toLowerCase());

/** The per-entry ratio cap for `path`: `maxRatio`, or none for a binary part. */
export const compressionRatioLimitFor = (path: string, maxRatio: number): number =>
  isRatioBoundedPart(path) ? maxRatio : Number.POSITIVE_INFINITY;

type CompressionRatioCheck = {
  inflatedBytes: number;
  compressedBytes: number | null;
  maxRatio: number;
};

/**
 * Whether `inflatedBytes` produced from `compressedBytes` passes the ratio cap.
 * Sizes under {@link DOCX_COMPRESSION_RATIO_GRACE_BYTES} never do, and an
 * unknown compressed size is left to the byte caps.
 */
export const exceedsCompressionRatio = ({
  inflatedBytes,
  compressedBytes,
  maxRatio,
}: CompressionRatioCheck): boolean =>
  compressedBytes !== null &&
  inflatedBytes > DOCX_COMPRESSION_RATIO_GRACE_BYTES &&
  inflatedBytes > compressedBytes * maxRatio;

/**
 * Inflated bytes charged across one package. `aborted` stops every inflation
 * sharing the budget at its next chunk, once one of them has failed the
 * package.
 */
export type InflationBudget = {
  readonly maxTotalBytes: number;
  inflatedBytes: number;
  aborted: boolean;
};

export const createInflationBudget = (maxTotalBytes: number): InflationBudget => ({
  maxTotalBytes,
  inflatedBytes: 0,
  aborted: false,
});

/**
 * The bound an inflation stopped at.
 *
 * - `declared-size` — the entry produced more than its declared size.
 * - `compression-ratio` — the entry passed the ratio cap for its compressed size.
 * - `entry` — the entry passed the per-entry cap the caller set.
 * - `total` — the package passed its cumulative budget.
 * - `aborted` — another inflation sharing the budget already failed.
 */
export type InflationLimit = "declared-size" | "compression-ratio" | "entry" | "total" | "aborted";

export type InflateEntryResult =
  | { readonly ok: true; readonly bytes: Uint8Array<ArrayBuffer> }
  | { readonly ok: false; readonly limit: InflationLimit };

export type InflateEntryOptions = {
  entry: JSZip.JSZipObject;
  /** Most bytes this entry may inflate to. */
  maxEntryBytes: number;
  /** Ratio cap while streaming; see {@link compressionRatioLimitFor}. */
  maxCompressionRatio: number;
  budget: InflationBudget;
  /**
   * `false` inflates the entry only to prove it stays within its bounds; the
   * chunks are dropped and the result carries no bytes. A package part that
   * is never read here is still inflated by a later save, so the bound has
   * to be established for it too.
   */
  retain?: boolean;
};

const EMPTY_BYTES = new Uint8Array(0);

const concatChunks = (
  chunks: readonly Uint8Array[],
  totalBytes: number,
): Uint8Array<ArrayBuffer> => {
  const merged = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return merged;
};

/**
 * Inflate one entry chunk by chunk, checking every bound before a chunk is
 * kept. Pausing the stream abandons the rest of the entry, so a limit bounds
 * memory rather than reporting an overrun afterwards.
 *
 * The budget is charged as chunks arrive, which keeps concurrent inflations
 * of one package honest, and refunded when this entry stops at a limit or
 * fails, since none of its bytes are kept. Stream errors (a corrupt entry)
 * reject; limits resolve with `ok: false` for the caller to map onto its own
 * error or skip policy.
 */
export const inflateEntryWithinLimits = async ({
  entry,
  maxEntryBytes,
  maxCompressionRatio,
  budget,
  retain = true,
}: InflateEntryOptions): Promise<InflateEntryResult> => {
  const { compressedBytes, uncompressedBytes } = getZipEntrySizes(entry);
  return await new Promise<InflateEntryResult>((resolve, reject) => {
    const stream = entry.internalStream("uint8array");
    const chunks: Uint8Array[] = [];
    let entryBytes = 0;
    let settled = false;

    const settle = (): boolean => {
      if (settled) {
        return false;
      }
      settled = true;
      return true;
    };
    const refund = () => {
      budget.inflatedBytes -= entryBytes;
    };
    const stop = (limit: InflationLimit) => {
      if (!settle()) {
        return;
      }
      stream.pause();
      refund();
      resolve({ ok: false, limit });
    };
    const limitFor = (): InflationLimit | null => {
      if (budget.aborted) {
        return "aborted";
      }
      if (uncompressedBytes !== null && entryBytes > uncompressedBytes) {
        return "declared-size";
      }
      if (
        exceedsCompressionRatio({
          inflatedBytes: entryBytes,
          compressedBytes,
          maxRatio: maxCompressionRatio,
        })
      ) {
        return "compression-ratio";
      }
      if (entryBytes > maxEntryBytes) {
        return "entry";
      }
      if (budget.inflatedBytes > budget.maxTotalBytes) {
        return "total";
      }
      return null;
    };

    stream
      .on("data", (chunk) => {
        if (settled) {
          return;
        }
        entryBytes += chunk.length;
        budget.inflatedBytes += chunk.length;
        const limit = limitFor();
        if (limit !== null) {
          stop(limit);
          return;
        }
        if (retain) {
          chunks.push(chunk);
        }
      })
      .on("end", () => {
        if (!settle()) {
          return;
        }
        resolve({
          ok: true,
          bytes: retain ? concatChunks(chunks, entryBytes) : EMPTY_BYTES,
        });
      })
      .on("error", (error) => {
        if (!settle()) {
          return;
        }
        refund();
        reject(error);
      })
      .resume();
  });
};
