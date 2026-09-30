import { TaggedError } from "better-result";
import JSZip from "jszip";

import {
  compressionRatioLimitFor,
  countCentralDirectoryRecords,
  createInflationBudget,
  DOCX_MAX_COMPRESSION_RATIO,
  exceedsCompressionRatio,
  getZipEntrySizes,
  inflateEntryWithinLimits,
  type InflationLimit,
} from "../archiveInflation";
import {
  assertXmlResourceLimits,
  createXmlPackageBudget,
  FOLIO_XML_RESOURCE_LIMITS,
  type XmlResourceLimits,
} from "../xmlResourceLimits";

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
    | "compression-ratio-exceeded"
    | "invalid-options";
  cause?: unknown;
}> {}

export type DocxArchiveOptions = {
  maxInputBytes?: number;
  maxEntryBytes?: number;
  maxTotalBytes?: number;
  maxEntries?: number;
  /**
   * Most inflated bytes allowed per compressed byte, for each markup or text
   * entry and for the package as a whole. Binary entries are bounded by the
   * byte limits and the package ratio only, and entries and packages under
   * 4 MiB inflated are exempt. Defaults to 200.
   */
  maxCompressionRatio?: number;
  /**
   * Bounds on parsed XML structure, applied to every XML part this archive
   * hands out as a string.
   *
   * Every consumer of `readEntryString` parses the result into an object tree,
   * and the byte caps above bound the markup, not the tree. Enforcing the
   * element and attribute bounds here rather than at each parse site means a
   * new consumer is bounded by construction: there is no way to obtain a part
   * string from this archive that has not been counted.
   */
  xmlLimits?: Partial<XmlResourceLimits>;
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

type ReadLimitErrorOptions = {
  limit: InflationLimit;
  path: string;
  maxEntryBytes: number;
  maxTotalBytes: number;
  maxCompressionRatio: number;
};

const readLimitError = ({
  limit,
  path,
  maxEntryBytes,
  maxTotalBytes,
  maxCompressionRatio,
}: ReadLimitErrorOptions): DocxArchiveError => {
  switch (limit) {
    case "declared-size":
      return new DocxArchiveError({
        message: `DOCX entry "${path}" inflated past its declared size`,
        reason: "entry-too-large",
      });
    case "compression-ratio":
      return new DocxArchiveError({
        message: `DOCX entry "${path}" exceeded the ${maxCompressionRatio}:1 compression ratio limit`,
        reason: "compression-ratio-exceeded",
      });
    case "entry":
      return new DocxArchiveError({
        message: `DOCX entry "${path}" exceeded the ${maxEntryBytes}-byte limit`,
        reason: "entry-too-large",
      });
    case "total":
    case "aborted":
      return new DocxArchiveError({
        message: `DOCX archive exceeded the ${maxTotalBytes}-byte cumulative limit while reading "${path}"`,
        reason: "total-too-large",
      });
  }
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

const asBytes = (bytes: ArrayBuffer | Uint8Array): Uint8Array =>
  bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);

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
  const maxCompressionRatio = resolveByteLimit({
    value: options.maxCompressionRatio,
    fallback: DOCX_MAX_COMPRESSION_RATIO,
    name: "DOCX compression ratio limit",
  });

  if (bytes.byteLength > maxInputBytes) {
    throw new DocxArchiveError({
      message: `DOCX input contains ${bytes.byteLength} bytes (max ${maxInputBytes})`,
      reason: "input-too-large",
    });
  }

  // Counted before JSZip builds an object per record: the entry cap has to
  // bound that allocation, not only what is inflated afterwards.
  if (countCentralDirectoryRecords(asBytes(bytes), maxEntries) > maxEntries) {
    throw new DocxArchiveError({
      message: `DOCX archive holds more than ${maxEntries} entries`,
      reason: "too-many-entries",
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
  let declaredTotalKnown = true;
  for (const entry of archiveEntries) {
    if (entry.dir) {
      continue;
    }
    const { compressedBytes, uncompressedBytes } = getZipEntrySizes(entry);
    if (uncompressedBytes === null) {
      declaredTotalKnown = false;
      continue;
    }
    if (uncompressedBytes > maxEntryBytes) {
      throw new DocxArchiveError({
        message: `DOCX entry "${entry.name}" declares ${uncompressedBytes} bytes (max ${maxEntryBytes})`,
        reason: "entry-too-large",
      });
    }
    if (
      exceedsCompressionRatio({
        inflatedBytes: uncompressedBytes,
        compressedBytes,
        maxRatio: compressionRatioLimitFor(entry.name, maxCompressionRatio),
      })
    ) {
      throw new DocxArchiveError({
        message: `DOCX entry "${entry.name}" declares more than ${maxCompressionRatio} bytes per compressed byte`,
        reason: "compression-ratio-exceeded",
      });
    }
    declaredTotalBytes += uncompressedBytes;
  }

  if (declaredTotalKnown && declaredTotalBytes > maxTotalBytes) {
    throw new DocxArchiveError({
      message: `DOCX archive declares ${declaredTotalBytes} cumulative bytes (max ${maxTotalBytes})`,
      reason: "total-too-large",
    });
  }
  if (
    exceedsCompressionRatio({
      inflatedBytes: declaredTotalBytes,
      compressedBytes: bytes.byteLength,
      maxRatio: maxCompressionRatio,
    })
  ) {
    throw new DocxArchiveError({
      message: `DOCX archive declares more than ${maxCompressionRatio} bytes per archive byte`,
      reason: "compression-ratio-exceeded",
    });
  }

  const xmlLimits: XmlResourceLimits = { ...FOLIO_XML_RESOURCE_LIMITS, ...options.xmlLimits };
  const xmlBudget = createXmlPackageBudget();
  const countedParts = new Set<string>();
  const inflationBudget = createInflationBudget(maxTotalBytes);
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
      const entryLimit = Math.min(requestedMaxBytes, maxEntryBytes);
      const result = await inflateEntryWithinLimits({
        entry,
        maxEntryBytes: entryLimit,
        maxCompressionRatio: compressionRatioLimitFor(path, maxCompressionRatio),
        budget: inflationBudget,
      });
      if (!result.ok) {
        throw readLimitError({
          limit: result.limit,
          path,
          maxEntryBytes: entryLimit,
          maxTotalBytes,
          maxCompressionRatio,
        });
      }
      return result.bytes;
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
        declaredUncompressedBytes: getZipEntrySizes(entry).uncompressedBytes,
      })),
    ),
    async readEntryString(path) {
      const content = await readEntry(path);
      // `ignoreBOM` keeps a leading U+FEFF in the string: OOXML parts written
      // by Word carry a UTF-8 BOM, and callers that splice a part and write it
      // back must not silently drop it.
      if (content === null) {
        return null;
      }
      const xml = new TextDecoder("utf-8", { ignoreBOM: true }).decode(content);
      const lower = path.toLowerCase();
      if (!lower.endsWith(".xml") && !lower.endsWith(".rels")) {
        return xml;
      }
      // A part re-read is not a second part: charge the package budget once per
      // path so a caller that reads `word/document.xml` twice is not refused
      // for a package it could parse once.
      const charged = countedParts.has(path);
      countedParts.add(path);
      assertXmlResourceLimits({
        xml,
        limits: xmlLimits,
        partPath: path,
        ...(charged ? {} : { budget: xmlBudget }),
      });
      return xml;
    },
    readEntryUint8: readEntry,
  };
};
