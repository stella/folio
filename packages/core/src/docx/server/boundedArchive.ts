import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import {
  assertXmlResourceLimits,
  createXmlPackageBudget,
  FOLIO_XML_RESOURCE_LIMITS,
  type XmlResourceLimits,
} from "@stll/docx-core/xml/limits";
import { ZipArchive, ZipError } from "@stll/docx-core/zip";

import { decodeText } from "../../utils/textEncoding";
import {
  DOCX_COMPRESSION_RATIO_GRACE_BYTES,
  DOCX_MAX_COMPRESSION_RATIO,
  isRatioBoundedPart,
} from "../archiveInflation";

export const DOCX_MAX_ENTRY_BYTES: number = 128 * 1024 * 1024;
export const DOCX_MAX_TOTAL_BYTES: number = 256 * 1024 * 1024;
export const DOCX_MAX_ENTRIES = 4096;
export const DOCX_MAX_INPUT_BYTES: number = 50 * 1024 * 1024;

const DocxArchiveErrorBase: TaggedErrorClass<"DocxArchiveError"> =
  TaggedError("DocxArchiveError");

/** Error raised when a DOCX archive cannot be loaded within configured limits. */
export class DocxArchiveError extends DocxArchiveErrorBase<{
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
  readonly declaredUncompressedBytes: number;
};

export type DocxArchiveReadOptions = {
  maxBytes?: number;
};

export type DocxArchive = {
  entries: readonly string[];
  entryMetadata: readonly DocxArchiveEntry[];
  readEntryString: (path: string) => Promise<string | null>;
  readEntryUint8: (
    path: string,
    options?: DocxArchiveReadOptions,
  ) => Promise<Uint8Array | null>;
};

type ArchiveLimits = {
  maxEntryBytes: number;
  maxTotalBytes: number;
  maxEntries: number;
  maxCompressionRatio: number;
};

/** The archive error for a read the ZIP reader refused. */
const archiveErrorFor = (
  error: ZipError,
  {
    maxEntryBytes,
    maxTotalBytes,
    maxEntries,
    maxCompressionRatio,
  }: ArchiveLimits,
): DocxArchiveError => {
  const entry = error.path === undefined ? "" : ` "${error.path}"`;
  switch (error.reason) {
    case "too-many-entries":
      return new DocxArchiveError({
        message: `DOCX archive holds more than ${maxEntries} entries`,
        reason: "too-many-entries",
      });
    case "entry-too-large":
      return new DocxArchiveError({
        message: `DOCX entry${entry} declares more than ${maxEntryBytes} bytes`,
        reason: "entry-too-large",
      });
    case "declared-size-exceeded":
      return new DocxArchiveError({
        message: `DOCX entry${entry} inflated past its declared size`,
        reason: "entry-too-large",
      });
    case "total-too-large":
      return new DocxArchiveError({
        message: `DOCX archive declares more than ${maxTotalBytes} cumulative bytes`,
        reason: "total-too-large",
      });
    case "compression-ratio":
      return new DocxArchiveError({
        message:
          error.path === undefined
            ? `DOCX archive declares more than ${maxCompressionRatio} bytes per archive byte`
            : `DOCX entry${entry} declares more than ${maxCompressionRatio} bytes per compressed byte`,
        reason: "compression-ratio-exceeded",
      });
    case "malformed":
    case "unsupported":
    case "invalid-entry-name":
    case "duplicate-entry":
    case "corrupt-entry":
      return new DocxArchiveError({
        message: "Failed to parse DOCX archive",
        reason: "load-failed",
        cause: error,
      });
  }
};

const toArchiveError = (cause: unknown, limits: ArchiveLimits): unknown =>
  cause instanceof ZipError ? archiveErrorFor(cause, limits) : cause;

type ResolveByteLimitOptions = {
  value: number | undefined;
  fallback: number;
  name: string;
};

const resolveByteLimit = ({
  value,
  fallback,
  name,
}: ResolveByteLimitOptions): number => {
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

  const limits: ArchiveLimits = {
    maxEntryBytes,
    maxTotalBytes,
    maxEntries,
    maxCompressionRatio,
  };
  const read = await Result.tryPromise({
    try: async () =>
      await ZipArchive.read(bytes, {
        limits: {
          maxEntries,
          maxEntryBytes,
          maxTotalBytes,
          maxCompressionRatio,
          compressionRatioGraceBytes: DOCX_COMPRESSION_RATIO_GRACE_BYTES,
        },
        // Binary parts can compress like a run of one byte; the byte limits
        // and the archive-wide ratio bound them.
        isCompressionRatioExempt: (path) => !isRatioBoundedPart(path),
      }),
    catch: (cause) => toArchiveError(cause, limits),
  });
  if (Result.isError(read)) {
    throw read.error;
  }
  const zip = read.value;
  const archiveEntries = zip.entries();

  const xmlLimits: XmlResourceLimits = {
    ...FOLIO_XML_RESOURCE_LIMITS,
    ...options.xmlLimits,
  };
  const xmlBudget = createXmlPackageBudget();
  const countedParts = new Set<string>();
  // Every read inflates at most its entry's declared size, and the declared
  // sizes fit `maxTotalBytes`; the running total bounds repeated reads too.
  let readBytes = 0;

  const readEntry = async (
    path: string,
    readOptions: DocxArchiveReadOptions = {},
  ): Promise<Uint8Array | null> => {
    const requestedMaxBytes = resolveByteLimit({
      value: readOptions.maxBytes,
      fallback: maxEntryBytes,
      name: "DOCX entry read byte limit",
    });
    const entry = zip.entry(path);
    if (!entry) {
      return null;
    }
    const entryLimit = Math.min(requestedMaxBytes, maxEntryBytes);
    if (entry.size > entryLimit) {
      throw new DocxArchiveError({
        message: `DOCX entry "${path}" exceeded the ${entryLimit}-byte limit`,
        reason: "entry-too-large",
      });
    }
    if (readBytes + entry.size > maxTotalBytes) {
      throw new DocxArchiveError({
        message: `DOCX archive exceeded the ${maxTotalBytes}-byte cumulative limit while reading "${path}"`,
        reason: "total-too-large",
      });
    }
    readBytes += entry.size;
    const content = await Result.tryPromise({
      try: async () => await entry.bytes(),
      catch: (cause) => toArchiveError(cause, limits),
    });
    if (Result.isError(content)) {
      readBytes -= entry.size;
      throw content.error;
    }
    return content.value;
  };

  return {
    entries: Object.freeze(archiveEntries.map(({ path }) => path)),
    entryMetadata: Object.freeze(
      archiveEntries.map((entry) => ({
        path: entry.path,
        declaredUncompressedBytes: entry.size,
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
      const xml = decodeText(content, { ignoreBOM: true });
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
