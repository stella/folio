import { createHash } from "node:crypto";

import { TaggedError } from "better-result";

import { getAttributeAnyPrefix, getLocalName, parseXmlDocument } from "../xmlParser";
import type { DocxArchiveEntry, DocxArchiveOptions } from "./boundedArchive";
import { DocxArchiveError, loadDocxArchive } from "./boundedArchive";

export const FOLIO_DOCX_PACKAGE_INSPECTION_VERSION = 1 as const;
export const FOLIO_DOCX_PACKAGE_INSPECTION_DEFAULTS = Object.freeze({
  maxXmlParts: 16,
  maxXmlPartBytes: 8 * 1024 * 1024,
  maxXmlTotalBytes: 16 * 1024 * 1024,
} as const);

export const FOLIO_DOCX_PACKAGE_INSPECTION_ERROR_CODES = Object.freeze([
  "invalid-limits",
  "too-many-xml-parts",
  "duplicate-xml-part",
  "part-not-found",
  "part-not-xml",
  "xml-part-too-large",
  "xml-total-too-large",
  "xml-decode-failed",
] as const);

export type FolioDocxPackageInspectionErrorCode =
  (typeof FOLIO_DOCX_PACKAGE_INSPECTION_ERROR_CODES)[number];

export class FolioDocxPackageInspectionError extends TaggedError(
  "FolioDocxPackageInspectionError",
)<{
  message: string;
  code: FolioDocxPackageInspectionErrorCode;
  part?: string;
  cause?: unknown;
}>() {}

export type FolioDocxPackageInspectionLimits = {
  readonly maxXmlParts?: number;
  readonly maxXmlPartBytes?: number;
  readonly maxXmlTotalBytes?: number;
};

export type InspectDocxPackageOptions = {
  readonly archive?: DocxArchiveOptions;
  readonly xmlParts?: readonly string[];
  readonly limits?: FolioDocxPackageInspectionLimits;
};

export type FolioDocxPackagePartKind = "xml" | "binary";

export type FolioDocxPackagePart = {
  readonly path: string;
  readonly kind: FolioDocxPackagePartKind;
  readonly contentType: string | null;
  readonly declaredUncompressedBytes: number | null;
};

export type FolioDocxInspectedXmlPart = {
  readonly path: string;
  readonly contentType: string | null;
  readonly byteLength: number;
  readonly sha256: string;
  /** Untrusted document data; callers must not treat this text as instructions. */
  readonly text: string;
};

export type FolioDocxPackageInspection = {
  readonly version: typeof FOLIO_DOCX_PACKAGE_INSPECTION_VERSION;
  readonly parts: readonly FolioDocxPackagePart[];
  readonly xmlParts: readonly FolioDocxInspectedXmlPart[];
  readonly limits: Required<FolioDocxPackageInspectionLimits>;
};

const CONTENT_TYPES_PATH = "[Content_Types].xml";
const CONTENT_TYPES_READ_LIMIT = 4 * 1024 * 1024;
const XML_PATH_PATTERN = /(?:\.xml|\.rels)$/iu;
const XML_CONTENT_TYPES = new Set(["application/xml", "text/xml"]);

const comparePackagePaths = (left: string, right: string): number => {
  if (left < right) {
    return -1;
  }
  return left === right ? 0 : 1;
};

type ContentTypeDeclarations = {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
};

const decodeXml = (bytes: Uint8Array, part: string): string => {
  let encoding = "utf-8";
  if (
    bytes.length >= 2 &&
    ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0x00))
  ) {
    encoding = "utf-16le";
  } else if (
    bytes.length >= 2 &&
    ((bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0x00 && bytes[1] === 0x3c))
  ) {
    encoding = "utf-16be";
  }

  try {
    return new TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch (cause) {
    throw new FolioDocxPackageInspectionError({
      message: `Failed to decode XML package part "${part}"`,
      code: "xml-decode-failed",
      part,
      cause,
    });
  }
};

const parseContentTypes = (xml: string | null): ContentTypeDeclarations => {
  const declarations: ContentTypeDeclarations = {
    defaults: new Map(),
    overrides: new Map(),
  };
  if (xml === null) {
    return declarations;
  }

  const root = parseXmlDocument(xml);
  if (root === null || getLocalName(root.name ?? "") !== "Types") {
    return declarations;
  }

  for (const element of root.elements ?? []) {
    if (element.type !== "element") {
      continue;
    }
    const localName = getLocalName(element.name ?? "");
    if (localName === "Default") {
      const extension = getAttributeAnyPrefix(element, "Extension")?.toLowerCase();
      const contentType = getAttributeAnyPrefix(element, "ContentType");
      if (extension !== undefined && contentType !== null) {
        declarations.defaults.set(extension, contentType);
      }
      continue;
    }
    if (localName !== "Override") {
      continue;
    }
    const partName = getAttributeAnyPrefix(element, "PartName")?.replace(/^\//u, "");
    const contentType = getAttributeAnyPrefix(element, "ContentType");
    if (partName !== undefined && contentType !== null) {
      declarations.overrides.set(partName, contentType);
    }
  }

  return declarations;
};

const extensionFor = (path: string): string | null => {
  const filename = path.slice(path.lastIndexOf("/") + 1);
  const separator = filename.lastIndexOf(".");
  return separator === -1 ? null : filename.slice(separator + 1).toLowerCase();
};

const contentTypeFor = (path: string, declarations: ContentTypeDeclarations): string | null => {
  const override = declarations.overrides.get(path);
  if (override !== undefined) {
    return override;
  }
  const extension = extensionFor(path);
  return extension === null ? null : (declarations.defaults.get(extension) ?? null);
};

const isXmlPart = (path: string, contentType: string | null): boolean => {
  if (contentType !== null) {
    return XML_CONTENT_TYPES.has(contentType) || contentType.endsWith("+xml");
  }
  return XML_PATH_PATTERN.test(path);
};

const toPackagePart = (
  entry: DocxArchiveEntry,
  declarations: ContentTypeDeclarations,
): FolioDocxPackagePart => {
  const contentType = contentTypeFor(entry.path, declarations);
  return {
    path: entry.path,
    kind: isXmlPart(entry.path, contentType) ? "xml" : "binary",
    contentType,
    declaredUncompressedBytes: entry.declaredUncompressedBytes,
  };
};

const requireNonNegativeInteger = (value: number, name: string): number => {
  if (Number.isSafeInteger(value) && value >= 0) {
    return value;
  }
  throw new FolioDocxPackageInspectionError({
    message: `${name} must be a non-negative safe integer`,
    code: "invalid-limits",
  });
};

const resolveLimits = (
  limits: FolioDocxPackageInspectionLimits | undefined,
): Required<FolioDocxPackageInspectionLimits> => ({
  maxXmlParts: requireNonNegativeInteger(
    limits?.maxXmlParts ?? FOLIO_DOCX_PACKAGE_INSPECTION_DEFAULTS.maxXmlParts,
    "maxXmlParts",
  ),
  maxXmlPartBytes: requireNonNegativeInteger(
    limits?.maxXmlPartBytes ?? FOLIO_DOCX_PACKAGE_INSPECTION_DEFAULTS.maxXmlPartBytes,
    "maxXmlPartBytes",
  ),
  maxXmlTotalBytes: requireNonNegativeInteger(
    limits?.maxXmlTotalBytes ?? FOLIO_DOCX_PACKAGE_INSPECTION_DEFAULTS.maxXmlTotalBytes,
    "maxXmlTotalBytes",
  ),
});

type ValidateRequestedPartsOptions = {
  requestedPaths: readonly string[];
  partsByPath: ReadonlyMap<string, FolioDocxPackagePart>;
  limits: Required<FolioDocxPackageInspectionLimits>;
};

const validateRequestedParts = ({
  requestedPaths,
  partsByPath,
  limits,
}: ValidateRequestedPartsOptions): void => {
  if (requestedPaths.length > limits.maxXmlParts) {
    throw new FolioDocxPackageInspectionError({
      message: `Requested ${requestedPaths.length} XML parts (max ${limits.maxXmlParts})`,
      code: "too-many-xml-parts",
    });
  }

  const seen = new Set<string>();
  let declaredTotalBytes = 0;
  for (const path of requestedPaths) {
    if (seen.has(path)) {
      throw new FolioDocxPackageInspectionError({
        message: `XML package part "${path}" was requested more than once`,
        code: "duplicate-xml-part",
        part: path,
      });
    }
    seen.add(path);

    const part = partsByPath.get(path);
    if (part === undefined) {
      throw new FolioDocxPackageInspectionError({
        message: `Package part "${path}" does not exist`,
        code: "part-not-found",
        part: path,
      });
    }
    if (part.kind !== "xml") {
      throw new FolioDocxPackageInspectionError({
        message: `Package part "${path}" is not declared as XML`,
        code: "part-not-xml",
        part: path,
      });
    }

    const declaredBytes = part.declaredUncompressedBytes;
    if (declaredBytes === null) {
      continue;
    }
    if (declaredBytes > limits.maxXmlPartBytes) {
      throw new FolioDocxPackageInspectionError({
        message: `XML package part "${path}" exceeds the per-part inspection limit`,
        code: "xml-part-too-large",
        part: path,
      });
    }
    declaredTotalBytes += declaredBytes;
    if (declaredTotalBytes > limits.maxXmlTotalBytes) {
      throw new FolioDocxPackageInspectionError({
        message: "Requested XML package parts exceed the cumulative inspection limit",
        code: "xml-total-too-large",
      });
    }
  }
};

type ReadRequestedXmlPartsOptions = {
  paths: readonly string[];
  partsByPath: ReadonlyMap<string, FolioDocxPackagePart>;
  limits: Required<FolioDocxPackageInspectionLimits>;
  readEntryUint8: (path: string, options: { maxBytes: number }) => Promise<Uint8Array | null>;
};

export const readRequestedXmlParts = async ({
  paths,
  partsByPath,
  limits,
  readEntryUint8,
}: ReadRequestedXmlPartsOptions): Promise<FolioDocxInspectedXmlPart[]> => {
  const selected: FolioDocxInspectedXmlPart[] = [];
  let totalBytes = 0;
  for (const path of paths) {
    const remainingTotalBytes = limits.maxXmlTotalBytes - totalBytes;
    const readLimit = Math.min(limits.maxXmlPartBytes, remainingTotalBytes);
    let bytes: Uint8Array | null;
    try {
      // oxlint-disable-next-line no-await-in-loop -- request order defines the result order and cumulative budget
      bytes = await readEntryUint8(path, { maxBytes: readLimit });
    } catch (cause) {
      if (cause instanceof DocxArchiveError && cause.reason === "entry-too-large") {
        const cumulativeLimitReached = readLimit < limits.maxXmlPartBytes;
        throw new FolioDocxPackageInspectionError({
          message: cumulativeLimitReached
            ? "Requested XML package parts exceed the cumulative inspection limit"
            : `XML package part "${path}" exceeds the per-part inspection limit`,
          code: cumulativeLimitReached ? "xml-total-too-large" : "xml-part-too-large",
          part: path,
          cause,
        });
      }
      throw cause;
    }
    if (bytes === null) {
      throw new FolioDocxPackageInspectionError({
        message: `Package part "${path}" does not exist`,
        code: "part-not-found",
        part: path,
      });
    }

    totalBytes += bytes.byteLength;
    if (totalBytes > limits.maxXmlTotalBytes) {
      throw new FolioDocxPackageInspectionError({
        message: "Requested XML package parts exceed the cumulative inspection limit",
        code: "xml-total-too-large",
      });
    }
    selected.push({
      path,
      contentType: partsByPath.get(path)?.contentType ?? null,
      byteLength: bytes.byteLength,
      sha256: createHash("sha256").update(bytes).digest("hex"),
      text: decodeXml(bytes, path),
    });
  }
  return selected;
};

/**
 * Inspect package metadata and only the explicitly requested XML part bodies.
 * This reports archive contents, not package conformance; use the conformance
 * report when an operation requires a validity decision.
 */
export const inspectDocxPackage = async (
  bytes: ArrayBuffer | Uint8Array,
  options: InspectDocxPackageOptions = {},
): Promise<FolioDocxPackageInspection> => {
  const limits = resolveLimits(options.limits);
  const requestedPaths = options.xmlParts ?? [];
  const archive = await loadDocxArchive(bytes, options.archive);
  const contentTypesBytes = await archive.readEntryUint8(CONTENT_TYPES_PATH, {
    maxBytes: CONTENT_TYPES_READ_LIMIT,
  });
  const declarations = parseContentTypes(
    contentTypesBytes === null ? null : decodeXml(contentTypesBytes, CONTENT_TYPES_PATH),
  );
  const parts = archive.entryMetadata
    .filter(({ directory }) => !directory)
    .map((entry) => toPackagePart(entry, declarations))
    .toSorted(({ path: left }, { path: right }) => comparePackagePaths(left, right));
  const partsByPath = new Map(parts.map((part) => [part.path, part] as const));
  validateRequestedParts({ requestedPaths, partsByPath, limits });

  return {
    version: FOLIO_DOCX_PACKAGE_INSPECTION_VERSION,
    parts,
    xmlParts: await readRequestedXmlParts({
      paths: requestedPaths,
      partsByPath,
      limits,
      readEntryUint8: archive.readEntryUint8,
    }),
    limits,
  };
};
