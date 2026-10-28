import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import { ZipArchive, type ZipEntry, ZipError } from "@stll/docx-core/zip";

import { decodeText } from "../utils/textEncoding";
import { DOCX_MAX_COMPRESSION_RATIO } from "./archiveInflation";
import { elementToXml, getLocalName, parseXmlDocument } from "./xmlParser";

export const FOLIO_DOCUMENT_METADATA_PROPERTIES: readonly [
  "title",
  "subject",
  "creator",
  "keywords",
  "description",
  "lastModifiedBy",
  "revision",
  "created",
  "modified",
] = Object.freeze([
  "title",
  "subject",
  "creator",
  "keywords",
  "description",
  "lastModifiedBy",
  "revision",
  "created",
  "modified",
] as const);

export type FolioDocumentMetadataProperty =
  (typeof FOLIO_DOCUMENT_METADATA_PROPERTIES)[number];

const isFolioDocumentMetadataProperty = (
  value: string,
): value is FolioDocumentMetadataProperty =>
  FOLIO_DOCUMENT_METADATA_PROPERTIES.some((property) => property === value);

export const FOLIO_DOCUMENT_PRIVACY_TRANSFORMS: readonly [
  "remove-attribution",
  "remove-timestamps",
  "remove-descriptive-metadata",
] = Object.freeze([
  "remove-attribution",
  "remove-timestamps",
  "remove-descriptive-metadata",
] as const);

export type FolioDocumentPrivacyTransform =
  (typeof FOLIO_DOCUMENT_PRIVACY_TRANSFORMS)[number];

export const isFolioDocumentPrivacyTransform = (
  value: unknown,
): value is FolioDocumentPrivacyTransform =>
  FOLIO_DOCUMENT_PRIVACY_TRANSFORMS.some((transform) => transform === value);

export type FolioDocumentPrivacyOptions = {
  transforms: readonly FolioDocumentPrivacyTransform[];
};

export type FolioDocumentPrivacyReport = {
  appliedTransforms: FolioDocumentPrivacyTransform[];
  removedMetadataProperties: FolioDocumentMetadataProperty[];
};

export type RewriteDocxMetadataPrivacyResult = {
  buffer: ArrayBuffer;
  privacyReport: FolioDocumentPrivacyReport;
};

const InvalidFolioDocumentPrivacyOptionsErrorBase: TaggedErrorClass<"InvalidFolioDocumentPrivacyOptionsError"> =
  TaggedError("InvalidFolioDocumentPrivacyOptionsError");

export class InvalidFolioDocumentPrivacyOptionsError extends InvalidFolioDocumentPrivacyOptionsErrorBase<{
  message: string;
  receivedValue: unknown;
}> {}

const FolioDocumentPrivacyArchiveErrorBase: TaggedErrorClass<"FolioDocumentPrivacyArchiveError"> =
  TaggedError("FolioDocumentPrivacyArchiveError");

export class FolioDocumentPrivacyArchiveError extends FolioDocumentPrivacyArchiveErrorBase<{
  message: string;
  reason:
    | "input-too-large"
    | "load-failed"
    | "too-many-entries"
    | "core-properties-too-large";
  cause?: unknown;
}> {}

export const PRIVATE_METADATA_PROPERTIES_BY_TRANSFORM: {
  readonly "remove-attribution": readonly ["creator", "lastModifiedBy"];
  readonly "remove-timestamps": readonly ["created", "modified"];
  readonly "remove-descriptive-metadata": readonly [
    "title",
    "subject",
    "keywords",
    "description",
  ];
} = {
  "remove-attribution": ["creator", "lastModifiedBy"],
  "remove-timestamps": ["created", "modified"],
  "remove-descriptive-metadata": [
    "title",
    "subject",
    "keywords",
    "description",
  ],
} as const satisfies Record<
  FolioDocumentPrivacyTransform,
  readonly FolioDocumentMetadataProperty[]
>;

export const resolveFolioDocumentPrivacyTransforms = (
  transforms: unknown,
): FolioDocumentPrivacyTransform[] => {
  if (
    !Array.isArray(transforms) ||
    transforms.some((transform) => !isFolioDocumentPrivacyTransform(transform))
  ) {
    throw new InvalidFolioDocumentPrivacyOptionsError({
      message: "Document privacy received an unrecognized transform",
      receivedValue: transforms,
    });
  }
  const requested = new Set(transforms);
  return FOLIO_DOCUMENT_PRIVACY_TRANSFORMS.filter((transform) =>
    requested.has(transform),
  );
};

const XML_DECLARATION_PATTERN = /^\s*<\?xml[^?]*\?>/u;
const MAX_INPUT_BYTES = 50 * 1024 * 1024;
const MAX_ARCHIVE_ENTRIES = 5000;
const MAX_CORE_PROPERTIES_BYTES = 1024 * 1024;
const CORE_PROPERTIES_PATH_LOWER = "docprops/core.xml";

/**
 * Look up a package part by its conventional lowercase path, falling back to
 * a case-insensitive scan. OPC part names are case-insensitive, so a
 * producer that wrote `docProps/Core.xml` (or another casing) must still be
 * found — otherwise the privacy rewrite silently no-ops and the unscrubbed
 * buffer round-trips untouched.
 */
const findZipEntryCaseInsensitive = (
  zip: ZipArchive,
  lowerPath: string,
): ZipEntry | null =>
  zip.entry(lowerPath) ??
  zip.entries().find(({ path }) => path.toLowerCase() === lowerPath) ??
  null;

const loadPrivacyArchive = async (buffer: ArrayBuffer): Promise<ZipArchive> => {
  if (buffer.byteLength > MAX_INPUT_BYTES) {
    throw new FolioDocumentPrivacyArchiveError({
      message: "Document privacy input exceeded the compressed-size limit",
      reason: "input-too-large",
    });
  }
  const read = await Result.tryPromise({
    try: async () =>
      await ZipArchive.read(buffer, {
        limits: {
          maxEntries: MAX_ARCHIVE_ENTRIES,
          maxCompressionRatio: DOCX_MAX_COMPRESSION_RATIO,
        },
      }),
    catch: (cause) => cause,
  });
  if (Result.isError(read)) {
    const cause = read.error;
    if (cause instanceof ZipError && cause.reason === "too-many-entries") {
      throw new FolioDocumentPrivacyArchiveError({
        message: "Document privacy input exceeded the package-entry limit",
        reason: "too-many-entries",
      });
    }
    throw new FolioDocumentPrivacyArchiveError({
      message: "Document privacy input is not a readable package",
      reason: "load-failed",
      cause,
    });
  }
  const zip = read.value;
  const coreProperties = findZipEntryCaseInsensitive(
    zip,
    CORE_PROPERTIES_PATH_LOWER,
  );
  if (coreProperties && coreProperties.size > MAX_CORE_PROPERTIES_BYTES) {
    throw new FolioDocumentPrivacyArchiveError({
      message: "Document privacy core properties exceeded the part-size limit",
      reason: "core-properties-too-large",
    });
  }
  return zip;
};

type RewriteCorePropertiesPrivacyResult = {
  xml: string;
  removedMetadataProperties: FolioDocumentMetadataProperty[];
};

const rewriteCorePropertiesPrivacy = (
  xml: string,
  transforms: readonly FolioDocumentPrivacyTransform[],
): RewriteCorePropertiesPrivacyResult => {
  const root = parseXmlDocument(xml);
  if (!root?.elements) {
    return { xml, removedMetadataProperties: [] };
  }
  const propertiesToRemove = new Set<FolioDocumentMetadataProperty>();
  for (const transform of transforms) {
    for (const property of PRIVATE_METADATA_PROPERTIES_BY_TRANSFORM[
      transform
    ]) {
      propertiesToRemove.add(property);
    }
  }
  const removedPropertySet = new Set<FolioDocumentMetadataProperty>();
  root.elements = root.elements.filter((element) => {
    if (element.type !== "element") {
      return true;
    }
    const property = getLocalName(element.name ?? "");
    if (!isFolioDocumentMetadataProperty(property)) {
      return true;
    }
    if (!propertiesToRemove.has(property)) {
      return true;
    }
    removedPropertySet.add(property);
    return false;
  });
  const declaration = xml.match(XML_DECLARATION_PATTERN)?.at(0) ?? "";
  return {
    xml: `${declaration}${elementToXml(root)}`,
    removedMetadataProperties: FOLIO_DOCUMENT_METADATA_PROPERTIES.filter(
      (property) => removedPropertySet.has(property),
    ),
  };
};

/**
 * Read the core-properties part. Its declared size was checked when the
 * archive loaded, and inflation stops at that size.
 */
const readCoreProperties = async (entry: ZipEntry): Promise<string> => {
  const read = await Result.tryPromise({
    try: async () => await entry.bytes(),
    catch: (cause) => cause,
  });
  if (Result.isError(read)) {
    const cause = read.error;
    throw new FolioDocumentPrivacyArchiveError(
      cause instanceof ZipError && cause.reason === "declared-size-exceeded"
        ? {
            message:
              "Document privacy core properties exceeded the part-size limit",
            reason: "core-properties-too-large",
          }
        : {
            message: "Document privacy input is not a readable package",
            reason: "load-failed",
            cause,
          },
    );
  }
  // `ignoreBOM` keeps a leading U+FEFF, as the previous string read did.
  return decodeText(read.value, { ignoreBOM: true });
};

/** Rewrite selected package metadata fields without changing other package parts. */
export const rewriteDocxMetadataPrivacy = async (
  buffer: ArrayBuffer,
  { transforms }: FolioDocumentPrivacyOptions,
): Promise<RewriteDocxMetadataPrivacyResult> => {
  const appliedTransforms = resolveFolioDocumentPrivacyTransforms(transforms);
  const zip = await loadPrivacyArchive(buffer);
  const coreProperties = findZipEntryCaseInsensitive(
    zip,
    CORE_PROPERTIES_PATH_LOWER,
  );
  if (!coreProperties) {
    return {
      buffer,
      privacyReport: { appliedTransforms, removedMetadataProperties: [] },
    };
  }
  const rewritten = rewriteCorePropertiesPrivacy(
    await readCoreProperties(coreProperties),
    appliedTransforms,
  );
  if (rewritten.removedMetadataProperties.length === 0) {
    return {
      buffer,
      privacyReport: { appliedTransforms, removedMetadataProperties: [] },
    };
  }
  // Write back under the entry's original name/casing so a non-conventional
  // producer doesn't end up with two core-properties-shaped parts.
  zip.set(coreProperties.path, rewritten.xml);
  return {
    buffer: (await zip.generate()).buffer,
    privacyReport: {
      appliedTransforms,
      removedMetadataProperties: rewritten.removedMetadataProperties,
    },
  };
};
