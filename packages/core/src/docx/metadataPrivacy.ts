import { TaggedError } from "better-result";
import JSZip from "jszip";

import { elementToXml, getLocalName, parseXmlDocument } from "./xmlParser";

export const FOLIO_DOCUMENT_METADATA_PROPERTIES = Object.freeze([
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

export type FolioDocumentMetadataProperty = (typeof FOLIO_DOCUMENT_METADATA_PROPERTIES)[number];

export const FOLIO_DOCUMENT_PRIVACY_TRANSFORMS = Object.freeze([
  "remove-attribution",
  "remove-timestamps",
  "remove-descriptive-metadata",
] as const);

export type FolioDocumentPrivacyTransform = (typeof FOLIO_DOCUMENT_PRIVACY_TRANSFORMS)[number];

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

export class InvalidFolioDocumentPrivacyOptionsError extends TaggedError(
  "InvalidFolioDocumentPrivacyOptionsError",
)<{
  message: string;
  receivedValue: unknown;
}>() {}

export class FolioDocumentPrivacyArchiveError extends TaggedError(
  "FolioDocumentPrivacyArchiveError",
)<{
  message: string;
  reason: "input-too-large" | "load-failed" | "too-many-entries" | "core-properties-too-large";
  cause?: unknown;
}>() {}

export const PRIVATE_METADATA_PROPERTIES_BY_TRANSFORM = {
  "remove-attribution": ["creator", "lastModifiedBy"],
  "remove-timestamps": ["created", "modified"],
  "remove-descriptive-metadata": ["title", "subject", "keywords", "description"],
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
  return FOLIO_DOCUMENT_PRIVACY_TRANSFORMS.filter((transform) => requested.has(transform));
};

const XML_DECLARATION_PATTERN = /^\s*<\?xml[^?]*\?>/u;
const MAX_INPUT_BYTES = 50 * 1024 * 1024;
const MAX_ARCHIVE_ENTRIES = 5000;
const MAX_CORE_PROPERTIES_BYTES = 1024 * 1024;

const loadPrivacyArchive = async (buffer: ArrayBuffer): Promise<JSZip> => {
  if (buffer.byteLength > MAX_INPUT_BYTES) {
    throw new FolioDocumentPrivacyArchiveError({
      message: "Document privacy input exceeded the compressed-size limit",
      reason: "input-too-large",
    });
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (cause) {
    throw new FolioDocumentPrivacyArchiveError({
      message: "Document privacy input is not a readable package",
      reason: "load-failed",
      cause,
    });
  }
  const entries = Object.values(zip.files);
  if (entries.length > MAX_ARCHIVE_ENTRIES) {
    throw new FolioDocumentPrivacyArchiveError({
      message: "Document privacy input exceeded the package-entry limit",
      reason: "too-many-entries",
    });
  }
  const coreProperties = zip.file("docProps/core.xml");
  if (!coreProperties) {
    return zip;
  }
  const data = "_data" in coreProperties ? coreProperties._data : undefined;
  const declaredBytes =
    typeof data === "object" && data !== null && "uncompressedSize" in data
      ? data.uncompressedSize
      : undefined;
  if (typeof declaredBytes !== "number" || declaredBytes > MAX_CORE_PROPERTIES_BYTES) {
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
    for (const property of PRIVATE_METADATA_PROPERTIES_BY_TRANSFORM[transform]) {
      propertiesToRemove.add(property);
    }
  }
  const removedPropertySet = new Set<FolioDocumentMetadataProperty>();
  root.elements = root.elements.filter((element) => {
    if (element.type !== "element") {
      return true;
    }
    const property = getLocalName(element.name ?? "");
    if (!FOLIO_DOCUMENT_METADATA_PROPERTIES.some((candidate) => candidate === property)) {
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
    removedMetadataProperties: FOLIO_DOCUMENT_METADATA_PROPERTIES.filter((property) =>
      removedPropertySet.has(property),
    ),
  };
};

/** Rewrite selected package metadata fields without changing other package parts. */
export const rewriteDocxMetadataPrivacy = async (
  buffer: ArrayBuffer,
  options: FolioDocumentPrivacyOptions,
): Promise<RewriteDocxMetadataPrivacyResult> => {
  const appliedTransforms = resolveFolioDocumentPrivacyTransforms(options.transforms);
  const zip = await loadPrivacyArchive(buffer);
  const coreProperties = zip.file("docProps/core.xml");
  if (!coreProperties) {
    return {
      buffer,
      privacyReport: { appliedTransforms, removedMetadataProperties: [] },
    };
  }
  const rewritten = rewriteCorePropertiesPrivacy(
    await coreProperties.async("text"),
    appliedTransforms,
  );
  if (rewritten.removedMetadataProperties.length === 0) {
    return {
      buffer,
      privacyReport: { appliedTransforms, removedMetadataProperties: [] },
    };
  }
  zip.file("docProps/core.xml", rewritten.xml);
  return {
    buffer: await zip.generateAsync({ type: "arraybuffer" }),
    privacyReport: {
      appliedTransforms,
      removedMetadataProperties: rewritten.removedMetadataProperties,
    },
  };
};
