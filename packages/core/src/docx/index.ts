// Barrel for the DOCX parse layer. Adapters import numbering helpers from
// `@stll/folio-core/docx`; the parser, serializer, and other modules stay
// reachable at their explicit subpaths via the core `"./*"` export wildcard.
export { getCachedNumberingMap } from "./numberingParser";
export {
  FOLIO_DOCUMENT_METADATA_PROPERTIES,
  FOLIO_DOCUMENT_PRIVACY_TRANSFORMS,
  FolioDocumentPrivacyArchiveError,
  InvalidFolioDocumentPrivacyOptionsError,
  isFolioDocumentPrivacyTransform,
  rewriteDocxMetadataPrivacy,
  type FolioDocumentMetadataProperty,
  type FolioDocumentPrivacyOptions,
  type FolioDocumentPrivacyReport,
  type FolioDocumentPrivacyTransform,
  type RewriteDocxMetadataPrivacyResult,
} from "./metadataPrivacy";
export {
  DOCX_ENCRYPTION_ERROR_CODES,
  DocxEncryptionError,
  decryptDocxIfNeeded,
  isDocxEncryptionError,
  openDocxBuffer,
  type DecryptDocxOptions,
  type DecryptDocxResult,
  type DocxEncryptionErrorCode,
} from "./encryption";
