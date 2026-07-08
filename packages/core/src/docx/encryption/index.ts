export {
  DOCX_ENCRYPTION_ERROR_CODES,
  DocxEncryptionError,
  type DocxEncryptionErrorCode,
  isDocxEncryptionError,
} from "./errors";
export {
  DOCX_CONTAINER_TYPES,
  detectDocxContainerType,
  type DocxContainerType,
} from "./containerFormat";
export {
  decryptDocxIfNeeded,
  openDocxBuffer,
  type DecryptDocxOptions,
  type DecryptDocxResult,
} from "./openEncryptedDocx";
