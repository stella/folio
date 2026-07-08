/**
 * Decrypt password-protected OOXML before ZIP extraction.
 */

import { decryptAgilePackage } from "./agileDecryptor";
import { DOCX_CONTAINER_TYPES, detectDocxContainerType } from "./detectContainer";
import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";
import { extractEncryptionStreams } from "./oleReader";
import { parseEncryptionInfo } from "./parseEncryptionInfo";

export type DecryptDocxOptions = {
  /** Password for Agile-encrypted .docx files (Office 2010+). */
  password?: string | undefined;
};

export type DecryptDocxResult = {
  /** Plaintext OOXML ZIP bytes, or the original buffer when unencrypted. */
  data: ArrayBuffer;
  /** True when the input was an encrypted CFB container and decryption ran. */
  wasEncrypted: boolean;
};

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

const validateDecryptedZip = (data: Uint8Array): void => {
  if (detectDocxContainerType(data) !== DOCX_CONTAINER_TYPES.ZIP) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "Decrypted output is not a valid ZIP archive — the file may be corrupt",
    });
  }
};

/**
 * Decrypt a password-protected `.docx` when needed, or return the input unchanged.
 *
 * Supports Agile Encryption (Office 2010+). Standard Encryption and legacy RC4
 * are rejected with `DOCX_ENCRYPTION_UNSUPPORTED`.
 *
 * Save produces a standard unencrypted OOXML ZIP; folio does not re-encrypt on export.
 */
export const decryptDocxIfNeeded = async (
  data: ArrayBuffer | Uint8Array,
  options: DecryptDocxOptions = {},
): Promise<DecryptDocxResult> => {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const containerType = detectDocxContainerType(bytes);

  if (containerType === DOCX_CONTAINER_TYPES.ZIP) {
    return { data: toArrayBuffer(bytes), wasEncrypted: false };
  }

  if (containerType !== DOCX_CONTAINER_TYPES.CFB) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "Unrecognized file format — expected a .docx (ZIP) or encrypted .docx (OLE/CFB)",
    });
  }

  if (!options.password) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.PASSWORD_REQUIRED,
      message: "This document is password-protected. A password is required to open it.",
    });
  }

  const streams = extractEncryptionStreams(bytes);
  const { params } = parseEncryptionInfo(streams.encryptionInfo);
  const decryptedZip = await decryptAgilePackage(options.password, params, streams.encryptedPackage);
  validateDecryptedZip(decryptedZip);

  return { data: toArrayBuffer(decryptedZip), wasEncrypted: true };
};

/**
 * Open a `.docx` buffer for parsing: decrypt when encrypted, then return the ZIP bytes.
 */
export const openDocxBuffer = async (
  data: ArrayBuffer | Uint8Array,
  options: DecryptDocxOptions = {},
): Promise<ArrayBuffer> => (await decryptDocxIfNeeded(data, options)).data;
