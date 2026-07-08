import { TaggedError } from "better-result";

/** Machine-readable codes for password-protected OOXML failures. */
export const DOCX_ENCRYPTION_ERROR_CODES = {
  PASSWORD_REQUIRED: "DOCX_PASSWORD_REQUIRED",
  PASSWORD_INVALID: "DOCX_PASSWORD_INVALID",
  ENCRYPTION_UNSUPPORTED: "DOCX_ENCRYPTION_UNSUPPORTED",
  DECRYPTION_FAILED: "DOCX_DECRYPTION_FAILED",
} as const;

export type DocxEncryptionErrorCode =
  (typeof DOCX_ENCRYPTION_ERROR_CODES)[keyof typeof DOCX_ENCRYPTION_ERROR_CODES];

export class DocxEncryptionError extends TaggedError("DocxEncryptionError")<{
  message: string;
  code: DocxEncryptionErrorCode;
  cause?: unknown;
}>() {}

export const isDocxEncryptionError = (error: unknown): error is DocxEncryptionError =>
  error instanceof DocxEncryptionError;
