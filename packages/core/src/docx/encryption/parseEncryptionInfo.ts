/**
 * Parse an OOXML EncryptionInfo stream (Agile Encryption, Office 2010+).
 *
 * @see https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
 */

import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";

export type AgileEncryptionParams = {
  keySalt: Uint8Array;
  keyBits: number;
  hashAlgorithm: string;
  hashSize: number;
  blockSize: number;
  cipherAlgorithm: string;
  cipherChaining: string;
  spinCount: number;
  passwordSalt: Uint8Array;
  passwordHashAlgorithm: string;
  passwordKeyBits: number;
  passwordBlockSize: number;
  encryptedVerifierHashInput: Uint8Array;
  encryptedVerifierHashValue: Uint8Array;
  encryptedKeyValue: Uint8Array;
  encryptedHmacKey: Uint8Array;
  encryptedHmacValue: Uint8Array;
};

export type ParsedEncryptionInfo = {
  type: "agile";
  params: AgileEncryptionParams;
};

const AGILE_VERSION = 0x0004;
const AGILE_RESERVED = 0x0004;
const STANDARD_RESERVED = 0x0003;
const MIN_HEADER_SIZE = 8;

const readUint16LE = (data: Uint8Array, offset: number): number =>
  data[offset]! | (data[offset + 1]! << 8);

const extractElementAttrs = (xml: string, localName: string): Record<string, string> | null => {
  const tagPattern = new RegExp(`<(?:\\w+:)?${localName}\\b([^>]*)/?>`);
  const match = xml.match(tagPattern);
  if (!match?.[1]) {
    return null;
  }

  const attrs: Record<string, string> = {};
  const attrPattern = /(\w+)="([^"]*)"/g;
  let attrMatch: RegExpExecArray | null;
  while ((attrMatch = attrPattern.exec(match[1])) !== null) {
    attrs[attrMatch[1]!] = attrMatch[2]!;
  }
  return attrs;
};

const base64ToBytes = (b64: string): Uint8Array => {
  if (typeof atob === "function") {
    const raw = atob(b64);
    const bytes = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i++) {
      bytes[i] = raw.charCodeAt(i);
    }
    return bytes;
  }
  return new Uint8Array(Buffer.from(b64, "base64"));
};

const requireAttr = (attrs: Record<string, string> | null, name: string, label: string): string => {
  const value = attrs?.[name];
  if (value === undefined) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `Missing required attribute "${name}" on <${label}>`,
    });
  }
  return value;
};

const requireBase64Attr = (
  attrs: Record<string, string> | null,
  name: string,
  label: string,
): Uint8Array => base64ToBytes(requireAttr(attrs, name, label));

const requireIntAttr = (attrs: Record<string, string> | null, name: string, label: string): number => {
  const value = Number.parseInt(requireAttr(attrs, name, label), 10);
  if (Number.isNaN(value)) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `Invalid integer for attribute "${name}" on <${label}>`,
    });
  }
  return value;
};

const parseAgileXml = (xmlBytes: Uint8Array): AgileEncryptionParams => {
  const xml = new TextDecoder("utf-8").decode(xmlBytes);
  const keyData = extractElementAttrs(xml, "keyData");
  const dataIntegrity = extractElementAttrs(xml, "dataIntegrity");
  const encryptedKey = extractElementAttrs(xml, "encryptedKey");

  return {
    keySalt: requireBase64Attr(keyData, "saltValue", "keyData"),
    keyBits: requireIntAttr(keyData, "keyBits", "keyData"),
    hashAlgorithm: requireAttr(keyData, "hashAlgorithm", "keyData"),
    hashSize: requireIntAttr(keyData, "hashSize", "keyData"),
    blockSize: requireIntAttr(keyData, "blockSize", "keyData"),
    cipherAlgorithm: requireAttr(keyData, "cipherAlgorithm", "keyData"),
    cipherChaining: requireAttr(keyData, "cipherChaining", "keyData"),
    spinCount: requireIntAttr(encryptedKey, "spinCount", "encryptedKey"),
    passwordSalt: requireBase64Attr(encryptedKey, "saltValue", "encryptedKey"),
    passwordHashAlgorithm: requireAttr(encryptedKey, "hashAlgorithm", "encryptedKey"),
    passwordKeyBits: requireIntAttr(encryptedKey, "keyBits", "encryptedKey"),
    passwordBlockSize: requireIntAttr(encryptedKey, "blockSize", "encryptedKey"),
    encryptedVerifierHashInput: requireBase64Attr(
      encryptedKey,
      "encryptedVerifierHashInput",
      "encryptedKey",
    ),
    encryptedVerifierHashValue: requireBase64Attr(
      encryptedKey,
      "encryptedVerifierHashValue",
      "encryptedKey",
    ),
    encryptedKeyValue: requireBase64Attr(encryptedKey, "encryptedKeyValue", "encryptedKey"),
    encryptedHmacKey: requireBase64Attr(dataIntegrity, "encryptedHmacKey", "dataIntegrity"),
    encryptedHmacValue: requireBase64Attr(dataIntegrity, "encryptedHmacValue", "dataIntegrity"),
  };
};

export const parseEncryptionInfo = (data: Uint8Array): ParsedEncryptionInfo => {
  if (data.length < MIN_HEADER_SIZE) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `EncryptionInfo too short: ${data.length} bytes (minimum ${MIN_HEADER_SIZE})`,
    });
  }

  const version = readUint16LE(data, 0);
  const reserved = readUint16LE(data, 2);

  if (version === AGILE_VERSION && reserved === AGILE_RESERVED) {
    return { type: "agile", params: parseAgileXml(data.subarray(MIN_HEADER_SIZE)) };
  }

  if ((version === 0x0003 || version === AGILE_VERSION) && reserved === STANDARD_RESERVED) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.ENCRYPTION_UNSUPPORTED,
      message:
        "Standard Encryption (Office 2007) is not supported. Only Agile Encryption (Office 2010+) is supported.",
    });
  }

  if (version <= 0x0002) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.ENCRYPTION_UNSUPPORTED,
      message: "Legacy RC4 encryption is not supported. Only Agile Encryption (Office 2010+) is supported.",
    });
  }

  throw new DocxEncryptionError({
    code: DOCX_ENCRYPTION_ERROR_CODES.ENCRYPTION_UNSUPPORTED,
    message: `Unrecognized EncryptionInfo version=${version} reserved=${reserved}`,
  });
};
