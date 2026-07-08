/**
 * Parse the EncryptionInfo stream for Agile Encryption (Office 2010+).
 *
 * @see https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/87020a34-e73f-4139-99bc-bbdf6cf6fa55
 */

import { decodeBase64 } from "./cryptoBytes";
import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";

export type AgileKeyMaterial = {
  packageSalt: Uint8Array;
  packageKeyBits: number;
  packageHash: string;
  packageHashBytes: number;
  cipherBlockBytes: number;
  cipherName: string;
  cipherMode: string;
  passwordIterations: number;
  passwordSalt: Uint8Array;
  passwordHash: string;
  passwordKeyBits: number;
  passwordBlockBytes: number;
  encryptedVerifierInput: Uint8Array;
  encryptedVerifierDigest: Uint8Array;
  encryptedIntermediateKey: Uint8Array;
  encryptedIntegrityKey: Uint8Array;
  encryptedIntegrityDigest: Uint8Array;
};

export type AgileEncryptionDescriptor = {
  scheme: "agile";
  material: AgileKeyMaterial;
};

const AGILE_STREAM_VERSION = 0x0004;
const AGILE_STREAM_KIND = 0x0004;
const STANDARD_STREAM_KIND = 0x0003;
const STREAM_PREFIX_BYTES = 8;

const readUint16Le = (bytes: Uint8Array, offset: number): number =>
  bytes[offset]! | (bytes[offset + 1]! << 8);

const findElement = (xml: string, localName: string): string | null => {
  let index = 0;
  while (index < xml.length) {
    const hit = xml.indexOf(localName, index);
    if (hit === -1) {
      return null;
    }
    const tagStart = xml.lastIndexOf("<", hit);
    if (tagStart === -1 || hit - tagStart > 24) {
      index = hit + localName.length;
      continue;
    }
    const tagEnd = xml.indexOf(">", hit);
    if (tagEnd === -1) {
      return null;
    }
    const tag = xml.slice(tagStart, tagEnd + 1);
    if (tag.includes(localName)) {
      return tag;
    }
    index = hit + localName.length;
  }
  return null;
};

const parseAttributes = (tag: string): Record<string, string> => {
  const attrs: Record<string, string> = {};
  let i = tag.indexOf(" ");
  if (i === -1) {
    return attrs;
  }
  while (i < tag.length) {
    while (i < tag.length && (tag[i] === " " || tag[i] === "\n" || tag[i] === "\r" || tag[i] === "\t")) {
      i++;
    }
    const nameStart = i;
    while (i < tag.length && tag[i] !== "=" && tag[i] !== ">" && tag[i] !== "/") {
      i++;
    }
    if (i >= tag.length || tag[i] !== "=") {
      break;
    }
    const name = tag.slice(nameStart, i).trim();
    i++;
    if (tag[i] !== '"') {
      break;
    }
    i++;
    const valueStart = i;
    while (i < tag.length && tag[i] !== '"') {
      i++;
    }
    attrs[name] = tag.slice(valueStart, i);
    i++;
  }
  return attrs;
};

const requireAttr = (
  attrs: Record<string, string>,
  name: string,
  element: string,
): string => {
  const value = attrs[name];
  if (value === undefined) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `EncryptionInfo <${element}> missing attribute "${name}"`,
    });
  }
  return value;
};

const requireInt = (attrs: Record<string, string>, name: string, element: string): number => {
  const parsed = Number.parseInt(requireAttr(attrs, name, element), 10);
  if (Number.isNaN(parsed)) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `EncryptionInfo <${element}> attribute "${name}" is not an integer`,
    });
  }
  return parsed;
};

const requireBase64 = (attrs: Record<string, string>, name: string, element: string): Uint8Array =>
  decodeBase64(requireAttr(attrs, name, element));

const parseAgileXml = (xmlBytes: Uint8Array): AgileKeyMaterial => {
  const xml = new TextDecoder("utf-8").decode(xmlBytes);
  const keyDataTag = findElement(xml, "keyData");
  const integrityTag = findElement(xml, "dataIntegrity");
  const encryptedKeyTag = findElement(xml, "encryptedKey");
  if (!keyDataTag || !integrityTag || !encryptedKeyTag) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "EncryptionInfo XML is missing required agile encryption elements",
    });
  }

  const keyData = parseAttributes(keyDataTag);
  const integrity = parseAttributes(integrityTag);
  const encryptedKey = parseAttributes(encryptedKeyTag);

  return {
    packageSalt: requireBase64(keyData, "saltValue", "keyData"),
    packageKeyBits: requireInt(keyData, "keyBits", "keyData"),
    packageHash: requireAttr(keyData, "hashAlgorithm", "keyData"),
    packageHashBytes: requireInt(keyData, "hashSize", "keyData"),
    cipherBlockBytes: requireInt(keyData, "blockSize", "keyData"),
    cipherName: requireAttr(keyData, "cipherAlgorithm", "keyData"),
    cipherMode: requireAttr(keyData, "cipherChaining", "keyData"),
    passwordIterations: requireInt(encryptedKey, "spinCount", "encryptedKey"),
    passwordSalt: requireBase64(encryptedKey, "saltValue", "encryptedKey"),
    passwordHash: requireAttr(encryptedKey, "hashAlgorithm", "encryptedKey"),
    passwordKeyBits: requireInt(encryptedKey, "keyBits", "encryptedKey"),
    passwordBlockBytes: requireInt(encryptedKey, "blockSize", "encryptedKey"),
    encryptedVerifierInput: requireBase64(encryptedKey, "encryptedVerifierHashInput", "encryptedKey"),
    encryptedVerifierDigest: requireBase64(
      encryptedKey,
      "encryptedVerifierHashValue",
      "encryptedKey",
    ),
    encryptedIntermediateKey: requireBase64(encryptedKey, "encryptedKeyValue", "encryptedKey"),
    encryptedIntegrityKey: requireBase64(integrity, "encryptedHmacKey", "dataIntegrity"),
    encryptedIntegrityDigest: requireBase64(integrity, "encryptedHmacValue", "dataIntegrity"),
  };
};

export const parseAgileEncryptionInfo = (stream: Uint8Array): AgileEncryptionDescriptor => {
  if (stream.length < STREAM_PREFIX_BYTES) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `EncryptionInfo stream too short (${stream.length} bytes)`,
    });
  }

  const version = readUint16Le(stream, 0);
  const kind = readUint16Le(stream, 2);

  if (version === AGILE_STREAM_VERSION && kind === AGILE_STREAM_KIND) {
    return { scheme: "agile", material: parseAgileXml(stream.subarray(STREAM_PREFIX_BYTES)) };
  }

  if ((version === 0x0003 || version === AGILE_STREAM_VERSION) && kind === STANDARD_STREAM_KIND) {
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
    message: `Unrecognized EncryptionInfo header version=${version} kind=${kind}`,
  });
};
