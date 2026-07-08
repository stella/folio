/**
 * Agile Encryption decryption via Web Crypto (MS-OFFCRYPTO).
 *
 * @see https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
 */

import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";
import type { AgileEncryptionParams } from "./parseEncryptionInfo";

const BLOCK_KEY_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_KEY_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_ENCRYPTED_KEY = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
const BLOCK_KEY_HMAC_KEY = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]);
const BLOCK_KEY_HMAC_VALUE = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]);

const SEGMENT_SIZE = 4096;
const PACKAGE_HEADER_SIZE = 8;

const getSubtleCrypto = (): SubtleCrypto => {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "Web Crypto API (SubtleCrypto) is not available in this environment",
    });
  }
  return subtle;
};

const toWebCryptoHash = (ooxmlName: string): AlgorithmIdentifier => {
  const map: Record<string, string> = {
    SHA1: "SHA-1",
    "SHA-1": "SHA-1",
    SHA256: "SHA-256",
    "SHA-256": "SHA-256",
    SHA384: "SHA-384",
    "SHA-384": "SHA-384",
    SHA512: "SHA-512",
    "SHA-512": "SHA-512",
  };
  const mapped = map[ooxmlName];
  if (!mapped) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `Unsupported hash algorithm: ${ooxmlName}`,
    });
  }
  return mapped;
};

const encodeUtf16le = (str: string): Uint8Array => {
  const buf = new Uint8Array(str.length * 2);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    buf[i * 2] = code & 0xff;
    buf[i * 2 + 1] = (code >> 8) & 0xff;
  }
  return buf;
};

const concatBytes = (...arrays: Uint8Array[]): Uint8Array => {
  const totalLength = arrays.reduce((sum, array) => sum + array.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  for (const array of arrays) {
    result.set(array, offset);
    offset += array.length;
  }
  return result;
};

const uint32LE = (value: number): Uint8Array => {
  const buf = new Uint8Array(4);
  buf[0] = value & 0xff;
  buf[1] = (value >>> 8) & 0xff;
  buf[2] = (value >>> 16) & 0xff;
  buf[3] = (value >>> 24) & 0xff;
  return buf;
};

const toBuffer = (arr: Uint8Array): ArrayBuffer =>
  arr.buffer.slice(arr.byteOffset, arr.byteOffset + arr.byteLength) as ArrayBuffer;

const hash = async (algorithm: AlgorithmIdentifier, data: Uint8Array): Promise<Uint8Array> => {
  const digest = await getSubtleCrypto().digest(algorithm, toBuffer(data));
  return new Uint8Array(digest);
};

const decryptAesCbc = async (
  keyBytes: Uint8Array,
  iv: Uint8Array,
  ciphertext: Uint8Array,
): Promise<Uint8Array> => {
  const subtle = getSubtleCrypto();
  const blockSize = 16;

  const key = await subtle.importKey("raw", toBuffer(keyBytes), { name: "AES-CBC" }, false, [
    "encrypt",
    "decrypt",
  ]);

  const lastBlock = ciphertext.subarray(ciphertext.length - blockSize);
  const pkcs7Plaintext = new Uint8Array(blockSize);
  pkcs7Plaintext.fill(blockSize);

  const encryptedPadding = new Uint8Array(
    await subtle.encrypt({ name: "AES-CBC", iv: toBuffer(lastBlock) }, key, pkcs7Plaintext),
  );
  const paddingBlock = encryptedPadding.subarray(0, blockSize);

  const withPadding = new Uint8Array(ciphertext.length + blockSize);
  withPadding.set(ciphertext);
  withPadding.set(paddingBlock, ciphertext.length);

  const decrypted = await subtle.decrypt({ name: "AES-CBC", iv: toBuffer(iv) }, key, withPadding);
  return new Uint8Array(decrypted).subarray(0, ciphertext.length);
};

const deriveKey = async (
  password: string,
  salt: Uint8Array,
  spinCount: number,
  keyBits: number,
  hashAlgorithm: string,
  blockKey: Uint8Array,
): Promise<Uint8Array> => {
  const algo = toWebCryptoHash(hashAlgorithm);
  const passwordBytes = encodeUtf16le(password);

  let h = await hash(algo, concatBytes(salt, passwordBytes));
  for (let i = 0; i < spinCount; i++) {
    h = await hash(algo, concatBytes(uint32LE(i), h));
  }

  const hDerived = await hash(algo, concatBytes(h, blockKey));
  const requiredBytes = keyBits / 8;
  if (hDerived.length >= requiredBytes) {
    return hDerived.subarray(0, requiredBytes);
  }

  const padded = new Uint8Array(requiredBytes);
  padded.set(hDerived);
  padded.fill(0x36, hDerived.length);
  return padded;
};

const generateIV = async (
  hashAlgorithm: string,
  salt: Uint8Array,
  blockKey: Uint8Array | undefined,
  blockSize: number,
): Promise<Uint8Array> => {
  const ivSource = blockKey
    ? await hash(toWebCryptoHash(hashAlgorithm), concatBytes(salt, blockKey))
    : salt;

  if (ivSource.length >= blockSize) {
    return ivSource.subarray(0, blockSize);
  }

  const padded = new Uint8Array(blockSize);
  padded.set(ivSource);
  padded.fill(0x36, ivSource.length);
  return padded;
};

const constantTimeEqual = (a: Uint8Array, b: Uint8Array): boolean => {
  if (a.length !== b.length) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a[i]! ^ b[i]!;
  }
  return diff === 0;
};

const verifyPasswordAndGetKey = async (
  password: string,
  params: AgileEncryptionParams,
): Promise<Uint8Array> => {
  const { passwordSalt, spinCount, passwordHashAlgorithm, passwordKeyBits, passwordBlockSize } = params;

  const keyVerifierInput = await deriveKey(
    password,
    passwordSalt,
    spinCount,
    passwordKeyBits,
    passwordHashAlgorithm,
    BLOCK_KEY_VERIFIER_INPUT,
  );
  const keyVerifierValue = await deriveKey(
    password,
    passwordSalt,
    spinCount,
    passwordKeyBits,
    passwordHashAlgorithm,
    BLOCK_KEY_VERIFIER_VALUE,
  );
  const keyEncryptedKey = await deriveKey(
    password,
    passwordSalt,
    spinCount,
    passwordKeyBits,
    passwordHashAlgorithm,
    BLOCK_KEY_ENCRYPTED_KEY,
  );

  const passwordVerifierIv = await generateIV(
    passwordHashAlgorithm,
    passwordSalt,
    undefined,
    passwordBlockSize,
  );

  const verifierHashInput = await decryptAesCbc(
    keyVerifierInput,
    passwordVerifierIv,
    params.encryptedVerifierHashInput,
  );
  const verifierHashValue = await decryptAesCbc(
    keyVerifierValue,
    passwordVerifierIv,
    params.encryptedVerifierHashValue,
  );

  const computedHash = await hash(toWebCryptoHash(passwordHashAlgorithm), verifierHashInput);
  const expectedHash = verifierHashValue.subarray(0, computedHash.length);
  if (!constantTimeEqual(computedHash, expectedHash)) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.PASSWORD_INVALID,
      message: "The password is incorrect",
    });
  }

  return decryptAesCbc(keyEncryptedKey, passwordVerifierIv, params.encryptedKeyValue);
};

const decryptPackageSegments = async (
  encryptionKey: Uint8Array,
  params: AgileEncryptionParams,
  encryptedPackage: Uint8Array,
): Promise<Uint8Array> => {
  const { keySalt, hashAlgorithm, blockSize } = params;
  const algo = toWebCryptoHash(hashAlgorithm);

  if (encryptedPackage.length < PACKAGE_HEADER_SIZE) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "EncryptedPackage too short to contain size header",
    });
  }

  const plaintextSize =
    encryptedPackage[0]! |
    (encryptedPackage[1]! << 8) |
    (encryptedPackage[2]! << 16) |
    ((encryptedPackage[3]! << 24) >>> 0);

  const encryptedData = encryptedPackage.subarray(PACKAGE_HEADER_SIZE);
  const segmentCount = Math.ceil(encryptedData.length / SEGMENT_SIZE);
  const decryptedChunks: Uint8Array[] = [];

  for (let i = 0; i < segmentCount; i++) {
    const segmentStart = i * SEGMENT_SIZE;
    const segmentEnd = Math.min(segmentStart + SEGMENT_SIZE, encryptedData.length);
    const segment = encryptedData.subarray(segmentStart, segmentEnd);

    const ivHash = await hash(algo, concatBytes(keySalt, uint32LE(i)));
    const iv = ivHash.length >= blockSize ? ivHash.subarray(0, blockSize) : ivHash;
    decryptedChunks.push(await decryptAesCbc(encryptionKey, iv, segment));
  }

  return concatBytes(...decryptedChunks).subarray(0, plaintextSize);
};

const verifyDataIntegrity = async (
  encryptionKey: Uint8Array,
  params: AgileEncryptionParams,
  encryptedPackage: Uint8Array,
): Promise<void> => {
  const { keySalt, hashAlgorithm, blockSize, hashSize } = params;
  const algo = toWebCryptoHash(hashAlgorithm);
  const subtle = getSubtleCrypto();

  const ivHmacKey = await generateIV(hashAlgorithm, keySalt, BLOCK_KEY_HMAC_KEY, blockSize);
  const hmacKeyRaw = await decryptAesCbc(encryptionKey, ivHmacKey, params.encryptedHmacKey);
  const hmacKey = hmacKeyRaw.subarray(0, hashSize);

  const ivHmacValue = await generateIV(hashAlgorithm, keySalt, BLOCK_KEY_HMAC_VALUE, blockSize);
  const expectedHmac = await decryptAesCbc(encryptionKey, ivHmacValue, params.encryptedHmacValue);

  const cryptoKey = await subtle.importKey("raw", toBuffer(hmacKey), { name: "HMAC", hash: algo }, false, [
    "sign",
  ]);
  const computedHmac = new Uint8Array(
    await subtle.sign("HMAC", cryptoKey, toBuffer(encryptedPackage)),
  );

  if (!constantTimeEqual(computedHmac, expectedHmac.subarray(0, computedHmac.length))) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message:
        "Data integrity check failed — the encrypted package may be corrupt or tampered with",
    });
  }
};

export const decryptAgilePackage = async (
  password: string,
  params: AgileEncryptionParams,
  encryptedPackage: Uint8Array,
): Promise<Uint8Array> => {
  try {
    const encryptionKey = await verifyPasswordAndGetKey(password, params);
    await verifyDataIntegrity(encryptionKey, params, encryptedPackage);
    return await decryptPackageSegments(encryptionKey, params, encryptedPackage);
  } catch (error) {
    if (error instanceof DocxEncryptionError) {
      throw error;
    }
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "Decryption failed unexpectedly",
      cause: error,
    });
  }
};
