/**
 * Agile Encryption decryption (Office 2010+) using Web Crypto.
 *
 * Algorithm steps follow [MS-OFFCRYPTO] sections 2.3.4.9 through 2.3.4.15.
 *
 * @see https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
 */

import {
  bytesEqual,
  joinBytes,
  padToBlock,
  passwordToUtf16Le,
  toArrayBuffer,
  writeUint32Le,
} from "./cryptoBytes";
import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";
import type { AgileKeyMaterial } from "./encryptionInfo";

/** [MS-OFFCRYPTO] 2.3.4.15 — block keys used during password verification. */
const PASSWORD_INPUT_BLOCK_KEY = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const PASSWORD_DIGEST_BLOCK_KEY = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const INTERMEDIATE_KEY_BLOCK_KEY = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
const INTEGRITY_KEY_BLOCK_KEY = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]);
const INTEGRITY_DIGEST_BLOCK_KEY = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]);

const PACKAGE_PLAINTEXT_SIZE_BYTES = 8;
const PACKAGE_SEGMENT_BYTES = 4096;

const subtleCrypto = (): SubtleCrypto => {
  const crypto = globalThis.crypto?.subtle;
  if (!crypto) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "Web Crypto API (SubtleCrypto) is not available in this environment",
    });
  }
  return crypto;
};

const webHashName = (officeName: string): AlgorithmIdentifier => {
  const aliases: Record<string, AlgorithmIdentifier> = {
    SHA1: "SHA-1",
    "SHA-1": "SHA-1",
    SHA256: "SHA-256",
    "SHA-256": "SHA-256",
    SHA384: "SHA-384",
    "SHA-384": "SHA-384",
    SHA512: "SHA-512",
    "SHA-512": "SHA-512",
  };
  const resolved = aliases[officeName];
  if (!resolved) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: `Unsupported hash algorithm: ${officeName}`,
    });
  }
  return resolved;
};

const digest = async (algorithm: AlgorithmIdentifier, input: Uint8Array): Promise<Uint8Array> => {
  const output = await subtleCrypto().digest(algorithm, toArrayBuffer(input));
  return new Uint8Array(output);
};

const importAesKey = (rawKey: Uint8Array, usages: KeyUsage[]): Promise<CryptoKey> =>
  subtleCrypto().importKey("raw", toArrayBuffer(rawKey), { name: "AES-CBC" }, false, usages);

const decryptAesCbc = async (
  rawKey: Uint8Array,
  iv: Uint8Array,
  ciphertext: Uint8Array,
): Promise<Uint8Array> => {
  const blockSize = 16;
  if (ciphertext.length === 0 || ciphertext.length % blockSize !== 0) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "AES ciphertext length must be a positive multiple of the block size",
    });
  }

  const subtle = subtleCrypto();
  const key = await importAesKey(rawKey, ["encrypt", "decrypt"]);

  // [MS-OFFCRYPTO] verifier fields are AES-CBC ciphertext without a PKCS#7
  // padding block. SubtleCrypto expects PKCS#7, so append a synthetic padding
  // block derived from the last ciphertext block (see MS-OFFCRYPTO 2.3.4.9).
  const lastBlock = ciphertext.subarray(ciphertext.length - blockSize);
  const pkcs7PadPlaintext = new Uint8Array(blockSize);
  pkcs7PadPlaintext.fill(blockSize);
  const syntheticPadding = new Uint8Array(
    await subtle.encrypt({ name: "AES-CBC", iv: toArrayBuffer(lastBlock) }, key, pkcs7PadPlaintext),
  ).subarray(0, blockSize);

  const paddedCiphertext = new Uint8Array(ciphertext.length + blockSize);
  paddedCiphertext.set(ciphertext);
  paddedCiphertext.set(syntheticPadding, ciphertext.length);

  const decrypted = await subtle.decrypt(
    { name: "AES-CBC", iv: toArrayBuffer(iv) },
    key,
    toArrayBuffer(paddedCiphertext),
  );
  return new Uint8Array(decrypted).subarray(0, ciphertext.length);
};

/** [MS-OFFCRYPTO] 2.3.4.9 — iterated hash of password + salt. */
const hashPassword = async (
  password: string,
  salt: Uint8Array,
  iterations: number,
  hashName: string,
): Promise<Uint8Array> => {
  const algorithm = webHashName(hashName);
  let state = await digest(algorithm, joinBytes([salt, passwordToUtf16Le(password)]));
  for (let i = 0; i < iterations; i++) {
    state = await digest(algorithm, joinBytes([writeUint32Le(i), state]));
  }
  return state;
};

/** [MS-OFFCRYPTO] 2.3.4.9 — finalize a derived key with a block key suffix. */
const finalizeDerivedKey = async (
  passwordHash: Uint8Array,
  blockKey: Uint8Array,
  hashName: string,
  keyBits: number,
): Promise<Uint8Array> => {
  const algorithm = webHashName(hashName);
  const hashed = await digest(algorithm, joinBytes([passwordHash, blockKey]));
  return hashed.subarray(0, keyBits / 8);
};

const segmentIv = async (
  packageSalt: Uint8Array,
  segmentIndex: number,
  hashName: string,
  blockBytes: number,
): Promise<Uint8Array> => {
  const hashed = await digest(
    webHashName(hashName),
    joinBytes([packageSalt, writeUint32Le(segmentIndex)]),
  );
  return hashed.subarray(0, blockBytes);
};

const integrityIv = async (
  packageSalt: Uint8Array,
  blockKey: Uint8Array,
  hashName: string,
  blockBytes: number,
): Promise<Uint8Array> => {
  const hashed = await digest(webHashName(hashName), joinBytes([packageSalt, blockKey]));
  return hashed.subarray(0, blockBytes);
};

const unlockIntermediateKey = async (
  password: string,
  material: AgileKeyMaterial,
): Promise<Uint8Array> => {
  const passwordHash = await hashPassword(
    password,
    material.passwordSalt,
    material.passwordIterations,
    material.passwordHash,
  );

  const verifierKey = await finalizeDerivedKey(
    passwordHash,
    PASSWORD_INPUT_BLOCK_KEY,
    material.passwordHash,
    material.passwordKeyBits,
  );
  const digestKey = await finalizeDerivedKey(
    passwordHash,
    PASSWORD_DIGEST_BLOCK_KEY,
    material.passwordHash,
    material.passwordKeyBits,
  );
  const intermediateKey = await finalizeDerivedKey(
    passwordHash,
    INTERMEDIATE_KEY_BLOCK_KEY,
    material.passwordHash,
    material.passwordKeyBits,
  );

  const verifierIv = padToBlock(material.passwordSalt, material.passwordBlockBytes);

  const verifierPlaintext = await decryptAesCbc(
    verifierKey,
    verifierIv,
    material.encryptedVerifierInput,
  );
  const digestPlaintext = await decryptAesCbc(
    digestKey,
    verifierIv,
    material.encryptedVerifierDigest,
  );

  const actualDigest = await digest(webHashName(material.passwordHash), verifierPlaintext);
  if (!bytesEqual(actualDigest, digestPlaintext.subarray(0, actualDigest.length))) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.PASSWORD_INVALID,
      message: "The password is incorrect",
    });
  }

  return decryptAesCbc(intermediateKey, verifierIv, material.encryptedIntermediateKey);
};

const assertPackageIntegrity = async (
  intermediateKey: Uint8Array,
  material: AgileKeyMaterial,
  encryptedPackage: Uint8Array,
): Promise<void> => {
  const ivForKey = await integrityIv(
    material.packageSalt,
    INTEGRITY_KEY_BLOCK_KEY,
    material.packageHash,
    material.cipherBlockBytes,
  );
  const ivForDigest = await integrityIv(
    material.packageSalt,
    INTEGRITY_DIGEST_BLOCK_KEY,
    material.packageHash,
    material.cipherBlockBytes,
  );

  const integrityKey = await decryptAesCbc(
    intermediateKey,
    ivForKey,
    material.encryptedIntegrityKey,
  );
  const expectedDigest = await decryptAesCbc(
    intermediateKey,
    ivForDigest,
    material.encryptedIntegrityDigest,
  );

  const hmacKey = await subtleCrypto().importKey(
    "raw",
    toArrayBuffer(integrityKey),
    { name: "HMAC", hash: webHashName(material.packageHash) },
    false,
    ["sign"],
  );
  const actualDigest = new Uint8Array(
    await subtleCrypto().sign("HMAC", hmacKey, toArrayBuffer(encryptedPackage)),
  );

  if (!bytesEqual(actualDigest, expectedDigest.subarray(0, actualDigest.length))) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message:
        "Data integrity check failed — the encrypted package may be corrupt or tampered with",
    });
  }
};

const decryptEncryptedPackage = async (
  intermediateKey: Uint8Array,
  material: AgileKeyMaterial,
  encryptedPackage: Uint8Array,
): Promise<Uint8Array> => {
  if (encryptedPackage.length < PACKAGE_PLAINTEXT_SIZE_BYTES) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "EncryptedPackage stream is too short to contain a size prefix",
    });
  }

  const sizeView = new DataView(
    encryptedPackage.buffer,
    encryptedPackage.byteOffset,
    PACKAGE_PLAINTEXT_SIZE_BYTES,
  );
  const plaintextSizeBig = sizeView.getBigUint64(0, true);
  if (plaintextSizeBig > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "EncryptedPackage plaintext size exceeds JavaScript's safe integer range",
    });
  }
  const plaintextSize = Number(plaintextSizeBig);

  const ciphertext = encryptedPackage.subarray(PACKAGE_PLAINTEXT_SIZE_BYTES);
  const segmentCount = Math.ceil(ciphertext.length / PACKAGE_SEGMENT_BYTES);
  const segments: Uint8Array[] = [];

  for (let index = 0; index < segmentCount; index++) {
    const start = index * PACKAGE_SEGMENT_BYTES;
    const end = Math.min(start + PACKAGE_SEGMENT_BYTES, ciphertext.length);
    const chunk = ciphertext.subarray(start, end);
    const iv = await segmentIv(
      material.packageSalt,
      index,
      material.packageHash,
      material.cipherBlockBytes,
    );
    segments.push(await decryptAesCbc(intermediateKey, iv, chunk));
  }

  const plaintext = joinBytes(segments);
  if (plaintextSize > plaintext.length) {
    throw new DocxEncryptionError({
      code: DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
      message: "EncryptedPackage plaintext size exceeds decrypted payload length",
    });
  }
  return plaintext.subarray(0, plaintextSize);
};

export const decryptAgileEncryptedPackage = async (
  password: string,
  material: AgileKeyMaterial,
  encryptedPackage: Uint8Array,
): Promise<Uint8Array> => {
  try {
    const intermediateKey = await unlockIntermediateKey(password, material);
    await assertPackageIntegrity(intermediateKey, material, encryptedPackage);
    return await decryptEncryptedPackage(intermediateKey, material, encryptedPackage);
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
