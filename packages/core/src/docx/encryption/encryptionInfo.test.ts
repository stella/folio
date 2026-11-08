import { parseAgileEncryptionInfo } from "./encryptionInfo";
import { DOCX_ENCRYPTION_ERROR_CODES, DocxEncryptionError } from "./errors";

const SALT = "AAAAAAAAAAAAAAAAAAAAAA==";
const BLOCK = "AAAAAAAAAAAAAAAAAAAAAA==";

type DescriptorOptions = {
  keyDataBlockSize: string;
  encryptedKeyBlockSize: string;
};

const agileStream = ({
  keyDataBlockSize,
  encryptedKeyBlockSize,
}: DescriptorOptions): Uint8Array => {
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<keyData saltSize="16" blockSize="${keyDataBlockSize}" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${SALT}"/>` +
    `<dataIntegrity encryptedHmacKey="${BLOCK}" encryptedHmacValue="${BLOCK}"/>` +
    `<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<p:encryptedKey spinCount="0" saltSize="16" blockSize="${encryptedKeyBlockSize}" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${SALT}" encryptedVerifierHashInput="${BLOCK}" encryptedVerifierHashValue="${BLOCK}" encryptedKeyValue="${BLOCK}"/>` +
    `</keyEncryptor></keyEncryptors></encryption>`;
  const body = new TextEncoder().encode(xml);
  const stream = new Uint8Array(8 + body.length);
  stream.set([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]);
  stream.set(body, 8);
  return stream;
};

const expectDecryptionFailure = (stream: Uint8Array): void => {
  try {
    parseAgileEncryptionInfo(stream);
    expect.unreachable("Expected DocxEncryptionError");
  } catch (error) {
    expect(error).toBeInstanceOf(DocxEncryptionError);
    expect((error as DocxEncryptionError).code).toBe(
      DOCX_ENCRYPTION_ERROR_CODES.DECRYPTION_FAILED,
    );
  }
};

describe("parseAgileEncryptionInfo block sizes", () => {
  test("accepts the AES block size on both key descriptors", () => {
    const descriptor = parseAgileEncryptionInfo(
      agileStream({ keyDataBlockSize: "16", encryptedKeyBlockSize: "16" }),
    );
    expect(descriptor.material.cipherBlockBytes).toBe(16);
    expect(descriptor.material.passwordBlockBytes).toBe(16);
  });

  test.each(["0", "-16", "15", "32", "4096", String(64 * 1024 * 1024)])(
    "rejects password key block size %s while parsing",
    (blockSize) => {
      expectDecryptionFailure(
        agileStream({
          keyDataBlockSize: "16",
          encryptedKeyBlockSize: blockSize,
        }),
      );
    },
  );

  test.each(["0", "8", String(64 * 1024 * 1024)])(
    "rejects package key block size %s while parsing",
    (blockSize) => {
      expectDecryptionFailure(
        agileStream({
          keyDataBlockSize: blockSize,
          encryptedKeyBlockSize: "16",
        }),
      );
    },
  );
});
