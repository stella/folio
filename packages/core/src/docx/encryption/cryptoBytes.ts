/** Byte helpers shared by MS-CFB and MS-OFFCRYPTO agile decryption. */

export const joinBytes = (parts: readonly Uint8Array[]): Uint8Array => {
  let total = 0;
  for (const part of parts) {
    total += part.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
};

export const writeUint32Le = (value: number): Uint8Array => {
  const out = new Uint8Array(4);
  out[0] = value & 0xff;
  out[1] = (value >>> 8) & 0xff;
  out[2] = (value >>> 16) & 0xff;
  out[3] = (value >>> 24) & 0xff;
  return out;
};

export const passwordToUtf16Le = (password: string): Uint8Array => {
  const out = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) {
    const code = password.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = code >> 8;
  }
  return out;
};

export const bytesEqual = (left: Uint8Array, right: Uint8Array): boolean => {
  if (left.length !== right.length) {
    return false;
  }
  let mismatch = 0;
  for (let i = 0; i < left.length; i++) {
    mismatch |= left[i]! ^ right[i]!;
  }
  return mismatch === 0;
};

export const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

export const decodeBase64 = (encoded: string): Uint8Array => {
  if (typeof atob === "function") {
    const binary = atob(encoded);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      out[i] = binary.charCodeAt(i);
    }
    return out;
  }
  return new Uint8Array(Buffer.from(encoded, "base64"));
};

export const padToBlock = (bytes: Uint8Array, blockSize: number, padByte = 0x36): Uint8Array => {
  if (bytes.length >= blockSize) {
    return bytes.subarray(0, blockSize);
  }
  const out = new Uint8Array(blockSize);
  out.set(bytes);
  out.fill(padByte, bytes.length);
  return out;
};
