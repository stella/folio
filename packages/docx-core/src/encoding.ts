import { panic } from "better-result";

/** Encode complete text before any transport divides its byte buffer. */
export const encodeUtf8 = (text: string): Uint8Array<ArrayBuffer> => new TextEncoder().encode(text);

type DecodeTextOptions = NonNullable<ConstructorParameters<typeof TextDecoder>[1]> & {
  encoding?: string;
};

/** Decode a complete byte buffer; chunked input must use a text reader. */
export const decodeText = (
  bytes: Uint8Array,
  { encoding = "utf-8", ...options }: DecodeTextOptions = {},
): string => new TextDecoder(encoding, options).decode(bytes);

/** Clipboard and file blobs are decoded from a complete byte buffer. */
export const readBlobText = async (blob: Blob): Promise<string> =>
  decodeText(new Uint8Array(await blob.arrayBuffer()));

/** One decoder retains incomplete scalars between byte chunks. */
export const createTextReader = ({ encoding = "utf-8", ...options }: DecodeTextOptions = {}) => {
  const decoder = new TextDecoder(encoding, options);
  return {
    read: (bytes: Uint8Array): string => decoder.decode(bytes, { stream: true }),
    finish: (): string => decoder.decode(),
    // A bounded prefix omits an incomplete trailing scalar rather than replacing it.
    finishPrefix: (): string => "",
  };
};

/** UTF-16BE is required by PDF text strings and Unicode character maps. */
export const encodeUtf16BE = (text: string): Uint8Array<ArrayBuffer> => {
  const bytes = new Uint8Array(text.length * 2);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < text.length; index++)
    view.setUint16(index * 2, text.charCodeAt(index));
  return bytes;
};

/** Binary strings carry byte values, rather than Unicode text. */
export const encodeBinaryString = (text: string): Uint8Array<ArrayBuffer> =>
  Uint8Array.from(text, (character) => {
    const byte = character.charCodeAt(0);
    if (byte > 0xff) return panic("Binary strings must contain only byte values");
    return byte;
  });

export const bytesToHex = (bytes: Uint8Array): string =>
  Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");

const BASE64_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
const BASE64_CODES = Uint8Array.from(BASE64_ALPHABET, (character) => character.charCodeAt(0));
const BASE64_PAD = 61;

/**
 * Encode into exactly one intermediate buffer: ASCII codes are written into a
 * single array and decoded once, so a megabyte-scale input leaves the output
 * string and nothing else.
 *
 * Exported because the runtimes the tests run on all carry a native encoder,
 * and the fallback is the branch the browsers that lack one take.
 */
export const encodeBase64Portable = (bytes: Uint8Array): string => {
  const groups = Math.ceil(bytes.length / 3);
  const encoded = new Uint8Array(groups * 4);
  let read = 0;
  let write = 0;
  for (; read + 2 < bytes.length; read += 3) {
    // SAFETY: the loop condition keeps all three reads inside the buffer.
    const triple =
      ((bytes[read] as number) << 16) |
      ((bytes[read + 1] as number) << 8) |
      (bytes[read + 2] as number);
    encoded[write] = BASE64_CODES[(triple >> 18) & 63] as number;
    encoded[write + 1] = BASE64_CODES[(triple >> 12) & 63] as number;
    encoded[write + 2] = BASE64_CODES[(triple >> 6) & 63] as number;
    encoded[write + 3] = BASE64_CODES[triple & 63] as number;
    write += 4;
  }
  const remaining = bytes.length - read;
  if (remaining > 0) {
    // SAFETY: `remaining` is 1 or 2, so `read` and the guarded `read + 1` are in range.
    const tail =
      ((bytes[read] as number) << 16) | (remaining === 2 ? (bytes[read + 1] as number) << 8 : 0);
    encoded[write] = BASE64_CODES[(tail >> 18) & 63] as number;
    encoded[write + 1] = BASE64_CODES[(tail >> 12) & 63] as number;
    encoded[write + 2] = remaining === 2 ? (BASE64_CODES[(tail >> 6) & 63] as number) : BASE64_PAD;
    encoded[write + 3] = BASE64_PAD;
  }
  // Every byte written is a base64 character, so UTF-8 decoding is exact.
  return decodeText(encoded);
};

/**
 * `Uint8Array.prototype.toBase64` is roughly nine times the portable encoder's
 * throughput on a multi-megabyte buffer, and the capability is a property of
 * the runtime rather than of the call, so it is probed once.
 */
const nativeToBase64 =
  "toBase64" in Uint8Array.prototype ? (bytes: Uint8Array) => bytes.toBase64() : undefined;

/** Encode bytes as canonical, padded, standard-alphabet base64. */
export const bytesToBase64 = nativeToBase64 ?? encodeBase64Portable;

/** Decode standard or URL-safe base64 into bytes. */
export const base64ToBytes = (text: string): Uint8Array<ArrayBuffer> =>
  Uint8Array.from(atob(text.replaceAll("-", "+").replaceAll("_", "/")), (character) =>
    character.charCodeAt(0),
  );

export const bytesToBase64Url = (bytes: Uint8Array): string =>
  bytesToBase64(bytes).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/u, "");
