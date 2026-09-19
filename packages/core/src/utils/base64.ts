/**
 * The package's one bytes-to-base64 encoder.
 *
 * Every call site that spelled this itself got a different answer to the same
 * two questions, and one of them got the first question wrong:
 *
 * - `btoa(new TextDecoder("latin1").decode(bytes))` is not a binary string.
 *   The Encoding standard maps the `latin1` label to windows-1252, so bytes
 *   0x80-0x9F decode to characters outside Latin-1 and `btoa` throws
 *   `InvalidCharacterError` on them. Ordinary image bytes contain those values,
 *   so the browser path of that spelling fails on real input.
 * - `String.fromCharCode(...chunk)` is correct but spreads a chunk's worth of
 *   arguments per call, and the intermediate binary string is as large as the
 *   input.
 *
 * Neither question has to be asked per call site, so it is asked here once.
 * The runtime's own encoder answers it when the runtime has one, and the
 * portable encoder below answers it when not; both produce canonical base64,
 * so the choice is invisible in the output.
 */

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
  return new TextDecoder().decode(encoded);
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

/** Encode bytes as a `data:` URL with the given media type. */
export const bytesToDataUrl = (bytes: Uint8Array, mimeType: string): string =>
  `data:${mimeType};base64,${bytesToBase64(bytes)}`;
