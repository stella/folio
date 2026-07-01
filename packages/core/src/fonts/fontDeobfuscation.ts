/**
 * Embedded-font de-obfuscation (ECMA-376 Part 4 §2.8.1, "Font Embedding").
 *
 * Word ships embedded fonts as obfuscated OpenType (`.odttf`): the first 32
 * bytes of the font binary are XOR-scrambled with a 16-byte key; the rest of
 * the file is untouched. The key is the font's `w:fontKey` GUID with its byte
 * order reversed, applied to bytes 0-15 and again to bytes 16-31.
 *
 * Example (from the spec): GUID `001B70DC-AA60-4AD5-90EC-18A0948E1EAE` yields
 * key bytes `AE 1E 8E 94 A0 18 EC 90 D5 4A 60 AA DC 70 1B 00`. The scheme is a
 * pure XOR, so the same operation obfuscates and de-obfuscates.
 *
 * Ported from eigenpal/docx-editor (see NOTICE.md). The ODTTF scheme is an
 * OOXML interop format, so the algorithm is reproduced faithfully.
 */

const HEADER_LENGTH = 32;
const KEY_LENGTH = 16;

/**
 * Strip a `w:fontKey` GUID down to its 16 key bytes, in the reversed order the
 * obfuscation applies. Returns null when the value is not a 32-hex-digit GUID.
 */
function fontKeyToReversedBytes(fontKey: string): Uint8Array | null {
  const hex = fontKey.replace(/[^0-9a-fA-F]/gu, "");
  if (hex.length !== KEY_LENGTH * 2) {
    return null;
  }

  const reversed = new Uint8Array(KEY_LENGTH);
  for (let i = 0; i < KEY_LENGTH; i++) {
    const byte = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16);
    if (Number.isNaN(byte)) {
      return null;
    }
    // Reverse the byte order while writing — this is the actual XOR key.
    reversed[KEY_LENGTH - 1 - i] = byte;
  }
  return reversed;
}

/**
 * Whether a string is a usable embedded-font obfuscation key: a 128-bit GUID,
 * with or without braces and hyphens.
 */
export function isValidFontKey(fontKey: string | undefined | null): boolean {
  if (!fontKey) {
    return false;
  }
  return fontKeyToReversedBytes(fontKey) !== null;
}

/**
 * De-obfuscate an embedded `.odttf` font into a usable OpenType/TrueType binary
 * by XOR-ing its first 32 bytes with the reversed `w:fontKey` GUID. Returns a
 * new array; the input is not mutated. Throws when `fontKey` is not a valid
 * 128-bit GUID (guard with {@link isValidFontKey} first).
 */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array<ArrayBuffer> {
  const key = fontKeyToReversedBytes(fontKey);
  if (!key) {
    throw new Error(`Invalid embedded-font key: "${fontKey}"`);
  }

  const out = new Uint8Array(data);
  const end = Math.min(HEADER_LENGTH, out.length);
  for (let i = 0; i < end; i++) {
    // In-bounds by construction: i < end <= out.length, and i % KEY_LENGTH
    // indexes the 16-byte key; `?? 0` only satisfies noUncheckedIndexedAccess.
    const source = out[i] ?? 0;
    const keyByte = key[i % KEY_LENGTH] ?? 0;
    out[i] = source ^ keyByte;
  }
  return out;
}
