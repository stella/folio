/**
 * A 128-bit hash of a string, for cache keys that must not keep the string.
 *
 * MurmurHash3's x86 128-bit construction over the string's UTF-16 code units,
 * two to a 32-bit word. It is not cryptographic: two strings an author did not
 * craft to collide share a hash with a chance of about n² / 2¹²⁹ among n keys.
 */

const C1 = 0x239b961b;
const C2 = 0xab0e9789;
const C3 = 0x38b34ae5;
const C4 = 0xa1e38b93;

const rotl = (value: number, by: number): number =>
  (value << by) | (value >>> (32 - by));

const fmix = (value: number): number => {
  let h = value;
  h ^= h >>> 16;
  h = Math.imul(h, 0x85ebca6b);
  h ^= h >>> 13;
  h = Math.imul(h, 0xc2b2ae35);
  h ^= h >>> 16;
  return h;
};

/** The word of code units `at` and `at + 1`, zero past the end. */
const word = (text: string, at: number): number =>
  (at < text.length ? text.charCodeAt(at) : 0) |
  (at + 1 < text.length ? text.charCodeAt(at + 1) << 16 : 0);

const mixK1 = (k: number): number => Math.imul(rotl(Math.imul(k, C1), 15), C2);
const mixK2 = (k: number): number => Math.imul(rotl(Math.imul(k, C2), 16), C3);
const mixK3 = (k: number): number => Math.imul(rotl(Math.imul(k, C3), 17), C4);
const mixK4 = (k: number): number => Math.imul(rotl(Math.imul(k, C4), 18), C1);

const hex = (value: number): string =>
  (value >>> 0).toString(16).padStart(8, "0");

/** `text`'s 128-bit hash, as 32 hexadecimal digits. */
export const contentHash128 = (text: string): string => {
  let h1 = 0;
  let h2 = 0;
  let h3 = 0;
  let h4 = 0;
  const length = text.length;
  // Eight code units make a 16-byte block of four words.
  const blocks = length - (length % 8);
  for (let at = 0; at < blocks; at += 8) {
    h1 ^= mixK1(word(text, at));
    h1 = (Math.imul(rotl(h1, 19) + h2, 5) + 0x561ccd1b) | 0;
    h2 ^= mixK2(word(text, at + 2));
    h2 = (Math.imul(rotl(h2, 17) + h3, 5) + 0x0bcaa747) | 0;
    h3 ^= mixK3(word(text, at + 4));
    h3 = (Math.imul(rotl(h3, 15) + h4, 5) + 0x96cd1c35) | 0;
    h4 ^= mixK4(word(text, at + 6));
    h4 = (Math.imul(rotl(h4, 13) + h1, 5) + 0x32ac3b17) | 0;
  }
  const tail = length - blocks;
  if (tail > 6) h4 ^= mixK4(word(text, blocks + 6));
  if (tail > 4) h3 ^= mixK3(word(text, blocks + 4));
  if (tail > 2) h2 ^= mixK2(word(text, blocks + 2));
  if (tail > 0) h1 ^= mixK1(word(text, blocks));

  const bytes = length * 2;
  h1 ^= bytes;
  h2 ^= bytes;
  h3 ^= bytes;
  h4 ^= bytes;
  h1 = (h1 + h2 + h3 + h4) | 0;
  h2 = (h2 + h1) | 0;
  h3 = (h3 + h1) | 0;
  h4 = (h4 + h1) | 0;
  h1 = fmix(h1);
  h2 = fmix(h2);
  h3 = fmix(h3);
  h4 = fmix(h4);
  h1 = (h1 + h2 + h3 + h4) | 0;
  h2 = (h2 + h1) | 0;
  h3 = (h3 + h1) | 0;
  h4 = (h4 + h1) | 0;
  return hex(h1) + hex(h2) + hex(h3) + hex(h4);
};
