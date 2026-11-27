import { expect, test } from "bun:test";

import { contentHash128 } from "./contentHash";

test("a string's hash is 32 hexadecimal digits, the same each time", () => {
  for (const text of ["", "a", "paragraph", "é€𝒜", "x".repeat(1000)]) {
    expect(contentHash128(text)).toMatch(/^[0-9a-f]{32}$/u);
    expect(contentHash128(text)).toBe(contentHash128(text));
  }
});

test("strings that differ in one code unit, or in length alone, hash apart", () => {
  const texts = new Set<string>();
  const base = "The Supplier shall deliver the goods within ten days.";
  for (let at = 0; at < base.length; at += 1)
    for (let code = 0; code < 512; code += 1)
      texts.add(
        `${base.slice(0, at)}${String.fromCharCode(code)}${base.slice(at + 1)}`,
      );
  // Every tail length, zero padding included.
  for (let length = 0; length <= 64; length += 1) {
    texts.add("\u0000".repeat(length));
    texts.add("a".repeat(length));
  }
  // Each of the four words of a block reaches the hash.
  for (let at = 0; at < 16; at += 1)
    texts.add(`${"b".repeat(at)}c${"b".repeat(15 - at)}`);
  const hashes = new Set([...texts].map(contentHash128));
  expect(hashes.size).toBe(texts.size);
});
