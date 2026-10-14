import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../test/property-testing";
import {
  base64ToBytes,
  bytesToBase64Url,
  createTextReader,
  decodeText,
  encodeBase64Portable,
  encodeBinaryString,
  encodeUtf16BE,
  encodeUtf16LE,
  encodeUtf8,
} from "./encoding";

const scalar = fc.oneof(
  fc.integer({ min: 0, max: 0xd7ff }),
  fc.integer({ min: 0xe000, max: 0x10ffff }),
);

test("byte transport preserves a leading BOM when requested", () => {
  const text = "\uFEFF👩🏽‍⚖️e\u0301שלום";
  const bytes = encodeUtf8(text);
  const reader = createTextReader({ ignoreBOM: true, fatal: true });
  const pieces: string[] = [];
  for (const byte of bytes) pieces.push(reader.read(Uint8Array.of(byte)));
  expect(pieces.join("") + reader.finish()).toBe(text);
  expect(decodeText(bytes, { ignoreBOM: true, fatal: true })).toBe(text);
});

test(
  "stream prefixes retain complete scalars and omit incomplete tails",
  async () => {
    await assertProperty(
      fc.property(
        fc.array(scalar, { maxLength: 30 }),
        fc.nat({ max: 150 }),
        fc.integer({ min: 1, max: 7 }),
        (points, cap, chunkSize) => {
          const text = points.map((point) => String.fromCodePoint(point)).join("");
          const bytes = encodeUtf8(text);
          const prefix = bytes.subarray(0, cap);
          const reader = createTextReader({ fatal: true, ignoreBOM: true });
          const pieces: string[] = [];
          for (let offset = 0; offset < prefix.length; offset += chunkSize)
            pieces.push(reader.read(prefix.subarray(offset, offset + chunkSize)));
          const actual = pieces.join("") + reader.finishPrefix();
          let used = 0;
          let expected = "";
          for (const character of text) {
            used += encodeUtf8(character).length;
            if (used <= cap) expected += character;
          }
          expect(actual).toBe(expected);
          expect(
            decodeText(encodeUtf16BE(text), { encoding: "utf-16be", ignoreBOM: true, fatal: true }),
          ).toBe(text);
          expect(
            decodeText(encodeUtf16LE(text), { encoding: "utf-16le", ignoreBOM: true, fatal: true }),
          ).toBe(text);
        },
      ),
      { numRuns: 100, id: "stream prefixes retain complete Unicode scalars" },
    );
  },
  propertyTestTimeout(10_000),
);

test(
  "portable and URL-safe base64 preserve every byte value",
  async () => {
    await assertProperty(
      fc.property(fc.uint8Array({ maxLength: 4096 }), (bytes) => {
        expect(base64ToBytes(encodeBase64Portable(bytes))).toEqual(bytes);
        expect(base64ToBytes(bytesToBase64Url(bytes))).toEqual(bytes);
        expect(encodeBinaryString(decodeText(bytes, { encoding: "binary" }))).toEqual(bytes);
      }),
      { numRuns: 100, id: "encoding owner preserves binary payloads" },
    );
  },
  propertyTestTimeout(10_000),
);
