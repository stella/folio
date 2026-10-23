import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { sliceText } from "./sliceText";

const scalar = fc.oneof(
  fc.integer({ min: 0, max: 0xd7ff }),
  fc.integer({ min: 0xe000, max: 0x10ffff }),
);
const unicodeText = fc
  .array(scalar, { maxLength: 60 })
  .map((points) => points.map((point) => String.fromCodePoint(point)).join(""));
const normalizeOffset = (offset: number, length: number): number =>
  offset < 0 ? Math.max(0, length + offset) : Math.min(length, offset);

test(
  "code-unit windows retain exactly their wholly contained Unicode characters",
  async () => {
    await assertProperty(
      fc.property(
        unicodeText,
        fc.integer({ min: -150, max: 150 }),
        fc.integer({ min: -150, max: 150 }),
        (text, start, end) => {
          const from = normalizeOffset(start, text.length);
          const to = normalizeOffset(end, text.length);
          let position = 0;
          let expected = "";
          for (const character of text) {
            const next = position + character.length;
            if (position >= from && next <= to) expected += character;
            position = next;
          }
          const result = sliceText({ text, units: "codeUnits", start, end });
          expect(result).toBe(expected);
          expect(/[\uD800-\uDFFF]/u.test(result)).toBe(false);
          expect(result.length).toBeLessThanOrEqual(Math.max(0, to - from));
          expect(sliceText({ text: result, units: "codeUnits" })).toBe(result);
        },
      ),
      {
        numRuns: 200,
        id: "code-unit windows retain complete Unicode characters",
      },
    );
  },
  propertyTestTimeout(10_000),
);

const grapheme = fc.constantFrom(
  "a",
  "e\u0301",
  "\u{1f1e8}\u{1f1ff}",
  "👨‍👩‍👧",
  "👩🏽‍⚖️",
);

test(
  "initials keep flags, joined emoji and combining marks whole",
  async () => {
    await assertProperty(
      fc.property(
        fc.array(grapheme, { minLength: 1, maxLength: 20 }),
        fc.nat({ max: 25 }),
        (characters, boundary) => {
          const text = characters.join("");
          const before = sliceText({ text, units: "graphemes", end: boundary });
          const after = sliceText({
            text,
            units: "graphemes",
            start: boundary,
          });
          expect(before).toBe(characters.slice(0, boundary).join(""));
          expect(after).toBe(characters.slice(boundary).join(""));
          expect(before + after).toBe(text);
        },
      ),
      { numRuns: 200, id: "initials retain whole grapheme clusters" },
    );
  },
  propertyTestTimeout(10_000),
);

test(
  "code-point partitions preserve text and count supplementary characters once",
  async () => {
    await assertProperty(
      fc.property(unicodeText, fc.nat({ max: 100 }), (text, boundary) => {
        const before = sliceText({ text, units: "codePoints", end: boundary });
        const after = sliceText({ text, units: "codePoints", start: boundary });
        expect(before + after).toBe(text);
        expect(/[\uD800-\uDFFF]/u.test(before)).toBe(false);
        expect(/[\uD800-\uDFFF]/u.test(after)).toBe(false);
        expect(Array.from(before).length).toBe(
          Math.min(boundary, Array.from(text).length),
        );
      }),
      { numRuns: 200, id: "code-point partitions preserve Unicode text" },
    );
  },
  propertyTestTimeout(10_000),
);
