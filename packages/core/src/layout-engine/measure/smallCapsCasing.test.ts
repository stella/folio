import { describe, expect, test } from "bun:test";

import {
  SMALL_CAPS_SCALE,
  isSmallCapsLetter,
  smallCapsMask,
  smallCapsSegments,
} from "./smallCapsCasing";

describe("isSmallCapsLetter", () => {
  test("is true for a cased lowercase letter", () => {
    expect(isSmallCapsLetter("a")).toBe(true);
    expect(isSmallCapsLetter("z")).toBe(true);
  });

  test("is false for an already-uppercase letter", () => {
    expect(isSmallCapsLetter("A")).toBe(false);
  });

  test("is false for digits and punctuation, which carry no case", () => {
    expect(isSmallCapsLetter("5")).toBe(false);
    expect(isSmallCapsLetter("/")).toBe(false);
    expect(isSmallCapsLetter(".")).toBe(false);
    expect(isSmallCapsLetter("(")).toBe(false);
  });

  test("is false for a caseless script (CJK has no upper/lower distinction)", () => {
    expect(isSmallCapsLetter("日")).toBe(false);
  });
});

describe("smallCapsMask", () => {
  test("marks only the lowercase letters small", () => {
    expect(smallCapsMask("Abc")).toEqual([false, true, true]);
  });

  test("a punctuation mark between two small-caps words is not small", () => {
    // "Ok/No" — the "/" itself carries no case and does not inherit the small
    // class from either neighbour: a "/" between two synthesized words
    // measures at the full run size.
    expect(smallCapsMask("Ok/No")).toEqual([
      false, // O
      true, // k
      false, // /
      false, // N
      true, // o
    ]);
  });

  test("a trailing space after a small-caps word takes its class", () => {
    // "Also On" — the space after "lso" (small) stays small even though the
    // next letter, "O", is a full-size capital.
    const mask = smallCapsMask("Also On");
    expect(mask[3]).toBe(true); // the trailing "o" of "lso"
    expect(mask[4]).toBe(true); // the space: inherits "o"'s class
    expect(mask[5]).toBe(false); // "O": a source capital
  });

  test("a leading space with nothing before it defaults to full size", () => {
    expect(smallCapsMask(" a")).toEqual([false, true]);
  });
});

describe("smallCapsSegments", () => {
  test("returns one full-size segment when nothing is lowercase", () => {
    expect(smallCapsSegments("HEADING")).toEqual([{ text: "HEADING", small: false }]);
  });

  test("returns one full-size segment for text with no case at all", () => {
    expect(smallCapsSegments("123 / 456")).toEqual([{ text: "123 / 456", small: false }]);
  });

  test("splits mixed-case text into alternating segments, uppercasing the small ones", () => {
    expect(smallCapsSegments("Also On Duty")).toEqual([
      { text: "A", small: false },
      { text: "LSO ", small: true },
      { text: "O", small: false },
      { text: "N ", small: true },
      { text: "D", small: false },
      { text: "UTY", small: true },
    ]);
  });

  test("keeps a punctuation mark between two words in the full-size class", () => {
    expect(smallCapsSegments("alpha / bravo")).toEqual([
      { text: "ALPHA ", small: true },
      { text: "/ ", small: false },
      { text: "BRAVO", small: true },
    ]);
  });
});

test("SMALL_CAPS_SCALE is the authored ratio, not a browser's synthesis", () => {
  // A 12pt run's synthesized capitals draw at exactly 0.8 of that (9.6pt) —
  // not Blink/WebKit's own small-caps synthesis multiplier (0.7).
  expect(SMALL_CAPS_SCALE).toBe(0.8);
});
