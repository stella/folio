import { describe, expect, test } from "bun:test";

import { englishCardinal, englishOrdinal } from "../spelledNumbers";

describe("counters written out in English", () => {
  test("cardinals capitalise the first word and join tens and units by a hyphen", () => {
    expect(
      [1, 12, 20, 21, 99, 100, 101, 1000, 1999, 12345].map(englishCardinal),
    ).toEqual([
      "One",
      "Twelve",
      "Twenty",
      "Twenty-one",
      "Ninety-nine",
      "One hundred",
      "One hundred one",
      "One thousand",
      "One thousand nine hundred ninety-nine",
      "Twelve thousand three hundred forty-five",
    ]);
  });

  test("ordinals change the last word only", () => {
    expect(
      [1, 2, 3, 5, 8, 9, 12, 20, 21, 40, 100, 101, 1000, 1999, 12348].map(
        englishOrdinal,
      ),
    ).toEqual([
      "First",
      "Second",
      "Third",
      "Fifth",
      "Eighth",
      "Ninth",
      "Twelfth",
      "Twentieth",
      "Twenty-first",
      "Fortieth",
      "One hundredth",
      "One hundred first",
      "One thousandth",
      "One thousand nine hundred ninety-ninth",
      "Twelve thousand three hundred forty-eighth",
    ]);
  });

  test("a counter outside the range written out is not known", () => {
    expect(englishCardinal(0)).toBeNull();
    expect(englishOrdinal(1_000_000)).toBeNull();
    expect(englishOrdinal(2.5)).toBeNull();
  });
});
