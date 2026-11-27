/**
 * Counters a list writes out in words (`cardinalText`, `ordinalText`), as
 * the reference consumer writes them in English: the first letter
 * capitalised, tens and units joined by a hyphen, no "and" (`One hundred
 * one`, `Twenty-first`, `One thousand nine hundred ninety-ninth`).
 */

const UNITS = [
  "",
  "one",
  "two",
  "three",
  "four",
  "five",
  "six",
  "seven",
  "eight",
  "nine",
  "ten",
  "eleven",
  "twelve",
  "thirteen",
  "fourteen",
  "fifteen",
  "sixteen",
  "seventeen",
  "eighteen",
  "nineteen",
] as const;

const TENS = [
  "",
  "",
  "twenty",
  "thirty",
  "forty",
  "fifty",
  "sixty",
  "seventy",
  "eighty",
  "ninety",
] as const;

/** Ordinals that do not add `th` to the cardinal. */
const IRREGULAR_ORDINALS: Readonly<Record<string, string>> = {
  one: "first",
  two: "second",
  three: "third",
  five: "fifth",
  eight: "eighth",
  nine: "ninth",
  twelve: "twelfth",
};

/** The largest counter written out here; a larger one is not known. */
const SPELLED_MAX = 999_999;

const belowHundred = (value: number): string => {
  if (value < 20) {
    return UNITS[value] ?? "";
  }
  const unit = value % 10;
  const tens = TENS[Math.floor(value / 10)] ?? "";
  return unit === 0 ? tens : `${tens}-${UNITS[unit] ?? ""}`;
};

const belowThousand = (value: number): string => {
  const hundreds = Math.floor(value / 100);
  const rest = value % 100;
  const words: string[] = [];
  if (hundreds > 0) {
    words.push(`${UNITS[hundreds] ?? ""} hundred`);
  }
  if (rest > 0) {
    words.push(belowHundred(rest));
  }
  return words.join(" ");
};

const lowerCardinal = (value: number): string => {
  const thousands = Math.floor(value / 1000);
  const rest = value % 1000;
  const words: string[] = [];
  if (thousands > 0) {
    words.push(`${belowThousand(thousands)} thousand`);
  }
  if (rest > 0) {
    words.push(belowThousand(rest));
  }
  return words.join(" ");
};

const capitalised = (words: string): string =>
  words.charAt(0).toUpperCase() + words.slice(1);

const ordinalWord = (word: string): string =>
  IRREGULAR_ORDINALS[word] ??
  (word.endsWith("y") ? `${word.slice(0, -1)}ieth` : `${word}th`);

const known = (value: number): boolean =>
  Number.isInteger(value) && value >= 1 && value <= SPELLED_MAX;

/** `cardinalText` in English, or null for a counter not written out here. */
export const englishCardinal = (value: number): string | null =>
  known(value) ? capitalised(lowerCardinal(value)) : null;

/** `ordinalText` in English, or null for a counter not written out here. */
export const englishOrdinal = (value: number): string | null => {
  if (!known(value)) {
    return null;
  }
  const cardinal = lowerCardinal(value);
  // The last word, or the unit after the hyphen, takes the ordinal form.
  const cut = Math.max(cardinal.lastIndexOf(" "), cardinal.lastIndexOf("-"));
  return capitalised(
    cardinal.slice(0, cut + 1) + ordinalWord(cardinal.slice(cut + 1)),
  );
};
