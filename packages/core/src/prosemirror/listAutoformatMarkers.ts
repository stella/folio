/**
 * The list markers that automatic bulleted and numbered lists convert when
 * typed at the start of a paragraph, and the list each one asks for.
 *
 * A marker's punctuation becomes the level text (`1.` → `%1.`, `1)` → `%1)`),
 * its letter case and numeral system the number format, and its value the
 * value the new list starts at. What is not a marker stays text: a number of
 * four or more digits (a year: `2024. `), zero, a word that happens to be
 * spelled in roman digits (`mix. `, `cm. `), and anything that is not the
 * whole text before the caret.
 */

import type { NumberFormat } from "../types/document";
import type { ListRequest } from "./listNumbering";

/**
 * A marker plus the space that triggers the conversion. The text before the
 * caret must be exactly this, so a marker typed mid-paragraph never matches.
 */
export const LIST_MARKER_PATTERN = /^(?:[-*]|\d{1,3}[.)]|[A-Za-z]+[.)]) $/u;

const BULLET_PATTERN = /^[-*] $/u;
const NUMBER_PATTERN = /^(?<value>\d{1,3})(?<punctuation>[.)]) $/u;
const LETTERS_PATTERN = /^(?<letters>[a-z]+|[A-Z]+)(?<punctuation>[.)]) $/u;

/**
 * The largest roman numeral read as a list marker. Past it the letters spell
 * words (`mix`, `civil`, `dim`) far more often than list items.
 */
const MAX_ROMAN_MARKER = 50;

const ROMAN_VALUES = [
  ["m", 1000],
  ["cm", 900],
  ["d", 500],
  ["cd", 400],
  ["c", 100],
  ["xc", 90],
  ["l", 50],
  ["xl", 40],
  ["x", 10],
  ["ix", 9],
  ["v", 5],
  ["iv", 4],
  ["i", 1],
] as const;

const toRoman = (value: number): string => {
  let remaining = value;
  let roman = "";
  for (const [digits, digitValue] of ROMAN_VALUES) {
    while (remaining >= digitValue) {
      roman += digits;
      remaining -= digitValue;
    }
  }
  return roman;
};

const ROMAN_DIGIT_VALUES: Readonly<Record<string, number>> = {
  i: 1,
  v: 5,
  x: 10,
  l: 50,
  c: 100,
  d: 500,
  m: 1000,
};

/** The value of a canonically spelled roman numeral, or `undefined`. */
const romanValue = (letters: string): number | undefined => {
  const lower = letters.toLowerCase();
  let total = 0;
  for (let index = 0; index < lower.length; index += 1) {
    const value = ROMAN_DIGIT_VALUES[lower.charAt(index)];
    if (value === undefined) {
      return undefined;
    }
    const next = ROMAN_DIGIT_VALUES[lower.charAt(index + 1)] ?? 0;
    total += value < next ? -value : value;
  }
  return total > 0 && toRoman(total) === lower ? total : undefined;
};

const numbered = (numFmt: NumberFormat, punctuation: string, start: number): ListRequest => ({
  kind: "numbered",
  format: { numFmt, lvlText: `%1${punctuation}` },
  start,
});

const letterFormat = (letters: string): NumberFormat =>
  letters === letters.toUpperCase() ? "upperLetter" : "lowerLetter";

const romanFormat = (letters: string): NumberFormat =>
  letters === letters.toUpperCase() ? "upperRoman" : "lowerRoman";

/**
 * The lists `text` (the paragraph's text up to and including the typed space)
 * may ask for, most likely first, or `null` when it is not a marker.
 *
 * A single letter that is also a roman digit is ambiguous: `i.` usually
 * starts a roman list, but directly under an `h.` item it continues the
 * letters. The caller prefers whichever interpretation a neighbouring list
 * already has, and otherwise takes the first.
 */
export const listRequestsForMarker = (text: string): ListRequest[] | null => {
  if (BULLET_PATTERN.test(text)) {
    return [{ kind: "bullet" }];
  }
  const number = NUMBER_PATTERN.exec(text)?.groups;
  if (number) {
    const value = Number.parseInt(number["value"] ?? "", 10);
    return value >= 1 ? [numbered("decimal", number["punctuation"] ?? ".", value)] : null;
  }
  const lettered = LETTERS_PATTERN.exec(text)?.groups;
  if (!lettered) {
    return null;
  }
  const letters = lettered["letters"] ?? "";
  const punctuation = lettered["punctuation"] ?? ".";
  const roman = romanValue(letters);
  const asRoman =
    roman !== undefined && roman <= MAX_ROMAN_MARKER
      ? numbered(romanFormat(letters), punctuation, roman)
      : undefined;
  if (letters.length > 1) {
    return asRoman ? [asRoman] : null;
  }
  const asLetter = numbered(
    letterFormat(letters),
    punctuation,
    letters.toLowerCase().charCodeAt(0) - "a".charCodeAt(0) + 1,
  );
  if (!asRoman) {
    return [asLetter];
  }
  return roman === 1 ? [asRoman, asLetter] : [asLetter, asRoman];
};
