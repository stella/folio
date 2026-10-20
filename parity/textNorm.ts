import { distance } from "@stll/fuzzy-search";

import { detectBaseDirection } from "../packages/core/src/utils/baseDirection";

/**
 * Text normalisation shared by every extractor and the comparator. Both sides
 * of the diff must normalise identically or alignment falls apart.
 */

const RTL_LEADER_LINE_PATTERN = /^(?<page>[\p{Decimal_Number}]+)\s+…\s+(?<title>.+)$/u;

const MIRRORABLE_BRACKET_PATTERN = /[()[\]{}]/u;
const MIRRORABLE_BRACKETS_PATTERN = /[()[\]{}]/gu;
const BRACKET_PAIR_REPRESENTATIVE: Record<string, string> = {
  "(": "(",
  ")": "(",
  "[": "[",
  "]": "[",
  "{": "{",
  "}": "{",
};

/**
 * A PDF records the mirrored glyph an RTL line actually paints, while a DOM
 * carries the logical character, so one extractor reports `)…(` where the other
 * reports `(…)`. Neither text model observes painted mirroring (the browser
 * mirrors at paint time), so comparing bracket direction on an RTL line only
 * manufactures text differences. Fold each pair to one representative instead;
 * LTR lines keep both characters distinct.
 */
const canonicalizeRtlBracketPairs = (text: string): string => {
  if (!MIRRORABLE_BRACKET_PATTERN.test(text) || detectBaseDirection(text) !== "rtl") {
    return text;
  }
  return text.replace(
    MIRRORABLE_BRACKETS_PATTERN,
    (bracket) => BRACKET_PAIR_REPRESENTATIVE[bracket] ?? bracket,
  );
};

const normalizeRtlLeaderOrder = (text: string): string => {
  const match = text.match(RTL_LEADER_LINE_PATTERN);
  const page = match?.groups?.["page"];
  const title = match?.groups?.["title"];
  if (!page || !title || detectBaseDirection(title) !== "rtl") {
    return text;
  }
  return `${title} … ${page}`;
};

/** NFC-normalise, fold known PDF glyph aliases, drop soft hyphens and
 * zero-width characters, fold all whitespace (incl. NBSP variants and tabs)
 * to single spaces, trim. */
export const normalizeLineText = (text: string): string => {
  const normalized = text
    .normalize("NFC")
    // Some CJK PDF font maps expose an ordinary ideograph as the equivalent
    // CJK or Kangxi radical (for example 乙 as U+2F04). The visual glyph and
    // line endpoint are unchanged, so canonicalize only those radical blocks.
    .replace(/[⺀-⿕]/gu, (character) => character.normalize("NFKC"))
    .replace(/[­​-‍﻿]/gu, "")
    .replace(/\uf0b7/gu, "•")
    .replace(/\uf0a8/gu, "☐")
    .replace(/\uf0e3/gu, "ã")
    // A PDF's ToUnicode map answers shaped Arabic glyphs with the Persian code
    // points (Farsi Yeh, Heh Doachashmee) even for documents that authored the
    // Arabic letters, so extraction spelling would read as a text difference.
    .replace(/\u06cc/gu, "\u064a")
    .replace(/\u06be/gu, "\u0647")
    .replace(/\s*\.{3,}\s*/gu, " … ")
    .replace(/(?:\s*…\s*){2,}/gu, " … ")
    .replace(/\s+/gu, " ")
    .trim();
  return canonicalizeRtlBracketPairs(normalizeRtlLeaderOrder(normalized));
};

const scalarLength = (text: string): number => {
  let length = 0;
  for (const _character of text) length += 1;
  return length;
};

/** Levenshtein similarity in [0, 1], with edits and lengths in Unicode scalars. */
export const textSimilarity = (a: string, b: string): number => {
  const left = a.toWellFormed();
  const right = b.toWellFormed();
  if (left === right) return 1;
  const maxLen = Math.max(scalarLength(left), scalarLength(right));
  if (maxLen === 0) return 1;
  return 1 - distance(left, right, "levenshtein") / maxLen;
};
