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

/** Levenshtein-based similarity in [0, 1]; 1 means identical. */
export const textSimilarity = (a: string, b: string): number => {
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
};

const levenshtein = (a: string, b: string): number => {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  // Two-row DP over Int32Array to avoid boxed-number arrays: this runs inside
  // alignFull's O(n²) loop (up to millions of calls), so per-call allocation
  // and GC pressure dominate. Typed arrays reuse a flat backing buffer.
  let prev = new Int32Array(b.length + 1);
  let curr = new Int32Array(b.length + 1);
  for (let i = 0; i <= b.length; i++) {
    prev[i] = i;
  }
  for (let i = 1; i <= a.length; i++) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min((curr[j - 1] ?? 0) + 1, (prev[j] ?? 0) + 1, (prev[j - 1] ?? 0) + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[b.length] ?? 0;
};
