/**
 * Text normalisation shared by every extractor and the comparator. Both sides
 * of the diff must normalise identically or alignment falls apart.
 */

/** NFC-normalise, drop soft hyphens and zero-width characters, fold all
 * whitespace (incl. NBSP variants and tabs) to single spaces, trim. */
export const normalizeLineText = (text: string): string =>
  text
    .normalize("NFC")
    .replace(/[­​-‍﻿]/gu, "")
    .replace(/\s+/gu, " ")
    .trim();

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
