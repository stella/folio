import { describe, expect, test } from "bun:test";

import { createCandidateContainment, createLiteralSearch } from "../literalSearch";
import { normalizeLineText } from "../textNorm";

describe("parity literal queries", () => {
  const fragments = [
    "",
    "a",
    "aa",
    "a+b[0]",
    "Novák",
    "Nováka",
    "Novak",
    "novák",
    "😀",
    "é",
    "e\u0301",
    "\ud800",
    "\udc00",
    "\ufffd",
    "a\ud800b",
  ];

  test("keeps exact literals, empty queries, overlaps and both containment directions", () => {
    for (const needle of fragments) {
      const search = createLiteralSearch(needle);
      const candidates = [...fragments, ...fragments.map((part) => needle + part)];
      const contains = createCandidateContainment(needle, candidates);
      for (const text of candidates) {
        const normalizedText = text.toWellFormed();
        const normalizedNeedle = needle.toWellFormed();
        expect(search(text)).toBe(normalizedText.includes(normalizedNeedle));
        expect(contains(text)).toBe(
          normalizedText.includes(normalizedNeedle) || normalizedNeedle.includes(normalizedText),
        );
      }
    }
  });

  test("uses caller normalization without adding accent or case folding", () => {
    const query = normalizeLineText("Nováka\u00a0 ve smlouvě").toLocaleLowerCase();
    const search = createLiteralSearch(query);
    expect(search(normalizeLineText("NOVÁKA ve\t smlouvě!").toLocaleLowerCase())).toBe(true);
    expect(search("novaka ve smlouve")).toBe(false);
    expect(search("Nováka ve smlouvě")).toBe(false);
    expect(createLiteralSearch("a+b[0]")("a+b[0]!")).toBe(true);
    expect(createLiteralSearch("a+b[0]")("aaab0")).toBe(false);
  });
});
