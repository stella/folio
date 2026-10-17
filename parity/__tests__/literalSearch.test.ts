import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../test/property-testing";

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

  test(
    "matches the native literal oracle for generated scalar strings",
    () => {
      const scalar = fc.constantFrom("a", "b", "é", "e\u0301", "😀", "𞤀", "+", "[");
      const text = fc.array(scalar, { maxLength: 40 }).map((parts) => parts.join(""));
      const pattern = fc.array(scalar, { maxLength: 12 }).map((parts) => parts.join(""));
      const candidates = fc.array(text, { maxLength: 12 });
      assertProperty(
        fc.property(pattern, text, candidates, (needle, haystack, values) => {
          const normalizedNeedle = needle.toWellFormed();
          const normalizedHaystack = haystack.toWellFormed();
          const normalizedCandidates = values.map((value) => value.toWellFormed());
          expect(createLiteralSearch(needle)(haystack)).toBe(
            normalizedHaystack.includes(normalizedNeedle),
          );
          expect(createCandidateContainment(needle, [...values, haystack])(haystack)).toBe(
            normalizedHaystack.includes(normalizedNeedle) ||
              normalizedNeedle.includes(normalizedHaystack),
          );
          expect(createCandidateContainment(needle, values)(haystack)).toBe(
            normalizedHaystack.includes(normalizedNeedle) ||
              ((normalizedCandidates.includes(normalizedHaystack) || normalizedHaystack === "") &&
                normalizedNeedle.includes(normalizedHaystack)),
          );
        }),
        { id: "literal search matches native scalar-string containment", numRuns: 100 },
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "normalizes generated lone surrogates before matching",
    () => {
      const unit = fc.constantFrom("a", "😀", "\ud800", "\udc00", "\ufffd");
      const text = fc.array(unit, { maxLength: 30 }).map((parts) => parts.join(""));
      assertProperty(
        fc.property(text, text, (needle, haystack) => {
          expect(createLiteralSearch(needle)(haystack)).toBe(
            haystack.toWellFormed().includes(needle.toWellFormed()),
          );
        }),
        { id: "literal search normalizes malformed UTF-16 at the boundary", numRuns: 100 },
      );
    },
    propertyTestTimeout(10_000),
  );
});
