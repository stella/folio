import { describe, expect, spyOn, test } from "bun:test";
import { AhoCorasick } from "@stll/aho-corasick";
import fc from "fast-check";
import ts from "typescript";
import { assertProperty, propertyTestTimeout } from "../../test/property-testing";

import { createLiteralSearchScope } from "../literalSearch";
import { normalizeLineText } from "../textNorm";

describe("parity literal queries", () => {
  test("the matcher module owns no mutable module state", async () => {
    const source = ts.createSourceFile(
      "literalSearch.ts",
      await Bun.file(new URL("../literalSearch.ts", import.meta.url)).text(),
      ts.ScriptTarget.Latest,
      true,
    );
    for (const statement of source.statements) {
      if (ts.isImportDeclaration(statement)) continue;
      expect(ts.isVariableStatement(statement)).toBe(true);
      if (!ts.isVariableStatement(statement)) continue;
      expect(statement.declarationList.flags & ts.NodeFlags.Const).toBeTruthy();
      for (const declaration of statement.declarationList.declarations) {
        expect(declaration.initializer && ts.isArrowFunction(declaration.initializer)).toBe(true);
      }
    }
  });

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
    const { createLiteralSearch, createCandidateContainment } = createLiteralSearchScope();
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
    const { createLiteralSearch } = createLiteralSearchScope();
    const query = normalizeLineText("Nováka\u00a0 ve smlouvě").toLocaleLowerCase();
    const search = createLiteralSearch(query);
    expect(search(normalizeLineText("NOVÁKA ve\t smlouvě!").toLocaleLowerCase())).toBe(true);
    expect(search("novaka ve smlouve")).toBe(false);
    expect(search("Nováka ve smlouvě")).toBe(false);
    expect(createLiteralSearch("a+b[0]")("a+b[0]!")).toBe(true);
    expect(createLiteralSearch("a+b[0]")("aaab0")).toBe(false);
  });

  test("candidate membership never enumerates repetitive occurrences", () => {
    const overlapping = spyOn(AhoCorasick.prototype, "findOverlappingIter");
    const nonOverlapping = spyOn(AhoCorasick.prototype, "findIter");
    try {
      const { createCandidateContainment } = createLiteralSearchScope();
      const query = "\ufffd".repeat(48);
      const candidates = ["\ud800", "\udc00", "\ufffd", ""].flatMap((unit) =>
        Array.from({ length: 12 }, () => unit.repeat(6)),
      );
      const contains = createCandidateContainment(query, candidates);
      for (const candidate of candidates) expect(contains(candidate)).toBe(true);
      expect(contains("\ufffd")).toBe(false);
      expect(contains("not a candidate")).toBe(false);
      expect(overlapping).not.toHaveBeenCalled();
      expect(nonOverlapping).not.toHaveBeenCalled();
    } finally {
      overlapping.mockRestore();
      nonOverlapping.mockRestore();
    }
  });

  test(
    "membership is independent of duplicate candidates and occurrence counts",
    () => {
      const unit = fc.constantFrom("a", "😀", "é", "e\u0301", "\ud800", "\udc00", "\ufffd");
      assertProperty(
        fc.property(
          unit,
          fc.integer({ min: 1, max: 40 }),
          fc.array(fc.integer({ min: 1, max: 20 }), { minLength: 1, maxLength: 12 }),
          fc.integer({ min: 1, max: 30 }),
          (fragment, occurrences, widths, duplicates) => {
            const query = fragment.repeat(occurrences);
            const candidates = widths.flatMap((width) =>
              Array.from({ length: duplicates }, () => fragment.repeat(width)),
            );
            const scope = createLiteralSearchScope();
            const contains = scope.createCandidateContainment(query, candidates);
            const wellFormedQuery = query.toWellFormed();
            for (const candidate of candidates) {
              const normalized = candidate.toWellFormed();
              expect(contains(candidate)).toBe(
                normalized.includes(wellFormedQuery) || wellFormedQuery.includes(normalized),
              );
            }
          },
        ),
        { id: "candidate membership ignores duplicate and repetitive occurrences", numRuns: 100 },
      );
    },
    propertyTestTimeout(10_000),
  );

  test(
    "matches the native literal oracle for generated scalar strings",
    () => {
      const { createLiteralSearch, createCandidateContainment } = createLiteralSearchScope();
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
      const { createLiteralSearch } = createLiteralSearchScope();
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
