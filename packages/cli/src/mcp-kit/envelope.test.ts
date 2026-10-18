import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { closestNames } from "./envelope";

setDefaultTimeout(propertyTestTimeout(5_000));

const scalarEditDistance = (a: string, b: string): number => {
  const left = Array.from(a.toWellFormed());
  const right = Array.from(b.toWellFormed());
  let previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let i = 1; i <= left.length; i += 1) {
    const current = [i];
    for (let j = 1; j <= right.length; j += 1) {
      current[j] = Math.min(
        (previous[j] ?? 0) + 1,
        (current[j - 1] ?? 0) + 1,
        (previous[j - 1] ?? 0) + (left[i - 1] === right[j - 1] ? 0 : 1),
      );
    }
    previous = current;
  }
  return previous[right.length] ?? 0;
};

const oldClosestNames = (name: string, known: readonly string[]): string[] => {
  const lowered = name.toLowerCase();
  return known
    .map((candidate) => {
      const lower = candidate.toLowerCase();
      const contains = lower.includes(lowered) || lowered.includes(lower);
      return { candidate, distance: contains ? 0 : scalarEditDistance(lowered, lower) };
    })
    .filter(({ distance }) => distance <= Math.max(2, Math.floor(name.length / 3)))
    .toSorted((a, b) => a.distance - b.distance || a.candidate.localeCompare(b.candidate))
    .slice(0, 3)
    .map(({ candidate }) => candidate);
};

const normalizedClosestNames = (name: string, known: readonly string[]): string[] => {
  const lowered = name.toWellFormed().toLowerCase();
  return known
    .map((candidate) => {
      const lower = candidate.toWellFormed().toLowerCase();
      const contains = lower.includes(lowered) || lowered.includes(lower);
      return { candidate, distance: contains ? 0 : scalarEditDistance(lowered, lower) };
    })
    .filter(({ distance }) => distance <= Math.max(2, Math.floor(name.length / 3)))
    .toSorted((a, b) => a.distance - b.distance || a.candidate.localeCompare(b.candidate))
    .slice(0, 3)
    .map(({ candidate }) => candidate);
};

describe("closestNames", () => {
  test("keeps case-insensitive containment matches and sorts ties by locale", () => {
    expect(closestNames("SEARCH", ["research", "search", "Search", "searching"])).toEqual([
      "research",
      "search",
      "Search",
    ]);
  });

  test("keeps the existing UTF-16 name-length distance threshold and returns at most three", () => {
    expect(
      closestNames("serach", [
        "search",
        "sear",
        "speech",
        "searche",
        "x",
        "seraih",
        "seraqh",
        "serafh",
        "serach",
      ]),
    ).toEqual(["serach", "serafh", "seraih"]);
  });

  test("uses ordinary code-unit distance for ASCII and BMP text", () => {
    expect(closestNames("café", ["cafe", "caff", "cafés", "cake"])).toEqual([
      "cafés",
      "cafe",
      "caff",
    ]);
  });

  test("keeps empty queries, duplicate candidates, Czech inflections, and punctuation", () => {
    expect(closestNames("", [])).toEqual([]);
    expect(closestNames("", ["search", ""])).toEqual(["", "search"]);
    expect(closestNames("search", ["search", "search", "search"])).toEqual([
      "search",
      "search",
      "search",
    ]);
    expect(closestNames("Příbrami", ["Příbram", "Příbramy", "Praha", "§ 12"])).toEqual([
      "Příbram",
      "Příbramy",
    ]);
    expect(closestNames("čl. 12", ["čl. 12", "čl. 13", "čl. 2", "článek 12"])).toEqual([
      "čl. 12",
      "čl. 13",
      "čl. 2",
    ]);
  });

  test("matches the previous oracle over seeded ASCII and BMP strings and lists", () => {
    const text = fc
      .array(fc.constantFrom("a", "A", "č", "ř", "é", "ž", " ", "-", ".", "?", "§", "1"), {
        maxLength: 16,
      })
      .map((characters) => characters.join(""));
    fc.assert(
      fc.property(text, fc.array(text, { maxLength: 12 }), (name, known) => {
        expect(closestNames(name, known)).toEqual(oldClosestNames(name, known));
      }),
      propertyConfig({ seed: 1_725_983, numRuns: 500 }),
    );
  });

  test("counts supplementary-plane characters as one edit", () => {
    expect(closestNames("😀a", ["b", "😀b", "c"])).toEqual(["😀b", "b", "c"]);
  });

  test("normalizes lone high and low surrogates to the replacement character", () => {
    expect(closestNames("\uD800", ["\uFFFD"])).toEqual(["\uFFFD"]);
    expect(closestNames("\uDC00", ["\uFFFD"])).toEqual(["\uFFFD"]);
    expect(closestNames("\uFFFD", ["\uD800", "\uDC00"])).toEqual(["\uD800", "\uDC00"]);
  });

  test("normalizes malformed surrogates embedded in names and candidates", () => {
    expect(closestNames("a\uD800b", ["a\uFFFDb", "a\uDFFFc", "a\uFFFDd"])).toEqual([
      "a\uFFFDb",
      "a\uDFFFc",
      "a\uFFFDd",
    ]);
    expect(closestNames("a\uFFFDb", ["a\uD800b"])).toEqual(["a\uD800b"]);
  });

  test("matches a seeded normalized oracle over malformed and well-formed Unicode", () => {
    const text = fc
      .array(fc.constantFrom("a", "c", "č", "\uD800", "\uDC00", "😀", "\uFFFD"), {
        maxLength: 10,
      })
      .map((characters) => characters.join(""));
    fc.assert(
      fc.property(text, fc.array(text, { maxLength: 10 }), (name, known) => {
        expect(closestNames(name, known)).toEqual(normalizedClosestNames(name, known));
      }),
      propertyConfig({ seed: 1_725_984, numRuns: 500 }),
    );
  });
});
