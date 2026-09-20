/**
 * The guard fires on the idiom it replaced, and on nothing else.
 *
 * A check nobody has seen fail is a check nobody knows the shape of, and this
 * one draws a line between two reads of the same element — the tolerance the
 * table now owns — and a `w:start` that is simply a different slot.
 */

import { describe, expect, test } from "bun:test";

import { handListedSpellings } from "./check-strict-name-tolerance";

describe("a reader that hand-lists a rename", () => {
  test("two spellings off one element is a finding", () => {
    expect(
      handListedSpellings(
        `const left = parseBorderSpec(
           findChild(bordersElement, "w", "left") ?? findChild(bordersElement, "w", "start"),
         );`,
        "tableParser.ts",
      ),
    ).toEqual([
      "tableParser.ts:2 reads w:left and w:start off bordersElement; the rename belongs to strictNames.gen.ts, so read it with findChildAnySpelling/numericAttributeAnySpelling/hasAttributeAnySpelling.",
    ]);
  });

  test("statements apart still count, because the drift does not care", () => {
    expect(
      handListedSpellings(
        `const left = parseNumericAttribute(ind, "w", "left");
         const firstLine = parseNumericAttribute(ind, "w", "firstLine");
         const start = parseNumericAttribute(ind, "w", "start");`,
      ),
    ).toHaveLength(1);
  });
});

describe("a name that is not a rename", () => {
  test("a list's first number reads alone", () => {
    expect(handListedSpellings(`const start = parseNumericAttribute(lvl, "w", "start");`)).toEqual(
      [],
    );
  });

  test("two elements are two slots, not one pair", () => {
    expect(
      handListedSpellings(
        `const margin = parseNumericAttribute(pgMar, "w", "left");
         const origin = parseNumericAttribute(lnNumType, "w", "start");`,
      ),
    ).toEqual([]);
  });

  test("the table's own readers are not a hand-list", () => {
    expect(
      handListedSpellings(
        `for (const spelling of spellingsOf(slot)) {
           const child = findChild(parent, "w", spelling);
         }`,
      ),
    ).toEqual([]);
  });
});
