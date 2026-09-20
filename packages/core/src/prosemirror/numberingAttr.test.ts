/**
 * The attr codec is the one crossing between the model's numbering union and
 * the persisted `numPr` attr. Before it existed the two tiers traded values
 * directly and each read the other's shape as its own: `toProseDoc` stored a
 * model `ParagraphFormatting` verbatim in `_originalFormatting` while the attr
 * beside it held the two `<w:numPr>` slots, and `paragraphRejectOriginalFormatting`
 * copied the slots back into that model-typed field — where
 * `sameStatedParagraphNumbering` switches on `kind` and a slot pair has none,
 * so an untouched save started materialising numbering as direct formatting.
 */

import { describe, expect, test } from "bun:test";
import { Panic } from "better-result";

import { paragraphRejectOriginalFormatting } from "./commands/propertyChangeScope";
import {
  paragraphNumberingAttr,
  paragraphNumberingFromAttrValue,
  readParagraphNumberingAttr,
} from "./numberingAttr";

describe("paragraphNumberingFromAttrValue", () => {
  test.each([
    ["a cancellation", { kind: "none" }, { kind: "none" }],
    [
      "a reference with a level",
      { kind: "reference", numId: 4, ilvl: 2 },
      { kind: "reference", numId: 4, ilvl: 2 },
    ],
    ["a reference without one", { kind: "reference", numId: 4 }, { kind: "reference", numId: 4 }],
    ["a level alone", { kind: "levelOnly", ilvl: 3 }, { kind: "levelOnly", ilvl: 3 }],
  ])("accepts %s", (_name, stored, expected) => {
    expect(paragraphNumberingFromAttrValue(stored)).toEqual(expected);
  });

  test.each([
    ["the pre-union slot pair", { numId: 4, ilvl: 2 }],
    ["the pre-union cancellation", { numId: 0 }],
    ["an unknown kind", { kind: "sentinel", numId: 4 }],
    ["the reserved id inside a reference", { kind: "reference", numId: 0 }],
    ["a negative id", { kind: "reference", numId: -1 }],
    ["a fractional level", { kind: "reference", numId: 1, ilvl: 0.5 }],
    ["a level alone with no level", { kind: "levelOnly" }],
    ["an array", [{ kind: "none" }]],
    ["a bare number", 4],
  ])("refuses %s", (_name, stored) => {
    expect(paragraphNumberingFromAttrValue(stored)).toBeNull();
  });

  test("a reference states no level unless the attr did", () => {
    expect(paragraphNumberingFromAttrValue({ kind: "reference", numId: 4 })).not.toHaveProperty(
      "ilvl",
    );
  });
});

describe("paragraphNumberingAttr", () => {
  test("mints a value the shape check accepts, for every arm", () => {
    for (const numbering of [
      { kind: "none" },
      { kind: "reference", numId: 2 },
      { kind: "reference", numId: 2, ilvl: 5 },
      { kind: "levelOnly", ilvl: 5 },
    ] as const) {
      const minted = paragraphNumberingAttr(numbering);
      expect(minted).toEqual(numbering);
      expect(paragraphNumberingFromAttrValue(minted)).toEqual(numbering);
    }
  });
});

describe("readParagraphNumberingAttr", () => {
  test("reads the attr's absent state as absent", () => {
    expect(readParagraphNumberingAttr(null)).toBeNull();
    expect(readParagraphNumberingAttr(undefined)).toBeNull();
  });

  test("panics on a shape this build does not store, instead of ignoring it", () => {
    expect(() => readParagraphNumberingAttr({ numId: 4, ilvl: 2 })).toThrow(Panic);
  });
});

describe("a rejected w:pPrChange rebuilds the serializer's pPr source", () => {
  test("carries the stored numbering across as the union", () => {
    expect(
      paragraphRejectOriginalFormatting(
        { numPr: paragraphNumberingAttr({ kind: "reference", numId: 1, ilvl: 0 }) },
        null,
      ),
    ).toEqual({ numPr: { kind: "reference", numId: 1, ilvl: 0 } });
  });

  test("drops the numbering a record states nothing about", () => {
    expect(paragraphRejectOriginalFormatting({ styleId: "Normal" }, null)).toEqual({
      styleId: "Normal",
    });
  });

  /**
   * The leak, as a refusal. A record carrying the pre-union slot pair used to
   * be copied into `_originalFormatting` unexamined, and the serializer then
   * compared it against a union and found it different from itself.
   */
  test("refuses a record carrying the pre-union shape", () => {
    // @ts-expect-error — the attr type no longer admits the slot pair; the
    // test states what a stored record from an older build would hold.
    expect(() => paragraphRejectOriginalFormatting({ numPr: { numId: 1, ilvl: 0 } }, null)).toThrow(
      Panic,
    );
  });
});
