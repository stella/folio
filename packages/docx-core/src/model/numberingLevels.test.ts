import { describe, expect, test } from "bun:test";
import type { ListLevel, NumberingDefinitions } from "./lists";
import { levelShowsMarker, numberingLevelEntries, resolveNumberingLevel } from "./numberingLevels";

const level = (ilvl: number, overrides: Partial<ListLevel> = {}): ListLevel => ({
  ilvl,
  numFmt: "decimal",
  lvlText: `%${ilvl + 1}.`,
  ...overrides,
});

const definitions = (overrides: Partial<NumberingDefinitions> = {}): NumberingDefinitions => ({
  abstractNums: [{ abstractNumId: 1, levels: [level(0)] }],
  nums: [{ numId: 1, abstractNumId: 1 }],
  ...overrides,
});

describe("resolveNumberingLevel", () => {
  test("rejects levels outside 0..8 without materialising them", () => {
    expect(resolveNumberingLevel(definitions(), 1, 9)).toEqual({
      kind: "undefined",
      reason: "levelOutOfRange",
    });
  });

  test("distinguishes missing instances and abstracts", () => {
    expect(resolveNumberingLevel(definitions(), 2, 0)).toEqual({
      kind: "undefined",
      reason: "noInstance",
    });
    expect(
      resolveNumberingLevel(definitions({ nums: [{ numId: 1, abstractNumId: 9 }] }), 1, 0),
    ).toEqual({ kind: "undefined", reason: "noAbstract" });
  });

  test("full overrides define an otherwise absent level", () => {
    const override = level(2, { numFmt: "lowerLetter" });
    expect(
      resolveNumberingLevel(
        definitions({
          nums: [{ numId: 1, abstractNumId: 1, levelOverrides: [{ ilvl: 2, lvl: override }] }],
        }),
        1,
        2,
      ),
    ).toEqual({ kind: "level", level: override, source: "override" });
  });

  test("start-only overrides patch an existing level", () => {
    const resolution = resolveNumberingLevel(
      definitions({
        nums: [{ numId: 1, abstractNumId: 1, levelOverrides: [{ ilvl: 0, startOverride: 7 }] }],
      }),
      1,
      0,
    );
    expect(resolution.kind === "level" ? resolution.level.start : undefined).toBe(7);
  });

  test("follows style links and reports an unresolved link", () => {
    const linked = definitions({
      abstractNums: [
        { abstractNumId: 1, numStyleLink: "Numbered", levels: [] },
        { abstractNumId: 2, styleLink: "Numbered", levels: [level(0)] },
      ],
    });
    expect(resolveNumberingLevel(linked, 1, 0)).toMatchObject({
      kind: "level",
      source: "styleLink",
    });
    linked.abstractNums.pop();
    expect(resolveNumberingLevel(linked, 1, 0)).toEqual({
      kind: "undefined",
      reason: "unresolvedStyleLink",
    });
  });

  test("first duplicate ids win", () => {
    const first = level(0, { numFmt: "upperRoman" });
    const duplicate = definitions({
      abstractNums: [
        { abstractNumId: 1, levels: [first] },
        { abstractNumId: 1, levels: [level(0, { numFmt: "bullet" })] },
      ],
      nums: [
        { numId: 1, abstractNumId: 1 },
        { numId: 1, abstractNumId: 99 },
      ],
    });
    expect(resolveNumberingLevel(duplicate, 1, 0)).toEqual({
      kind: "level",
      level: first,
      source: "abstract",
    });
  });

  test("entries include override-only levels", () => {
    const override = level(3);
    expect(
      numberingLevelEntries(
        definitions({
          nums: [{ numId: 1, abstractNumId: 1, levelOverrides: [{ ilvl: 3, lvl: override }] }],
        }),
      ),
    ).toContainEqual({ numId: 1, ilvl: 3, level: override });
  });
});

describe("levelShowsMarker", () => {
  test.each([
    ["visible", level(0), true],
    ["hidden", level(0, { rPr: { hidden: true } }), false],
    ["none", level(0, { numFmt: "none" }), true],
    ["none placeholder", level(0, { numFmt: "none", lvlText: "%1" }), false],
    ["none legal", level(0, { numFmt: "none", isLgl: true, lvlText: "%1" }), true],
    ["empty decimal", level(0, { lvlText: "" }), true],
    ["empty none", level(0, { numFmt: "none", lvlText: "" }), false],
  ] as const)("handles %s", (_name, candidate, expected) => {
    expect(
      levelShowsMarker(
        definitions({ abstractNums: [{ abstractNumId: 1, levels: [candidate] }] }),
        1,
        0,
      ),
    ).toBe(expected);
  });

  test("the current level's legal mode renders cross-level placeholders as decimal", () => {
    expect(
      levelShowsMarker(
        definitions({
          abstractNums: [
            {
              abstractNumId: 1,
              levels: [
                level(0, { numFmt: "none", lvlText: "%1" }),
                level(1, { isLgl: true, lvlText: "%1" }),
              ],
            },
          ],
        }),
        1,
        1,
      ),
    ).toBe(true);
  });
});
