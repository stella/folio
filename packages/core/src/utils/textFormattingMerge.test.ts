import { describe, expect, test } from "bun:test";

import { mergeStyleTextFormatting, mergeTextFormatting } from "./textFormattingMerge";

describe("mergeTextFormatting", () => {
  test("per-slot merge of fontFamily preserves inherited ascii", () => {
    const result = mergeTextFormatting(
      { fontFamily: { ascii: "Arial Narrow" } },
      { fontFamily: { eastAsia: "Calibri" } },
    );

    expect(result?.fontFamily).toEqual({
      ascii: "Arial Narrow",
      eastAsia: "Calibri",
    });
  });

  test("shallow-merges object-shaped fields", () => {
    const result = mergeTextFormatting(
      { underline: { style: "single", color: { rgb: "FF0000" } } },
      { underline: { style: "double" } },
    );

    expect(result?.underline).toEqual({
      style: "double",
      color: { rgb: "FF0000" },
    });
  });

  test('color w:val="auto" clears an inherited explicit color', () => {
    const result = mergeTextFormatting({ color: { rgb: "FF0000" } }, { color: { auto: true } });

    expect(result?.color).toEqual({ auto: true });
  });
});

describe("mergeStyleTextFormatting", () => {
  test("a style's own off cancels a bold value inherited from its basedOn ancestor", () => {
    const result = mergeStyleTextFormatting({ bold: true }, { bold: false });

    expect(result?.bold).toBe(false);
  });

  test("a style's own on wins over an ancestor's off", () => {
    const result = mergeStyleTextFormatting({ italic: false }, { italic: true });

    expect(result?.italic).toBe(true);
  });

  test("a style that does not mention the toggle inherits the ancestor's state unchanged", () => {
    const onResult = mergeStyleTextFormatting({ strike: true }, { allCaps: true });
    const offResult = mergeStyleTextFormatting({ strike: false }, { allCaps: true });

    expect(onResult?.strike).toBe(true);
    expect(offResult?.strike).toBe(false);
  });

  test("no ancestor and an explicit off yields off, not an absent property", () => {
    const result = mergeStyleTextFormatting(undefined, { hidden: false });

    expect(result?.hidden).toBe(false);
  });
});
