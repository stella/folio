import { describe, expect, test } from "bun:test";

import type { SelectionFormatting } from "./toolbarPrimitives";
import { areSelectionFormattingEqual } from "./toolbarUtils";

describe("selection formatting equality", () => {
  test("treats identical nullable values as equal", () => {
    expect(areSelectionFormattingEqual(undefined, undefined)).toBe(true);
    expect(areSelectionFormattingEqual(null, null)).toBe(true);
  });

  test("treats one missing formatting value as different", () => {
    expect(areSelectionFormattingEqual(undefined, {})).toBe(false);
    expect(areSelectionFormattingEqual({}, null)).toBe(false);
  });

  test("compares formatting fields structurally", () => {
    const formatting = {
      bold: true,
      fontSize: 22,
      listState: { type: "bullet", level: 1 },
    } satisfies SelectionFormatting;
    const equivalent = {
      bold: true,
      fontSize: 22,
      listState: { type: "bullet", level: 1 },
    } satisfies SelectionFormatting;
    const different = {
      bold: true,
      fontSize: 22,
      listState: { type: "numbered", level: 1 },
    } satisfies SelectionFormatting;

    expect(areSelectionFormattingEqual(formatting, equivalent)).toBe(true);
    expect(areSelectionFormattingEqual(formatting, different)).toBe(false);
  });
});
