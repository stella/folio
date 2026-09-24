import { describe, expect, test } from "bun:test";

import { resolveAnchorLayoutInCellCompatibility } from "./anchorLayoutInCellCompatibility";

describe("resolveAnchorLayoutInCellCompatibility", () => {
  test.each([undefined, 11, 12, 14])(
    "honours the authored attribute at compatibilityMode %p",
    (mode) => {
      expect(resolveAnchorLayoutInCellCompatibility(mode)).toBe(false);
    },
  );

  test.each([15, 16])("forces in-cell layout at compatibilityMode %p", (mode) => {
    expect(resolveAnchorLayoutInCellCompatibility(mode)).toBe(true);
  });
});
