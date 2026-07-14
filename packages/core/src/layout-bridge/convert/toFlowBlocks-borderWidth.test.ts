import { describe, expect, test } from "bun:test";

import { convertBorderSpecToLayout } from "./toFlowBlocks";

describe("border width conversion", () => {
  test("keeps eighth-point widths fractional in layout coordinates", () => {
    expect(convertBorderSpecToLayout({ style: "single", size: 2 })?.width).toBeCloseTo(1 / 3);
    expect(convertBorderSpecToLayout({ style: "single", size: 4 })?.width).toBeCloseTo(2 / 3);
  });

  test("uses a visible hairline when width is omitted", () => {
    expect(convertBorderSpecToLayout({ style: "single" })?.width).toBe(1);
  });
});
