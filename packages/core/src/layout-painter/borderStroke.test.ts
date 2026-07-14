import { describe, expect, test } from "bun:test";

import {
  borderStrokeToCss,
  resolveCssBorderStroke,
  resolveParagraphBorderHorizontalOutsets,
} from "./borderStroke";

describe("border stroke painting", () => {
  test("projects authored subpixel width into optical coverage", () => {
    expect(borderStrokeToCss({ width: 1 / 3, style: "solid", color: "#123456" })).toBe(
      "1px solid color-mix(in srgb, #123456 33.3333%, transparent)",
    );
  });

  test("keeps full-pixel and explicit hairline strokes opaque", () => {
    expect(resolveCssBorderStroke({ width: 2, style: "dashed", color: "#123456" })).toEqual({
      width: 2,
      style: "dashed",
      color: "#123456",
    });
    expect(resolveCssBorderStroke({ width: 0, style: "solid" })).toEqual({
      width: 1,
      style: "solid",
      color: "#000000",
    });
  });

  test("extends horizontal paragraph rules beyond text when side borders are absent", () => {
    expect(resolveParagraphBorderHorizontalOutsets({ bottom: { width: 1 } }, true)).toEqual({
      left: 2,
      right: 2,
    });
    expect(
      resolveParagraphBorderHorizontalOutsets(
        { left: { width: 2, space: 3 }, bottom: { width: 1 } },
        true,
      ),
    ).toEqual({ left: 5, right: 2 });
  });
});
