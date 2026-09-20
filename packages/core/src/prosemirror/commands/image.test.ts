import { describe, expect, test } from "bun:test";

import { computeImageTransform, constrainImageSize, resolveImageWrap } from "./image";

describe("resolveImageWrap", () => {
  test("inline keeps the image in the text flow", () => {
    expect(resolveImageWrap("inline")).toEqual({
      wrapType: "inline",
      displayMode: "inline",
      cssFloat: undefined,
    });
  });

  test("square/tight/through float left", () => {
    for (const wrapType of ["square", "tight", "through"] as const) {
      expect(resolveImageWrap(wrapType)).toEqual({
        wrapType,
        displayMode: "float",
        cssFloat: "left",
      });
    }
  });

  test("wrapLeft and wrapRight map to a square float on the opposite side", () => {
    expect(resolveImageWrap("wrapLeft")).toEqual({
      wrapType: "square",
      displayMode: "float",
      cssFloat: "right",
    });
    expect(resolveImageWrap("wrapRight")).toEqual({
      wrapType: "square",
      displayMode: "float",
      cssFloat: "left",
    });
  });

  test("behind/inFront float with no css float", () => {
    expect(resolveImageWrap("behind")).toEqual({
      wrapType: "behind",
      displayMode: "float",
      cssFloat: "none",
    });
  });

  test("unknown wrap modes are rejected", () => {
    expect(resolveImageWrap("nope")).toBeNull();
  });
});

describe("computeImageTransform", () => {
  test("rotate wraps within 0..359 degrees", () => {
    expect(computeImageTransform({ rotation: 180 }, "rotateCW")).toEqual({ rotation: 270 });
    expect(computeImageTransform(undefined, "rotateCCW")).toEqual({ rotation: 270 });
  });

  test("rotating back to 0 degrees states the zero", () => {
    expect(computeImageTransform({ rotation: 270 }, "rotateCW")).toEqual({ rotation: 0 });
  });

  test("flips toggle and combine with rotation", () => {
    expect(computeImageTransform(undefined, "flipH")).toEqual({ flipH: true });
    expect(computeImageTransform({ flipH: true }, "flipH")).toEqual({ flipH: false });
    expect(computeImageTransform({ rotation: 90, flipH: true }, "flipV")).toEqual({
      rotation: 90,
      flipH: true,
      flipV: true,
    });
  });
});

describe("constrainImageSize", () => {
  test("leaves images within the max width untouched", () => {
    expect(constrainImageSize(300, 200)).toEqual({ width: 300, height: 200 });
  });

  test("scales down oversized images preserving aspect ratio", () => {
    expect(constrainImageSize(1224, 612)).toEqual({ width: 612, height: 306 });
  });
});
