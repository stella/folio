import { describe, expect, test } from "bun:test";

import {
  calculateResizedImageDimensions,
  IMAGE_RESIZE_HANDLES,
  type ImageResizeHandle,
} from "./imageResize";

const resizeHandles = IMAGE_RESIZE_HANDLES;
const cornerHandles = IMAGE_RESIZE_HANDLES.filter(
  (handle) => handle.length === 2,
);
const horizontalReflections = {
  nw: "ne",
  ne: "nw",
  se: "sw",
  sw: "se",
  n: "n",
  s: "s",
  e: "w",
  w: "e",
} as const satisfies Record<ImageResizeHandle, ImageResizeHandle>;
const verticalReflections = {
  nw: "sw",
  ne: "se",
  se: "ne",
  sw: "nw",
  n: "s",
  s: "n",
  e: "e",
  w: "w",
} as const satisfies Record<ImageResizeHandle, ImageResizeHandle>;

describe("calculateResizedImageDimensions", () => {
  test.each(cornerHandles)(
    "locked %s corner preserves aspect ratio",
    (handle) => {
      const result = calculateResizedImageDimensions({
        handle,
        deltaX: handle.includes("w") ? -100 : 100,
        deltaY: handle.includes("n") ? -10 : 10,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      });

      expect(result).toEqual({ width: 300, height: 150 });
      expect(result.width / result.height).toBeCloseTo(2, 5);
      expect(result.width).toBeGreaterThan(200);
      expect(result.height).toBeGreaterThan(100);
    },
  );

  test("free corner resizes only the dimensions selected by the drag", () => {
    const result = calculateResizedImageDimensions({
      handle: "se",
      deltaX: 100,
      deltaY: 0,
      startWidth: 200,
      startHeight: 100,
      aspect: "free",
    });

    expect(result).toEqual({ width: 300, height: 100 });
  });

  test("edge handles resize only their driven axis", () => {
    expect(
      calculateResizedImageDimensions({
        handle: "e",
        deltaX: 80,
        deltaY: 999,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      }),
    ).toEqual({ width: 280, height: 100 });
    expect(
      calculateResizedImageDimensions({
        handle: "w",
        deltaX: 50,
        deltaY: 0,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      }),
    ).toEqual({ width: 150, height: 100 });
    expect(
      calculateResizedImageDimensions({
        handle: "s",
        deltaX: 999,
        deltaY: 60,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      }),
    ).toEqual({ width: 200, height: 160 });
    expect(
      calculateResizedImageDimensions({
        handle: "n",
        deltaX: 0,
        deltaY: 40,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      }),
    ).toEqual({ width: 200, height: 60 });
  });

  test.each(cornerHandles)(
    "locked %s corner shrinks along the dominant axis",
    (handle) => {
      const result = calculateResizedImageDimensions({
        handle,
        deltaX: handle.includes("w") ? 100 : -100,
        deltaY: handle.includes("n") ? 10 : -10,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      });

      expect(result).toEqual({ width: 100, height: 50 });
    },
  );

  test.each(cornerHandles)(
    "locked %s corner clamps the common scale",
    (handle) => {
      const result = calculateResizedImageDimensions({
        handle,
        deltaX: handle.includes("w") ? -10_000 : 10_000,
        deltaY: handle.includes("n") ? -10_000 : 10_000,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      });

      expect(result).toEqual({ width: 2000, height: 1000 });
      expect(result.width / result.height).toBeCloseTo(2, 5);
    },
  );

  test.each(cornerHandles)(
    "locked %s corner caps impossible minimum constraints",
    (handle) => {
      const result = calculateResizedImageDimensions({
        handle,
        deltaX: handle.includes("w") ? -100 : 100,
        deltaY: handle.includes("n") ? -100 : 100,
        startWidth: 700,
        startHeight: 1,
        aspect: "locked",
      });

      expect(result.width).toBeLessThanOrEqual(2000);
      expect(result.height).toBeLessThanOrEqual(2000);
      expect(result.height).toBeLessThan(20);
      expect(result.width / result.height).toBeCloseTo(700, 5);
    },
  );

  test.each(resizeHandles)(
    "%s keeps its non-driven axis unchanged and bounds driven axes",
    (handle) => {
      const result = calculateResizedImageDimensions({
        handle,
        deltaX: 10_000,
        deltaY: -10_000,
        startWidth: 200,
        startHeight: 5000,
        aspect: "free",
      });

      if (handle.includes("w") || handle.includes("e")) {
        expect(result.width).toBe(handle.includes("w") ? 20 : 2000);
      } else {
        expect(result.width).toBe(200);
      }
      if (handle.includes("n") || handle.includes("s")) {
        expect(result.height).toBe(handle.includes("n") ? 2000 : 20);
      } else {
        expect(result.height).toBe(5000);
      }
    },
  );

  test("reflected handles produce matching dimensions for reflected deltas", () => {
    for (const handle of resizeHandles) {
      const horizontalResult = calculateResizedImageDimensions({
        handle,
        deltaX: 37,
        deltaY: -19,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      });
      const horizontalReflection = calculateResizedImageDimensions({
        handle: horizontalReflections[handle],
        deltaX: -37,
        deltaY: -19,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      });
      const verticalResult = calculateResizedImageDimensions({
        handle: verticalReflections[handle],
        deltaX: 37,
        deltaY: 19,
        startWidth: 200,
        startHeight: 100,
        aspect: "locked",
      });

      expect(horizontalReflection).toEqual(horizontalResult);
      expect(verticalResult).toEqual(horizontalResult);
    }
  });

  test("zero start dimensions remain finite", () => {
    const result = calculateResizedImageDimensions({
      handle: "se",
      deltaX: 100,
      deltaY: 50,
      startWidth: 0,
      startHeight: 0,
      aspect: "locked",
    });

    expect(Number.isFinite(result.width)).toBe(true);
    expect(Number.isFinite(result.height)).toBe(true);
  });
});
