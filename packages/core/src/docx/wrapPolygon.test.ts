import { describe, expect, test } from "bun:test";

import { requiredWrapPolygon, wrapPolygonFor } from "./wrapPolygon";

const AUTHORED = {
  edited: true,
  start: { x: 1_200, y: 0 },
  lineTo: [
    { x: 0, y: 8_400 },
    { x: 21_600, y: 21_600 },
  ],
};

describe("the wrap polygon a wrap of each kind states", () => {
  test("only a tight or through wrap has one", () => {
    for (const type of ["inline", "square", "topAndBottom", "behind", "inFront"] as const) {
      expect(wrapPolygonFor(type, undefined)).toBeUndefined();
      // An outline the drawing already carries is not written on a wrap whose
      // type declares none, and not thrown away either: the caller keeps it.
      expect(wrapPolygonFor(type, AUTHORED)).toBeUndefined();
    }
  });

  test("becoming tight or through mints the rectangle", () => {
    for (const type of ["tight", "through"] as const) {
      expect(wrapPolygonFor(type, undefined)).toEqual({
        edited: false,
        start: { x: 0, y: 0 },
        lineTo: [
          { x: 0, y: 21_600 },
          { x: 21_600, y: 21_600 },
          { x: 21_600, y: 0 },
          { x: 0, y: 0 },
        ],
      });
    }
  });

  test("the rectangle is minted once: a second pass returns what the first left", () => {
    const minted = wrapPolygonFor("tight", undefined);
    expect(wrapPolygonFor("tight", minted)).toBe(minted);
    expect(wrapPolygonFor("through", minted)).toBe(minted);
  });

  test("an authored outline is never replaced", () => {
    expect(wrapPolygonFor("tight", AUTHORED)).toBe(AUTHORED);
    expect(requiredWrapPolygon(AUTHORED)).toBe(AUTHORED);
  });

  test("a path shorter than the schema admits is written back as it stands", () => {
    // folio is not the validator of its input. Substituting the full extent of
    // the drawing would move text the source flows through the object.
    const short = { start: { x: 1, y: 1 }, lineTo: [{ x: 2, y: 2 }] };
    expect(requiredWrapPolygon(short)).toBe(short);
  });
});
