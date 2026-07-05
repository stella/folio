import { describe, expect, test } from "bun:test";

import { measureDirectiveGutter } from "./rangeProjection";

describe("measureDirectiveGutter", () => {
  test("returns null for a non-positive zoom instead of dividing by it", () => {
    // Every measurement in the function divides by `zoom`; a zero/negative factor
    // would produce Infinity/NaN rail geometry, so the guard must short-circuit
    // before touching the element. A bare stub therefore stands in for it.
    // SAFETY: the guard returns before the container is dereferenced, so its
    // concrete shape is never read on this path.
    const stub = {} as unknown as HTMLElement;

    expect(measureDirectiveGutter(stub, 0)).toBeNull();
    expect(measureDirectiveGutter(stub, -0.5)).toBeNull();
  });
});
