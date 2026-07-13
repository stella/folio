import { describe, expect, test } from "bun:test";

import { canReusePagePngCache } from "../pdfReference";

describe("PDF reference page cache", () => {
  test("does not treat a bounded render as a complete unbounded cache", () => {
    expect(
      canReusePagePngCache({
        existing: ["/cache/p1.png"],
        completePageCount: null,
      }),
    ).toBe(false);
  });

  test("reuses bounded and explicitly complete sequential page caches", () => {
    const existing = ["/cache/p1.png", "/cache/p2.png"];

    expect(canReusePagePngCache({ existing, completePageCount: null, maxPages: 2 })).toBe(true);
    expect(canReusePagePngCache({ existing, completePageCount: 2 })).toBe(true);
  });

  test("rejects incomplete or non-sequential caches", () => {
    expect(
      canReusePagePngCache({
        existing: ["/cache/p1.png"],
        completePageCount: 2,
      }),
    ).toBe(false);
    expect(
      canReusePagePngCache({
        existing: ["/cache/p1.png", "/cache/p3.png"],
        completePageCount: 2,
      }),
    ).toBe(false);
  });
});
