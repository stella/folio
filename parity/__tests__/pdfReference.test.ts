import { describe, expect, test } from "bun:test";

import { canReusePagePngCache, refreshCachedReferenceGeom } from "../pdfReference";
import type { DocGeom } from "../types";

describe("PDF reference page cache", () => {
  test("recomputes stale text-derived facts", () => {
    const cached = {
      source: "word",
      file: "/old.docx",
      meta: {},
      pages: [
        {
          number: 1,
          widthPt: 100,
          heightPt: 100,
          lines: [
            {
              text: "عنوان عربي",
              normText: "stale",
              xPt: 0,
              yPt: 0,
              widthPt: 10,
              heightPt: 10,
              region: "body",
              direction: "ltr",
            },
            {
              text: "English",
              normText: "stale",
              xPt: 0,
              yPt: 10,
              widthPt: 10,
              heightPt: 10,
              region: "body",
              direction: "rtl",
            },
            {
              text: "(...) 123",
              normText: "stale",
              xPt: 0,
              yPt: 20,
              widthPt: 10,
              heightPt: 10,
              region: "body",
              direction: "rtl",
            },
          ],
        },
      ],
    } as const satisfies DocGeom;

    const refreshed = refreshCachedReferenceGeom(cached, "/current.docx");

    expect(refreshed.file).toBe("/current.docx");
    expect(refreshed.pages[0]?.lines.map(({ direction }) => direction)).toEqual([
      "rtl",
      "ltr",
      "unknown",
    ]);
    expect(refreshed.pages[0]?.lines[0]?.normText).toBe("عنوان عربي");
  });

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
