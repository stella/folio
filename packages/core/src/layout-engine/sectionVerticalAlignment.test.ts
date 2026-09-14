import { describe, expect, test } from "bun:test";

import { applySectionVerticalAlignment } from "./sectionVerticalAlignment";
import type { Layout, Page } from "./types";

const page = (): Page => ({
  number: 1,
  logicalNumber: 1,
  sectionIndex: 0,
  margins: { top: 100, right: 100, bottom: 100, left: 100 },
  size: { w: 800, h: 1_000 },
  fragments: [
    {
      kind: "paragraph",
      blockId: "body",
      x: 100,
      y: 100,
      width: 600,
      height: 100,
      fromLine: 0,
      toLine: 1,
    },
    {
      kind: "image",
      blockId: "page-anchor",
      x: 0,
      y: 20,
      width: 40,
      height: 40,
      isAnchored: true,
    },
  ],
});

const layout = (): Layout => ({ pageSize: { w: 800, h: 1_000 }, pages: [page()] });

describe("section vertical alignment", () => {
  test("centers flowing body content without moving page-positioned objects", () => {
    const aligned = applySectionVerticalAlignment(layout(), ["center"]);

    expect(aligned.pages[0]?.fragments.map(({ y }) => y)).toEqual([450, 20]);
  });

  test("bottom-aligns flowing body content", () => {
    const aligned = applySectionVerticalAlignment(layout(), ["bottom"]);

    expect(aligned.pages[0]?.fragments.map(({ y }) => y)).toEqual([800, 20]);
  });

  test("does not move top-aligned or vertically justified content", () => {
    const original = layout();

    expect(applySectionVerticalAlignment(original, ["top"])).toBe(original);
    expect(applySectionVerticalAlignment(original, ["both"])).toBe(original);
  });
});
