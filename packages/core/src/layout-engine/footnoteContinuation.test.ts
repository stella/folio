import { describe, expect, test } from "bun:test";

import {
  applyFootnoteContinuation,
  footnoteBreakOffsets,
  footnoteReservationHeight,
  footnoteReservationsSettled,
  planFootnoteContinuation,
} from "./footnoteContinuation";
import type { Fragment, Measure, Page } from "./types";

const LINE = 20;
const SIZE = { w: 800, h: 1000 };
const MARGINS = { top: 100, bottom: 100, left: 100, right: 100 };
const CONTENT = 800;
const SEPARATOR = 12;

const noteOf = (lines: number) => ({
  height: lines * LINE,
  measures: [
    {
      kind: "paragraph",
      totalHeight: lines * LINE,
      lines: Array.from({ length: lines }, () => ({ lineHeight: LINE })),
    } as unknown as Measure,
  ],
});

const pageOf = (number: number, bodyHeight: number, footnoteIds?: number[]) =>
  ({
    number,
    size: SIZE,
    margins: MARGINS,
    fragments:
      bodyHeight > 0
        ? [
            {
              kind: "paragraph",
              y: MARGINS.top,
              height: bodyHeight,
            } as Fragment,
          ]
        : [],
    ...(footnoteIds ? { footnoteIds } : {}),
  }) as Page;

describe("a footnote too tall for the room on its page", () => {
  test("every line shows once, in order, across the pages", () => {
    const contentById = new Map([[1, noteOf(150)]]);
    const pages = [pageOf(1, 200, [1]), pageOf(2, 100), pageOf(3, 40)];
    const plan = planFootnoteContinuation({ pages, contentById });
    const applied = applyFootnoteContinuation(pages, plan);
    const windows = applied.flatMap((page) => page.footnoteWindows ?? []);
    expect(windows.at(0)?.from).toBe(0);
    expect(windows.at(-1)?.to).toBe(150 * LINE);
    for (const [index, shown] of windows.entries()) {
      if (index > 0) expect(shown.from).toBe(windows[index - 1]?.to ?? -1);
      expect(shown.from % LINE).toBe(0);
      expect(shown.to % LINE).toBe(0);
    }
  });

  test("the first page's window fills what the body leaves", () => {
    const contentById = new Map([[1, noteOf(150)]]);
    const plan = planFootnoteContinuation({
      pages: [pageOf(1, 200, [1]), pageOf(2, 100)],
      contentById,
    });
    const first = plan.pages.get(1)?.windows[0];
    const room = CONTENT - 200 - SEPARATOR;
    expect(first?.to).toBe(Math.floor(room / LINE) * LINE);
  });

  test("the next page reserves room for the rest, up to its share", () => {
    const contentById = new Map([[1, noteOf(150)]]);
    const plan = planFootnoteContinuation({
      pages: [pageOf(1, 200, [1]), pageOf(2, 100)],
      contentById,
    });
    expect(plan.reservedHeights.get(2)).toBe(CONTENT * 0.75);
    expect(plan.reservedHeights.get(3)).toBe(CONTENT * 0.75);
    expect(plan.pages.get(1)?.height).toBeLessThanOrEqual(CONTENT - 200);
  });

  test("a note longer than the document's pages adds pages for its rest", () => {
    const contentById = new Map([[1, noteOf(150)]]);
    const pages = [pageOf(1, 600, [1])];
    const plan = planFootnoteContinuation({ pages, contentById });
    expect(plan.extraPages).toBeGreaterThan(0);
    const applied = applyFootnoteContinuation(pages, plan);
    expect(applied.length).toBe(1 + plan.extraPages);
    expect(applied.at(-1)?.fragments).toEqual([]);
    expect(applied.at(-1)?.footnoteWindows?.at(-1)?.to).toBe(150 * LINE);
  });

  test("a later note waits behind the one still running", () => {
    const contentById = new Map([
      [1, noteOf(150)],
      [2, noteOf(2)],
    ]);
    const plan = planFootnoteContinuation({
      pages: [pageOf(1, 200, [1, 2]), pageOf(2, 100), pageOf(3, 100)],
      contentById,
    });
    expect(plan.pages.get(1)?.ids).toEqual([1]);
    const placed = [...plan.pages.values()];
    expect(placed.at(-1)?.ids).toContain(2);
    expect(placed.flatMap(({ ids }) => ids).at(-1)).toBe(2);
  });
});

describe("a footnote that fits", () => {
  test("changes nothing", () => {
    const contentById = new Map([[1, noteOf(3)]]);
    const pages = [pageOf(1, 200, [1]), pageOf(2, 100)];
    const plan = planFootnoteContinuation({ pages, contentById });
    expect(plan.pages.size).toBe(0);
    expect(plan.reservedHeights.size).toBe(0);
    expect(footnoteReservationsSettled(undefined, plan)).toBe(true);
    expect(applyFootnoteContinuation(pages, plan)).toEqual(pages);
  });

  test("reserves its own height", () => {
    expect(
      footnoteReservationHeight(60, { size: SIZE, margins: MARGINS }),
    ).toBe(60);
  });
});

test("a window ends at a line bottom or the note's end", () => {
  expect(footnoteBreakOffsets(noteOf(3))).toEqual([20, 40, 60]);
});
