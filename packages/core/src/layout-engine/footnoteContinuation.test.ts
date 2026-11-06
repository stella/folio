import { describe, expect, test } from "bun:test";

import {
  applyFootnoteContinuation,
  footnoteBreakOffsets,
  footnoteContinuationSettled,
  footnoteReservationHeights,
  planFootnoteContinuation,
} from "./footnoteContinuation";
import { layoutDocument } from "./index";
import type { FlowBlock, Fragment, Measure, Page } from "./types";

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
    const pages = [
      pageOf(1, 200, [1]),
      pageOf(2, 100),
      pageOf(3, 40),
      pageOf(4, 40),
      pageOf(5, 40),
    ];
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

  test("a note longer than the document's pages asks for pages for its rest", () => {
    const contentById = new Map([[1, noteOf(150)]]);
    const pages = [pageOf(1, 600, [1])];
    const plan = planFootnoteContinuation({ pages, contentById });
    expect(plan.extraPages).toBeGreaterThan(0);
    expect(plan.trailingPages).toBe(plan.extraPages);

    const laidOut = [
      ...pages,
      ...Array.from({ length: plan.extraPages }, (_, index) =>
        pageOf(index + 2, 0),
      ),
    ];
    const next = planFootnoteContinuation({
      pages: laidOut,
      contentById,
      trailingPages: plan.trailingPages,
    });
    expect(next.extraPages).toBe(0);
    expect(next.trailingPages).toBe(plan.trailingPages);
    const applied = applyFootnoteContinuation(laidOut, next);
    expect(applied.length).toBe(laidOut.length);
    expect(applied.at(-1)?.fragments).toEqual([]);
    expect(applied.at(-1)?.footnoteWindows?.at(-1)?.to).toBe(150 * LINE);
  });

  test("a trailing page no note needs any more goes away", () => {
    const plan = planFootnoteContinuation({
      pages: [pageOf(1, 100, [1]), pageOf(2, 0), pageOf(3, 0)],
      contentById: new Map([[1, noteOf(3)]]),
      trailingPages: 2,
    });
    expect(plan.trailingPages).toBe(0);
  });

  test("a line that lacks room here waits for the next page", () => {
    const contentById = new Map([[1, noteOf(30)]]);
    const plan = planFootnoteContinuation({
      pages: [pageOf(1, CONTENT - SEPARATOR - 5, [1]), pageOf(2, 100)],
      contentById,
    });
    expect(plan.pages.get(1)).toEqual({ ids: [], windows: [], height: 0 });
    expect(plan.pages.get(2)?.ids).toEqual([1]);
    expect(plan.pages.get(2)?.windows).toEqual([]);
  });

  test("a line taller than any page still shows", () => {
    const contentById = new Map([
      [
        1,
        {
          height: 2000,
          measures: [
            {
              kind: "paragraph",
              totalHeight: 2000,
              lines: [{ lineHeight: 2000 }],
            } as unknown as Measure,
          ],
        },
      ],
    ]);
    const plan = planFootnoteContinuation({
      pages: [pageOf(1, 100, [1])],
      contentById,
    });
    // It shows whole on its reference's page, past the room there is.
    expect(plan.pages.size).toBe(0);
    expect(plan.extraPages).toBe(0);
  });

  test("a page with no content area ends the plan", () => {
    const noRoom = {
      ...pageOf(1, 0, [1]),
      margins: { ...MARGINS, top: 500, bottom: 500 },
    };
    const plan = planFootnoteContinuation({
      pages: [noRoom],
      contentById: new Map([[1, noteOf(3)]]),
    });
    expect(plan.reservedHeights.size).toBe(0);
    expect(plan.extraPages).toBeGreaterThan(0);
  });

  test("each page reserves from its own size", () => {
    const smaller = { ...pageOf(2, 100), size: { w: 800, h: 600 } };
    const plan = planFootnoteContinuation({
      pages: [pageOf(1, 200, [1]), smaller, pageOf(3, 100)],
      contentById: new Map([[1, noteOf(150)]]),
    });
    expect(plan.reservedHeights.get(2)).toBe(400 * 0.75);
    expect(plan.reservedHeights.get(3)).toBe(CONTENT * 0.75);
  });

  test("a later note waits behind the one still running", () => {
    const contentById = new Map([
      [1, noteOf(150)],
      [2, noteOf(2)],
    ]);
    const plan = planFootnoteContinuation({
      pages: [
        pageOf(1, 200, [1, 2]),
        pageOf(2, 100),
        pageOf(3, 100),
        pageOf(4, 100),
        pageOf(5, 100),
        pageOf(6, 100),
      ],
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
    expect(
      footnoteContinuationSettled(
        { reservedHeights: undefined, trailingPages: 0 },
        plan,
      ),
    ).toBe(true);
    expect(applyFootnoteContinuation(pages, plan)).toEqual(pages);
  });

  test("reserves its own height", () => {
    expect(footnoteReservationHeights(new Map([[1, 60]]), CONTENT)).toEqual(
      new Map([[1, 60]]),
    );
    expect(footnoteReservationHeights(new Map([[1, 900]]), CONTENT)).toEqual(
      new Map([[1, CONTENT * 0.75]]),
    );
  });
});

describe("where a note can be cut", () => {
  test("below a line or at the note's end", () => {
    expect(footnoteBreakOffsets(noteOf(3))).toEqual([20, 40, 60]);
  });

  test("below the lines the painters draw, with paragraph spacing after them", () => {
    const measures = [
      {
        kind: "paragraph",
        totalHeight: 50,
        lines: [{ lineHeight: 20 }, { lineHeight: 20 }],
      } as unknown as Measure,
    ];
    expect(footnoteBreakOffsets({ height: 50, measures })).toEqual([20, 50]);
  });

  test("below a table's rows", () => {
    const measures = [
      {
        kind: "table",
        totalHeight: 60,
        rows: [{ height: 30 }, { height: 10 }, { height: 20 }],
      } as unknown as Measure,
    ];
    expect(footnoteBreakOffsets({ height: 60, measures })).toEqual([
      30, 40, 60,
    ]);
  });

  test("past a picture or a text box, counting their heights", () => {
    const measures = [
      { kind: "image", height: 40 } as unknown as Measure,
      { kind: "textBox", height: 30 } as unknown as Measure,
      ...noteOf(2).measures,
    ];
    expect(footnoteBreakOffsets({ height: 110, measures })).toEqual([
      40, 70, 90, 110,
    ]);
  });
});

describe("the layout's own pages for a note's rest", () => {
  const blocks: FlowBlock[] = [
    {
      kind: "paragraph",
      id: "body",
      runs: [{ kind: "text", text: "body", footnoteRefId: 1 }],
      attrs: {},
    },
  ];
  const measures: Measure[] = [
    {
      kind: "paragraph",
      totalHeight: LINE,
      lines: [
        {
          fromRun: 0,
          fromChar: 0,
          toRun: 0,
          toChar: 4,
          width: 100,
          ascent: 16,
          descent: 4,
          lineHeight: LINE,
        },
      ],
    },
  ];
  const options = { pageSize: SIZE, margins: MARGINS };

  test("number on from the last body page", () => {
    const { pages } = layoutDocument(blocks, measures, {
      ...options,
      footnoteContinuation: { trailingPages: 2 },
    });
    expect(pages.map((page) => page.number)).toEqual([1, 2, 3]);
    expect(pages.map((page) => page.logicalNumber)).toEqual([1, 2, 3]);
    expect(pages.map((page) => page.sectionPageNumber)).toEqual([1, 2, 3]);
  });

  test("a reference reserves at most a share of its page", () => {
    const { pages } = layoutDocument(blocks, measures, {
      ...options,
      footnoteHeightById: new Map([[1, 5000]]),
      footnoteContinuation: { trailingPages: 0 },
    });
    expect(pages.length).toBe(1);
    expect(pages[0]?.footnoteReservedHeight).toBe(CONTENT * 0.75 + SEPARATOR);
  });
});
