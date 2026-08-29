import { describe, expect, test } from "bun:test";

import { reflowFootnoteColumns } from "./footnoteColumnReflow";
import type { Fragment, Layout } from "./types";

const overlappingLayout = (fragments: Fragment[]): Layout => ({
  pageSize: { w: 100, h: 100 },
  columns: { count: 2, gap: 10 },
  pages: [
    {
      number: 1,
      logicalNumber: 1,
      fragments,
      margins: { top: 0, right: 0, bottom: 0, left: 0 },
      size: { w: 100, h: 100 },
      columns: { count: 2, gap: 10 },
      footnoteReservedHeight: 20,
    },
  ],
});

const fragmentPosition = {
  blockId: "out-of-flow",
  x: 0,
  y: 75,
  width: 10,
  height: 10,
};

describe("footnote column reflow", () => {
  test("ignores fragments that do not participate in body flow", () => {
    const layout = overlappingLayout([
      {
        kind: "table",
        ...fragmentPosition,
        fromRow: 0,
        toRow: 1,
        isFloating: true,
      },
      { kind: "image", ...fragmentPosition, isAnchored: true },
      { kind: "textBox", ...fragmentPosition, isPositioned: true },
    ]);
    let passes = 0;

    const result = reflowFootnoteColumns({
      initialLayout: layout,
      runLayout: () => {
        passes += 1;
        return layout;
      },
    });

    expect(result.pages[0]!.footnoteReservedHeight).toBeUndefined();
    expect(passes).toBe(0);
  });

  test("returns a stable fallback when flow overlap persists", () => {
    const flowFragments: Fragment[] = [
      {
        kind: "paragraph",
        ...fragmentPosition,
        fromLine: 0,
        toLine: 1,
      },
      {
        kind: "table",
        ...fragmentPosition,
        fromRow: 0,
        toRow: 1,
      },
      { kind: "image", ...fragmentPosition },
      { kind: "textBox", ...fragmentPosition },
    ];

    for (const fragment of flowFragments) {
      const layout = overlappingLayout([fragment]);
      let passes = 0;

      const result = reflowFootnoteColumns({
        initialLayout: layout,
        runLayout: () => {
          passes += 1;
          return layout;
        },
      });

      expect(result.pages[0]!.footnoteReservedHeight).toBeUndefined();
      expect(passes).toBe(1);
    }
  });

  test("stops when the reserve-floor state repeats", () => {
    const first = overlappingLayout([
      {
        kind: "paragraph",
        ...fragmentPosition,
        fromLine: 0,
        toLine: 1,
      },
    ]);
    const second = overlappingLayout([
      {
        kind: "paragraph",
        ...fragmentPosition,
        y: 76,
        fromLine: 0,
        toLine: 1,
      },
    ]);
    first.pages[0]!.footnoteIds = [7];
    second.pages[0]!.footnoteIds = [7];
    let passes = 0;

    const result = reflowFootnoteColumns({
      initialLayout: first,
      runLayout: () => {
        passes += 1;
        return passes % 2 === 1 ? second : first;
      },
    });

    expect(result).toBe(second);
    expect(passes).toBe(1);
  });

  test("cleans orphan reservations on a repeated two-page floor state", () => {
    const initial = overlappingLayout([
      {
        kind: "paragraph",
        ...fragmentPosition,
        fromLine: 0,
        toLine: 1,
      },
    ]);
    initial.pages[0]!.footnoteIds = [7];
    const returned = {
      ...initial,
      pages: [
        {
          ...initial.pages[0]!,
          fragments: [{ ...initial.pages[0]!.fragments[0]!, y: 76 }],
        },
        {
          ...initial.pages[0]!,
          number: 2,
          logicalNumber: 2,
          fragments: [],
          footnoteIds: [],
        },
      ],
    };
    const returnedPage = returned.pages[1]!;

    const result = reflowFootnoteColumns({
      initialLayout: initial,
      runLayout: () => returned,
    });

    expect(result).not.toBe(returned);
    expect(result.pages[1]).not.toBe(returnedPage);
    expect(result.pages[1]!.footnoteReservedHeight).toBeUndefined();
    expect(returnedPage.footnoteReservedHeight).toBe(20);
  });

  test("keeps a caller floor when observed reservation is smaller", () => {
    const initial = overlappingLayout([
      {
        kind: "paragraph",
        ...fragmentPosition,
        fromLine: 0,
        toLine: 1,
      },
    ]);
    initial.pages[0]!.footnoteIds = [7];
    let receivedFloors: Map<number, number> | undefined;

    const result = reflowFootnoteColumns({
      initialLayout: initial,
      initialReserveFloors: new Map([[1, 80]]),
      runLayout: (floors) => {
        receivedFloors = floors;
        const page = initial.pages[0]!;
        return {
          ...initial,
          pages: [
            {
              ...page,
              fragments: [{ ...page.fragments[0]!, y: 70 }],
              footnoteReservedHeight: floors.get(1),
            },
          ],
        };
      },
    });

    expect(receivedFloors?.get(1)).toBe(80);
    expect(result.pages[0]!.footnoteReservedHeight).toBe(80);
  });

  test("cleans orphan reservations with copy-on-write", () => {
    const initial = overlappingLayout([
      {
        kind: "paragraph",
        ...fragmentPosition,
        fromLine: 0,
        toLine: 1,
      },
    ]);
    const initialPage = initial.pages[0]!;
    const cleanedInitial = reflowFootnoteColumns({
      initialLayout: {
        ...initial,
        pages: [{ ...initialPage, fragments: [{ ...initialPage.fragments[0]!, y: 70 }] }],
      },
      runLayout: () => initial,
    });

    expect(cleanedInitial).not.toBe(initial);
    expect(cleanedInitial.pages[0]).not.toBe(initialPage);
    expect(cleanedInitial.pages[0]!.footnoteReservedHeight).toBeUndefined();
    expect(initialPage.footnoteReservedHeight).toBe(20);

    const returned = overlappingLayout([
      {
        kind: "paragraph",
        ...fragmentPosition,
        y: 70,
        fromLine: 0,
        toLine: 1,
      },
    ]);
    const returnedPage = returned.pages[0]!;
    const cleanedReturned = reflowFootnoteColumns({
      initialLayout: initial,
      runLayout: () => returned,
    });

    expect(cleanedReturned).not.toBe(returned);
    expect(cleanedReturned.pages[0]).not.toBe(returnedPage);
    expect(cleanedReturned.pages[0]!.footnoteReservedHeight).toBeUndefined();
    expect(returnedPage.footnoteReservedHeight).toBe(20);
  });
});
