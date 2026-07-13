import { describe, expect, test } from "bun:test";

import { createPaginator } from "./paginator";

const SIZE = { w: 800, h: 1000 };
const MARGINS = { top: 50, right: 50, bottom: 50, left: 50 };

describe("paginator mirrored margins", () => {
  test("swaps left and right margins on even physical pages", () => {
    const paginator = createPaginator({
      pageSize: SIZE,
      margins: { ...MARGINS, left: 90, right: 72 },
      mirrorMargins: true,
    });

    const first = paginator.getCurrentState();
    expect(first.page.margins.left).toBe(90);
    expect(paginator.getColumnX(0)).toBe(90);

    const second = paginator.forcePageBreak();
    expect(second.page.margins.left).toBe(72);
    expect(second.page.margins.right).toBe(90);
    expect(paginator.getColumnX(0)).toBe(72);
  });
});

describe("paginator even-page margins", () => {
  test("uses the section-specific clearance on even section pages", () => {
    const evenMargins = { ...MARGINS, top: 30, bottom: 35 };
    const paginator = createPaginator({
      pageSize: SIZE,
      margins: MARGINS,
      sectionEvenPageMargins: [evenMargins],
    });

    expect(paginator.getCurrentState().page.margins).toEqual(MARGINS);
    expect(paginator.forcePageBreak().page.margins).toEqual(evenMargins);
    expect(paginator.forcePageBreak().page.margins).toEqual(MARGINS);
  });
});

describe("paginator forcePageBreak", () => {
  test("two consecutive forcePageBreak calls preserve an explicit blank page", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.forcePageBreak();
    paginator.forcePageBreak();

    expect(paginator.pages.length).toBe(2);
  });

  test("coalesceBlankPage reuses an empty page with the active layout", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.forcePageBreak({ coalesceBlankPage: true });
    paginator.forcePageBreak({ coalesceBlankPage: true });

    expect(paginator.pages.length).toBe(1);
  });

  test("forcePageBreak after content followed by another forcePageBreak preserves a blank page", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    const state = paginator.getCurrentState();
    state.cursorY += 100;
    state.page.fragments.push({ kind: "paragraph" } as never);

    paginator.forcePageBreak();
    paginator.forcePageBreak();

    expect(paginator.pages.length).toBe(3);
  });

  test("forcePageBreak creates a fresh blank page after the active layout changes", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.forcePageBreak();

    const nextSize = { w: 600, h: 700 };
    const nextMargins = { top: 30, right: 40, bottom: 50, left: 60 };
    paginator.updatePageLayout(nextSize, nextMargins);
    const state = paginator.forcePageBreak({ coalesceBlankPage: true });

    expect(paginator.pages.length).toBe(2);
    expect(state.page.size).toEqual(nextSize);
    expect(state.page.margins).toEqual(nextMargins);
    expect(state.topMargin).toBe(nextMargins.top);
    expect(state.contentBottom).toBe(nextSize.h - nextMargins.bottom);
  });

  test("retargetCurrentBlankPage applies the active layout and section metadata", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.forcePageBreak();

    const nextSize = { w: 600, h: 700 };
    const nextMargins = { top: 30, right: 40, bottom: 50, left: 60 };
    paginator.updatePageLayout(nextSize, nextMargins);
    paginator.startSection(1);

    expect(paginator.retargetCurrentBlankPage()).toBe(true);
    expect(paginator.pages.length).toBe(1);

    const state = paginator.getCurrentState();
    expect(state.page.size).toEqual(nextSize);
    expect(state.page.margins).toEqual(nextMargins);
    expect(state.page.sectionIndex).toBe(1);
    expect(state.page.sectionPageNumber).toBe(1);
    expect(state.topMargin).toBe(nextMargins.top);
    expect(state.cursorY).toBe(nextMargins.top);
    expect(state.contentBottom).toBe(nextSize.h - nextMargins.bottom);
  });

  test("retargetCurrentBlankPage leaves nonblank pages unchanged", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    const state = paginator.getCurrentState();
    state.page.fragments.push({ kind: "paragraph" } as never);

    const nextSize = { w: 600, h: 700 };
    paginator.updatePageLayout(nextSize, MARGINS);
    paginator.startSection(1);

    expect(paginator.retargetCurrentBlankPage()).toBe(false);
    expect(state.page.size).toEqual(SIZE);
    expect(state.page.sectionIndex).toBe(0);
  });
});

describe("paginator block spacing", () => {
  test("collapses adjacent paragraph spacing to the larger side", () => {
    const paginator = createPaginator({
      pageSize: { w: 100, h: 100 },
      margins: { top: 10, right: 10, bottom: 10, left: 10 },
    });

    paginator.addFragment({ kind: "paragraph" } as never, 10, 0, 20);
    const result = paginator.addFragment({ kind: "paragraph" } as never, 10, 10, 0);

    expect(result.y).toBe(40);
  });

  test("does not carry trailing spacing to the top of a new page", () => {
    const paginator = createPaginator({
      pageSize: { w: 100, h: 100 },
      margins: { top: 10, right: 10, bottom: 10, left: 10 },
    });

    paginator.addFragment({ kind: "paragraph" } as never, 70, 0, 20);
    const result = paginator.addFragment({ kind: "paragraph" } as never, 20, 0, 0);

    expect(paginator.pages.length).toBe(2);
    expect(result.y).toBe(10);
  });

  test("preserves explicit spaceBefore at the top of a new page", () => {
    const paginator = createPaginator({
      pageSize: { w: 100, h: 100 },
      margins: { top: 10, right: 10, bottom: 10, left: 10 },
    });

    paginator.addFragment({ kind: "paragraph" } as never, 70, 0, 20);
    const result = paginator.addFragment({ kind: "paragraph" } as never, 20, 5, 0);

    expect(paginator.pages.length).toBe(2);
    expect(result.y).toBe(15);
  });
});
