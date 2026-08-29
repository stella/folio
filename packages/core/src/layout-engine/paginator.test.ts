import { describe, expect, test } from "bun:test";

import { createPaginator } from "./paginator";
import type { ParagraphFragment } from "./types";

const SIZE = { w: 800, h: 1000 };
const MARGINS = { top: 50, right: 50, bottom: 50, left: 50 };

const paragraphFragment = (blockId: string): ParagraphFragment => ({
  kind: "paragraph",
  blockId,
  x: 0,
  y: 0,
  width: 100,
  height: 20,
  fromLine: 0,
  toLine: 1,
});

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

  test("uses authored page-number parity after a restart", () => {
    const evenMargins = { ...MARGINS, top: 30, bottom: 35 };
    const paginator = createPaginator({
      pageSize: SIZE,
      margins: MARGINS,
      pageNumbering: { type: "restart", start: 2 },
      sectionEvenPageMargins: [evenMargins],
    });

    expect(paginator.getCurrentState().page.margins).toEqual(evenMargins);
  });

  test("retargets a coalesced section page with its restarted parity", () => {
    const evenMargins = { ...MARGINS, top: 30, bottom: 35 };
    const paginator = createPaginator({
      pageSize: SIZE,
      margins: MARGINS,
      sectionEvenPageMargins: [undefined, evenMargins],
    });
    paginator.forcePageBreak();
    paginator.startSection(1, { type: "restart", start: 2 });

    const page = paginator.forcePageBreak({ coalesceBlankPage: true }).page;

    expect(paginator.pages).toHaveLength(1);
    expect(page.logicalNumber).toBe(2);
    expect(page.margins).toEqual(evenMargins);
  });
});

describe("paginator logical page numbers", () => {
  test("keeps physical, logical, and section page numbers distinct", () => {
    const paginator = createPaginator({
      pageSize: SIZE,
      margins: MARGINS,
      pageNumbering: { type: "restart", start: 13 },
    });

    const first = paginator.getCurrentState().page;
    const second = paginator.forcePageBreak().page;
    paginator.startSection(1, { type: "continue" });
    const continued = paginator.forcePageBreak().page;
    paginator.startSection(2, { type: "restart", start: 4 });
    const restarted = paginator.forcePageBreak().page;

    expect([first.number, second.number, continued.number, restarted.number]).toEqual([1, 2, 3, 4]);
    expect([
      first.logicalNumber,
      second.logicalNumber,
      continued.logicalNumber,
      restarted.logicalNumber,
    ]).toEqual([13, 14, 15, 4]);
    expect([
      first.sectionPageNumber,
      second.sectionPageNumber,
      continued.sectionPageNumber,
      restarted.sectionPageNumber,
    ]).toEqual([1, 2, 1, 1]);
  });

  test("consumes a section restart on its first shared-page fragment", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.addFragment(paragraphFragment("outgoing"), 20);

    paginator.startSection(1, { type: "restart", start: 2 });
    paginator.addFragment(paragraphFragment("incoming"), 20);
    const shared = paginator.pages[0];
    const following = paginator.forcePageBreak().page;

    expect(shared?.fragments.map(({ blockId }) => blockId)).toEqual(["outgoing", "incoming"]);
    expect(shared).toMatchObject({ logicalNumber: 1, sectionIndex: 0, sectionPageNumber: 1 });
    expect(following).toMatchObject({ logicalNumber: 3, sectionIndex: 1, sectionPageNumber: 1 });
  });

  test("consumes a section restart on its first unflowed shared-page fragment", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.addFragment(paragraphFragment("outgoing"), 20);

    paginator.startSection(1, { type: "restart", start: 2 });
    paginator.addUnflowedFragment({
      kind: "image",
      blockId: "anchored-incoming",
      x: 100,
      y: 100,
      width: 20,
      height: 20,
      isAnchored: true,
    });
    const following = paginator.forcePageBreak().page;

    expect(paginator.pages[0]).toMatchObject({ logicalNumber: 1, sectionIndex: 0 });
    expect(following).toMatchObject({ logicalNumber: 3, sectionIndex: 1, sectionPageNumber: 1 });
  });

  test("defers a restart when the first section fragment advances to a fresh page", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.addFragment(paragraphFragment("outgoing"), 900);

    paginator.startSection(1, { type: "restart", start: 2 });
    paginator.addFragment(paragraphFragment("incoming"), 20);

    expect(paginator.pages).toHaveLength(2);
    expect(paginator.pages[0]).toMatchObject({ logicalNumber: 1, sectionIndex: 0 });
    expect(paginator.pages[1]).toMatchObject({ logicalNumber: 2, sectionIndex: 1 });
  });

  test("defers a restart when the first section fragment follows a forced break", () => {
    const paginator = createPaginator({ pageSize: SIZE, margins: MARGINS });
    paginator.addFragment(paragraphFragment("outgoing"), 20);

    paginator.startSection(1, { type: "restart", start: 2 });
    paginator.forcePageBreak();
    paginator.addFragment(paragraphFragment("incoming"), 20);

    expect(paginator.pages.map(({ logicalNumber }) => logicalNumber)).toEqual([1, 2]);
    expect(paginator.pages[0]?.sectionIndex).toBe(0);
    expect(paginator.pages[1]).toMatchObject({ sectionIndex: 1, sectionPageNumber: 1 });
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
    paginator.addUnflowedFragment(paragraphFragment("content"));

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
    paginator.addUnflowedFragment(paragraphFragment("content"));

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
