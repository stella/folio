import { describe, expect, test } from "bun:test";

import { layoutDocument } from "./index";
import type {
  ColumnBreakBlock,
  FlowBlock,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
  SectionBreakBlock,
} from "./types";

function paragraph(
  id: string,
  height: number,
): { block: ParagraphBlock; measure: ParagraphMeasure } {
  return {
    block: {
      kind: "paragraph",
      id,
      pmStart: 0,
      pmEnd: 0,
      runs: [{ kind: "text", text: id }],
      attrs: {},
    },
    measure: {
      kind: "paragraph",
      lines: [
        {
          fromRun: 0,
          fromChar: 0,
          toRun: 0,
          toChar: 0,
          width: 100,
          ascent: 10,
          descent: 3,
          lineHeight: height,
        },
      ],
      totalHeight: height,
    },
  };
}

type EmptyParagraphOptions = {
  id: string;
  attrs?: ParagraphBlock["attrs"];
  lineHeight?: number;
};

function emptyParagraph({ id, attrs = {}, lineHeight = 20 }: EmptyParagraphOptions): {
  block: ParagraphBlock;
  measure: ParagraphMeasure;
} {
  return {
    block: {
      kind: "paragraph",
      id,
      pmStart: 0,
      pmEnd: 0,
      runs: [],
      attrs,
    },
    measure: {
      kind: "paragraph",
      lines: [
        {
          fromRun: 0,
          fromChar: 0,
          toRun: 0,
          toChar: 0,
          width: 0,
          ascent: 0,
          descent: 0,
          lineHeight,
        },
      ],
      totalHeight: lineHeight,
    },
  };
}

describe("continuous section break geometry", () => {
  test("balances a short paragraph-only multi-column section", () => {
    const intro = paragraph("intro", 100);
    const firstBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "first-break",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const first = paragraph("first", 100);
    const second = paragraph("second", 100);
    const third = paragraph("third", 100);
    const fourth = paragraph("fourth", 100);
    const secondBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "second-break",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      columns: { count: 2, gap: 20 },
    };
    const outro = paragraph("outro", 100);
    const blocks: FlowBlock[] = [
      intro.block,
      firstBreak,
      first.block,
      second.block,
      third.block,
      fourth.block,
      secondBreak,
      outro.block,
    ];
    const measures = [
      intro.measure,
      { kind: "sectionBreak" },
      first.measure,
      second.measure,
      third.measure,
      fourth.measure,
      { kind: "sectionBreak" },
      outro.measure,
    ] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalPageSize: { w: 800, h: 1000 },
      finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    expect(result.pages).toHaveLength(1);
    const page = result.pages[0];
    const firstColumn = page?.fragments.filter(
      (fragment) => fragment.kind === "paragraph" && ["first", "second"].includes(fragment.blockId),
    );
    const secondColumn = page?.fragments.filter(
      (fragment) => fragment.kind === "paragraph" && ["third", "fourth"].includes(fragment.blockId),
    );
    expect(firstColumn?.map(({ x }) => x)).toEqual([50, 50]);
    expect(secondColumn?.map(({ x }) => x)).toEqual([410, 410]);
    expect(secondColumn?.map(({ y }) => y)).toEqual([150, 250]);
  });

  test("a final multi-column section keeps a forced column break on the current page", () => {
    const first = paragraph("a", 100);
    const sectionBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "sb",
      type: "continuous",
    };
    const firstColumn = paragraph("b", 100);
    const columnBreak: ColumnBreakBlock = { kind: "columnBreak", id: "cb" };
    const secondColumn = paragraph("c", 100);
    const blocks: FlowBlock[] = [
      first.block,
      sectionBreak,
      firstColumn.block,
      columnBreak,
      secondColumn.block,
    ];
    const measures = [
      first.measure,
      { kind: "sectionBreak" },
      firstColumn.measure,
      { kind: "columnBreak" },
      secondColumn.measure,
    ] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalColumns: { count: 2, gap: 20, widths: [200, 300], gaps: [100] },
    });

    expect(result.pages).toHaveLength(1);
    const secondColumnFragment = result.pages[0]?.fragments.find(
      (fragment) => fragment.kind === "paragraph" && fragment.blockId === "c",
    );
    expect(secondColumnFragment?.x).toBe(350);
    expect(secondColumnFragment?.width).toBe(300);
  });

  test("an omitted transition defaults to a new page at its own boundary", () => {
    const first = paragraph("first", 100);
    const firstBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "first-break",
      type: "continuous",
    };
    const second = paragraph("second", 100);
    const omittedTypeBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "omitted-type-break",
    };
    const third = paragraph("third", 100);
    const blocks: FlowBlock[] = [
      first.block,
      firstBreak,
      second.block,
      omittedTypeBreak,
      third.block,
    ];
    const measures = [
      first.measure,
      { kind: "sectionBreak" },
      second.measure,
      { kind: "sectionBreak" },
      third.measure,
    ] satisfies Measure[];

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]?.fragments.map((fragment) => fragment.blockId)).toEqual([
      "first",
      "second",
    ]);
    expect(result.pages[1]?.fragments.map((fragment) => fragment.blockId)).toEqual(["third"]);
  });

  test("a hard break reuses the blank page opened by a preceding section boundary", () => {
    const cover = paragraph("cover", 100);
    const spacer = emptyParagraph({ id: "spacer" });
    const carrier = emptyParagraph({
      id: "carrier",
      attrs: { suppressEmptyParagraphHeight: true },
      lineHeight: 0,
    });
    const body = paragraph("body", 100);
    body.block.attrs = { renderedPageBreakBefore: true };
    const blocks: FlowBlock[] = [
      cover.block,
      { kind: "sectionBreak", id: "cover-end" },
      { kind: "sectionBreak", id: "continuous-marker", type: "continuous" },
      spacer.block,
      { kind: "pageBreak", id: "redundant-page-break" },
      carrier.block,
      body.block,
    ];
    const measures = [
      cover.measure,
      { kind: "sectionBreak" },
      { kind: "sectionBreak" },
      spacer.measure,
      { kind: "pageBreak" },
      carrier.measure,
      body.measure,
    ] satisfies Measure[];

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    expect(result.pages).toHaveLength(2);
    expect(result.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual([
      "spacer",
      "carrier",
      "body",
    ]);
  });

  test("a hard-break page is not coalesced through a continuous section", () => {
    const first = paragraph("first", 100);
    const body = paragraph("body", 100);
    const result = layoutDocument(
      [
        first.block,
        { kind: "pageBreak", id: "first-hard-break" },
        { kind: "sectionBreak", id: "continuous-marker", type: "continuous" },
        { kind: "pageBreak", id: "second-hard-break" },
        body.block,
      ],
      [
        first.measure,
        { kind: "pageBreak" },
        { kind: "sectionBreak" },
        { kind: "pageBreak" },
        body.measure,
      ],
      {
        pageSize: { w: 800, h: 1000 },
        margins: { top: 50, right: 50, bottom: 50, left: 50 },
      },
    );

    expect(result.pages).toHaveLength(3);
    expect(result.pages[2]?.fragments.map(({ blockId }) => blockId)).toEqual(["body"]);
  });

  test("a column break consumes section-break coalescence before a hard break", () => {
    const cover = paragraph("cover", 100);
    const body = paragraph("body", 100);
    const result = layoutDocument(
      [
        cover.block,
        { kind: "sectionBreak", id: "cover-end" },
        { kind: "columnBreak", id: "column-break" },
        { kind: "pageBreak", id: "page-break" },
        body.block,
      ],
      [
        cover.measure,
        { kind: "sectionBreak" },
        { kind: "columnBreak" },
        { kind: "pageBreak" },
        body.measure,
      ],
      {
        pageSize: { w: 800, h: 1000 },
        margins: { top: 50, right: 50, bottom: 50, left: 50 },
        finalColumns: { count: 2, gap: 20 },
      },
    );

    expect(result.pages).toHaveLength(3);
    expect(result.pages[1]?.fragments).toHaveLength(0);
    expect(result.pages[2]?.fragments.map(({ blockId, x }) => ({ blockId, x }))).toEqual([
      { blockId: "body", x: 50 },
    ]);
  });

  test("visible empty-paragraph decoration consumes section-break coalescence", () => {
    const cover = paragraph("cover", 100);
    const decorated = emptyParagraph({ id: "decorated", attrs: { shading: "#FFFF00" } });
    const body = paragraph("body", 100);
    const result = layoutDocument(
      [
        cover.block,
        { kind: "sectionBreak", id: "cover-end" },
        decorated.block,
        { kind: "pageBreak", id: "hard-break" },
        body.block,
      ],
      [
        cover.measure,
        { kind: "sectionBreak" },
        decorated.measure,
        { kind: "pageBreak" },
        body.measure,
      ],
      {
        pageSize: { w: 800, h: 1000 },
        margins: { top: 50, right: 50, bottom: 50, left: 50 },
      },
    );

    expect(result.pages).toHaveLength(3);
    expect(result.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual(["decorated"]);
    expect(result.pages[2]?.fragments.map(({ blockId }) => blockId)).toEqual(["body"]);
  });

  test("does not apply final section config to a preceding omitted transition", () => {
    const first = paragraph("first", 100);
    const second = paragraph("second", 100);
    const result = layoutDocument(
      [first.block, { kind: "sectionBreak", id: "break" }, second.block],
      [first.measure, { kind: "sectionBreak" }, second.measure],
      {
        pageSize: { w: 800, h: 1000 },
        margins: { top: 50, right: 50, bottom: 50, left: 50 },
      },
    );

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]?.fragments.map(({ blockId }) => blockId)).toEqual(["first"]);
    expect(result.pages[1]?.fragments.map(({ blockId }) => blockId)).toEqual(["second"]);
  });

  test("content after a continuous column section resumes below its tallest column", () => {
    const intro = paragraph("intro", 100);
    const firstBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "first-break",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const first = paragraph("first", 100);
    const second = paragraph("second", 100);
    const columnBreak: ColumnBreakBlock = { kind: "columnBreak", id: "column-break" };
    const third = paragraph("third", 50);
    const secondBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "second-break",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      columns: { count: 2, gap: 20 },
    };
    const outro = paragraph("outro", 100);
    const blocks: FlowBlock[] = [
      intro.block,
      firstBreak,
      first.block,
      second.block,
      columnBreak,
      third.block,
      secondBreak,
      outro.block,
    ];
    const measures = [
      intro.measure,
      { kind: "sectionBreak" },
      first.measure,
      second.measure,
      { kind: "columnBreak" },
      third.measure,
      { kind: "sectionBreak" },
      outro.measure,
    ] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalPageSize: { w: 800, h: 1000 },
      finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    const outroFragment = result.pages[0]?.fragments.find(
      (fragment) => fragment.kind === "paragraph" && fragment.blockId === "outro",
    );
    expect(outroFragment?.y).toBe(350);
  });

  test("current page keeps old geometry and overflow page picks up new geometry", () => {
    const first = paragraph("a", 200);
    const sectionBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "sb",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const second = paragraph("b", 200);
    const third = paragraph("c", 800);

    const blocks: FlowBlock[] = [first.block, sectionBreak, second.block, third.block];
    const measures = [
      first.measure,
      { kind: "sectionBreak" },
      second.measure,
      third.measure,
    ] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalPageSize: { w: 1200, h: 700 },
      finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    expect(result.pages[0]?.size.w).toBe(800);
    const lastPage = result.pages.at(-1);
    expect(lastPage?.size).toEqual({ w: 1200, h: 700 });
  });

  test("orientation-changing continuous break is promoted to a page break", () => {
    // Regression (eigenpal/docx-editor#841): a `continuous` break normally
    // defers the new geometry, but a break that changes page size/orientation
    // cannot share a physical sheet with the preceding section. Word and
    // LibreOffice promote it to a page break; match that.
    const first = paragraph("a", 200);
    // The break block describes the section it terminates (the portrait first
    // section, which sets the initial page geometry); the next/body section is
    // landscape via `finalPageSize`.
    const sectionBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "sb",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const second = paragraph("b", 200);

    const blocks: FlowBlock[] = [first.block, sectionBreak, second.block];
    const measures = [first.measure, { kind: "sectionBreak" }, second.measure] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalPageSize: { w: 1000, h: 800 },
      finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    // "b" must land on a NEW page that already carries the landscape geometry,
    // not share the portrait page with "a" (the pre-fix behavior).
    expect(result.pages.length).toBeGreaterThanOrEqual(2);
    const pageWithB = result.pages.find((page) =>
      page.fragments.some((f) => f.kind === "paragraph" && f.blockId === "b"),
    );
    expect(pageWithB?.size).toEqual({ w: 1000, h: 800 });
    const pageWithA = result.pages.find((page) =>
      page.fragments.some((f) => f.kind === "paragraph" && f.blockId === "a"),
    );
    expect(pageWithA?.size).toEqual({ w: 800, h: 1000 });
    // The promoted break starts the next section, so the new page carries the
    // next section's index (and thus its header/footer references).
    expect(pageWithA?.sectionIndex).toBe(0);
    expect(pageWithB?.sectionIndex).toBe(1);
  });

  test("a leading size-changing continuous break does not strand a blank page", () => {
    // With no content laid out yet there is no sheet to share, so the break
    // defers instead of materializing a blank page just to compare geometry;
    // the first content opens directly on a new-geometry page.
    const sectionBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "sb",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const body = paragraph("b", 200);

    const blocks: FlowBlock[] = [sectionBreak, body.block];
    const measures = [{ kind: "sectionBreak" }, body.measure] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalPageSize: { w: 1000, h: 800 },
      finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    // Exactly one page, carrying the body — no empty leading page.
    expect(result.pages).toHaveLength(1);
    expect(result.pages[0]?.size).toEqual({ w: 1000, h: 800 });
    expect(result.pages[0]?.sectionIndex).toBe(1);
    expect(
      result.pages[0]?.fragments.some((f) => f.kind === "paragraph" && f.blockId === "b"),
    ).toBe(true);
  });

  test("a promoted continuous break reuses an already blank current page", () => {
    const first = paragraph("a", 200);
    const sectionBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "sb",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const body = paragraph("b", 200);

    const blocks: FlowBlock[] = [
      first.block,
      { kind: "pageBreak", id: "pb" },
      sectionBreak,
      body.block,
    ];
    const measures = [
      first.measure,
      { kind: "pageBreak" },
      { kind: "sectionBreak" },
      body.measure,
    ] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalPageSize: { w: 1000, h: 800 },
      finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    expect(result.pages).toHaveLength(2);
    expect(
      result.pages[0]?.fragments.some((f) => f.kind === "paragraph" && f.blockId === "a"),
    ).toBe(true);
    expect(result.pages[1]?.size).toEqual({ w: 1000, h: 800 });
    expect(result.pages[1]?.sectionIndex).toBe(1);
    expect(
      result.pages[1]?.fragments.some((f) => f.kind === "paragraph" && f.blockId === "b"),
    ).toBe(true);
  });

  test("a same-size continuous break starts the next section on later pages", () => {
    const first = paragraph("a", 700);
    const sectionBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "sb",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const second = paragraph("b", 150);
    const third = paragraph("c", 700);

    const blocks: FlowBlock[] = [first.block, sectionBreak, second.block, third.block];
    const measures = [
      first.measure,
      { kind: "sectionBreak" },
      second.measure,
      third.measure,
    ] as never;

    const result = layoutDocument(blocks, measures, {
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
      finalPageSize: { w: 800, h: 1000 },
      finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
    });

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]?.sectionIndex).toBe(0);
    expect(result.pages[1]?.sectionIndex).toBe(1);
    expect(result.pages[1]?.sectionPageNumber).toBe(1);
    expect(
      result.pages[1]?.fragments.some((f) => f.kind === "paragraph" && f.blockId === "c"),
    ).toBe(true);
  });

  test("advances a restart past incoming content retained on the shared page", () => {
    const first = paragraph("outgoing", 700);
    const sectionBreak: SectionBreakBlock = {
      kind: "sectionBreak",
      id: "sb",
      type: "continuous",
      pageSize: { w: 800, h: 1000 },
      margins: { top: 50, right: 50, bottom: 50, left: 50 },
    };
    const shared = paragraph("shared-incoming", 150);
    const next = paragraph("next-incoming", 700);

    const result = layoutDocument(
      [first.block, sectionBreak, shared.block, next.block],
      [first.measure, { kind: "sectionBreak" }, shared.measure, next.measure] as never,
      {
        pageSize: { w: 800, h: 1000 },
        margins: { top: 50, right: 50, bottom: 50, left: 50 },
        finalPageSize: { w: 800, h: 1000 },
        finalMargins: { top: 50, right: 50, bottom: 50, left: 50 },
        finalPageNumbering: { type: "restart", start: 2 },
      },
    );

    expect(result.pages.map(({ logicalNumber }) => logicalNumber)).toEqual([1, 3]);
    expect(result.pages[0]).toMatchObject({ sectionIndex: 0, sectionPageNumber: 1 });
    expect(result.pages[1]).toMatchObject({ sectionIndex: 1, sectionPageNumber: 1 });
    expect(
      result.pages[0]?.fragments.some(
        (fragment) => fragment.kind === "paragraph" && fragment.blockId === "shared-incoming",
      ),
    ).toBe(true);
  });
});
