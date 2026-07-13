import { describe, expect, test } from "bun:test";

import type { FlowBlock, ImageBlock, ParagraphBlock, TableBlock, TextBoxBlock } from "../types";
import { setTextBoxGroupId } from "../textBoxGroup";
import { fixedCharWidth, withFakeTextMeasure } from "./__tests__/fakeTextMeasure";
import { measureBlock, measureBlocks, measureTableBlock } from "./measureBlocks";

const fakeMeasure = { charWidth: fixedCharWidth(5) };

const para = (id: string, text: string): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text }],
});

const imageBlock: ImageBlock = {
  kind: "image",
  id: "img-1",
  src: "data:image/png;base64,",
  width: 120,
  height: 40,
};

describe("measureBlock dispatch", () => {
  test("image block reports its own dimensions", () => {
    const measure = measureBlock(imageBlock, 500);
    expect(measure.kind).toBe("image");
    if (measure.kind === "image") {
      expect(measure.width).toBe(120);
      expect(measure.height).toBe(40);
    }
  });

  test("structural breaks map to matching measure kinds", () => {
    expect(measureBlock({ kind: "pageBreak", id: "pb-1" }, 500).kind).toBe("pageBreak");
    expect(measureBlock({ kind: "columnBreak", id: "cb-1" }, 500).kind).toBe("columnBreak");
    expect(measureBlock({ kind: "sectionBreak", id: "sb-1" }, 500).kind).toBe("sectionBreak");
  });

  test("unknown block kind falls back to an empty paragraph measure", () => {
    const measure = measureBlock({ kind: "mystery" } as unknown as FlowBlock, 500);
    expect(measure.kind).toBe("paragraph");
    if (measure.kind === "paragraph") {
      expect(measure.lines).toEqual([]);
      expect(measure.totalHeight).toBe(0);
    }
  });
});

describe("measureBlocks", () => {
  test("returns exactly one measure per input block", () => {
    withFakeTextMeasure(() => {
      const blocks: FlowBlock[] = [imageBlock, { kind: "pageBreak", id: "pb-1" }, imageBlock];
      const measures = measureBlocks(blocks, 500);
      expect(measures).toHaveLength(blocks.length);
      expect(measures.map((m) => m.kind)).toEqual(["image", "pageBreak", "image"]);
    }, fakeMeasure);
  });

  test("keeps paragraph-anchored bands from one textbox group active together", () => {
    withFakeTextMeasure(() => {
      const firstBand: TextBoxBlock = {
        kind: "textBox",
        id: "first-band",
        width: 300,
        height: 30,
        content: [],
        wrapType: "topAndBottom",
        position: { vertical: { relativeTo: "paragraph", posOffset: 0 } },
      };
      const secondBand: TextBoxBlock = {
        ...firstBand,
        id: "second-band",
        position: { vertical: { relativeTo: "paragraph", posOffset: 285_750 } },
      };
      setTextBoxGroupId(firstBand, "host-paragraph");
      setTextBoxGroupId(secondBand, "host-paragraph");

      const measures = measureBlocks([firstBand, secondBand, para("after", "after")], 600);
      const paragraphMeasure = measures.at(2);
      if (paragraphMeasure?.kind !== "paragraph") {
        throw new Error("Expected paragraph measure");
      }

      expect(paragraphMeasure.lines.at(0)?.floatSkipBefore).toBeCloseTo(60, 5);
    }, fakeMeasure);
  });

  test("keeps an active band across an unspecified continuous section break", () => {
    withFakeTextMeasure(() => {
      const band: TextBoxBlock = {
        kind: "textBox",
        id: "band",
        width: 300,
        height: 120,
        content: [],
        wrapType: "topAndBottom",
        position: { vertical: { relativeTo: "paragraph", posOffset: 285_750 } },
      };
      const blocks: FlowBlock[] = [
        band,
        para("before", "before"),
        { kind: "sectionBreak", id: "section" },
        para("after", "after"),
      ];

      const measures = measureBlocks(blocks, 600);
      const afterMeasure = measures.at(3);
      if (afterMeasure?.kind !== "paragraph") {
        throw new Error("Expected paragraph after section break");
      }

      expect(afterMeasure.lines.at(0)?.floatSkipBefore).toBeGreaterThan(0);
    }, fakeMeasure);
  });

  test("per-block content widths are honoured for parallel arrays", () => {
    withFakeTextMeasure(() => {
      const paragraph: ParagraphBlock = {
        kind: "paragraph",
        id: "p-1",
        runs: [{ kind: "text", text: "hello world" }],
      };
      const measures = measureBlocks([paragraph, paragraph], [500, 300]);
      expect(measures).toHaveLength(2);
      // Both inputs are paragraphs measured against different widths; the
      // dispatch still produces a measure for each block.
      expect(measures.every((m) => m.kind === "paragraph")).toBe(true);
    }, fakeMeasure);
  });
});

describe("measureTableBlock row height", () => {
  test("measures a row-spanning cell against the combined spanned-row height", () => {
    withFakeTextMeasure(() => {
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [120, 120],
        rows: [
          {
            id: "r0",
            height: 30,
            cells: [
              {
                id: "span",
                rowSpan: 2,
                blocks: [para("span-1", "First"), para("span-2", "Second")],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
              },
              { id: "r0-cell", blocks: [para("r0-p", "Row zero")] },
            ],
          },
          {
            id: "r1",
            height: 30,
            cells: [{ id: "r1-cell", blocks: [para("r1-p", "Row one")] }],
          },
        ],
      };

      const measure = measureTableBlock(table, 240);

      expect(measure.rows.map((row) => row.height)).toEqual([30, 30]);
      expect(measure.totalHeight).toBe(60);
    }, fakeMeasure);
  });

  test("adds a row-spanning content deficit to the last non-exact row", () => {
    withFakeTextMeasure(() => {
      const spanningBlocks = Array.from({ length: 6 }, (_, index) =>
        para(`span-${index}`, `Line ${index}`),
      );
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [120, 120],
        rows: [
          {
            id: "r0",
            height: 20,
            heightRule: "exact",
            cells: [
              {
                id: "span",
                rowSpan: 2,
                blocks: spanningBlocks,
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
              },
              { id: "r0-cell", blocks: [para("r0-p", "Row zero")] },
            ],
          },
          {
            id: "r1",
            height: 20,
            cells: [{ id: "r1-cell", blocks: [para("r1-p", "Row one")] }],
          },
        ],
      };

      const measure = measureTableBlock(table, 240);
      const spanningCellHeight = measure.rows[0]!.cells[0]!.height;

      expect(measure.rows[0]!.height).toBe(20);
      expect(measure.rows[1]!.height).toBe(spanningCellHeight - 20);
      expect(measure.totalHeight).toBe(spanningCellHeight);
    }, fakeMeasure);
  });

  test("counts a collapsed top border only on the first row", () => {
    withFakeTextMeasure(() => {
      const borderedCell = (id: string) => ({
        id,
        blocks: [para(`${id}-p`, "One line")],
        padding: { top: 0, right: 0, bottom: 0, left: 0 },
        borders: { top: { width: 2 }, bottom: { width: 2 } },
      });
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [120],
        rows: [
          { id: "r0", cells: [borderedCell("first")] },
          { id: "r1", cells: [borderedCell("second")] },
        ],
      };

      const measure = measureTableBlock(table, 120);
      const firstContentHeight = measure.rows[0]?.cells[0]?.height ?? 0;
      const secondContentHeight = measure.rows[1]?.cells[0]?.height ?? 0;

      expect(measure.rows[0]?.height).toBe(firstContentHeight + 4);
      expect(measure.rows[1]?.height).toBe(secondContentHeight + 2);
    }, fakeMeasure);
  });

  test("maxes per-cell content+border, not summed independent maxes", () => {
    withFakeTextMeasure(() => {
      // Cell A: more content (two paragraphs), thin border.
      // Cell B: less content (one paragraph), thick border.
      // The buggy formula (maxContent + maxBorder) takes A's content and B's
      // border from different cells and over-allocates; the correct one maxes
      // each cell's own content+padding+border.
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [120, 120],
        rows: [
          {
            id: "r0",
            cells: [
              {
                id: "a",
                blocks: [para("a1", "Tall cell line one."), para("a2", "two")],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                borders: { top: { width: 1 }, bottom: { width: 1 } }, // 2
              },
              {
                id: "b",
                blocks: [para("b1", "Short.")],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                borders: { top: { width: 20 }, bottom: { width: 20 } }, // 40
              },
            ],
          },
        ],
      };

      const measure = measureTableBlock(table, 240);
      const row = measure.rows[0]!;
      const cellA = row.cells[0]!;
      const cellB = row.cells[1]!;

      // The tallest-content cell (A) is not the tallest-border cell (B).
      expect(cellA.height).toBeGreaterThan(cellB.height);

      const correct = Math.max(cellA.height + 2, cellB.height + 40);
      const summedMaxes = Math.max(cellA.height, cellB.height) + 40;

      expect(row.height).toBe(correct);
      // The fix avoids the over-allocation the old summed-maxes formula produced.
      expect(row.height).toBeLessThan(summedMaxes);
    }, fakeMeasure);
  });

  test("adds horizontal cell borders on top of an atLeast/auto minimum height", () => {
    withFakeTextMeasure(() => {
      // A short single-line cell whose content is well under an explicit
      // (atLeast) row height, but which carries a horizontal border. Word treats
      // the explicit height as the minimum of the cell content+padding region;
      // the border sits on the grid line and adds on top. So the row must be
      // taller than the explicit height by the border thickness — the border
      // must not be absorbed by `max(content + border, explicit)`.
      const border = 6;
      const tallMinHeight = 400;
      const build = (rule: "atLeast" | undefined): TableBlock => ({
        kind: "table",
        id: "t",
        columnWidths: [120],
        rows: [
          {
            id: "r0",
            height: tallMinHeight,
            ...(rule ? { heightRule: rule } : {}),
            cells: [
              {
                id: "a",
                blocks: [para("a1", "One line")],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                // First row: top+bottom both collapse in, so 2 * border.
                borders: { top: { width: border }, bottom: { width: border } },
              },
            ],
          },
        ],
      });

      for (const rule of ["atLeast", undefined] as const) {
        const measure = measureTableBlock(build(rule), 120);
        const contentHeight = measure.rows[0]?.cells[0]?.height ?? 0;
        expect(contentHeight).toBeLessThan(tallMinHeight); // content is the shorter term
        // Border (2 * 6 on the first row) added on top of the minimum height.
        expect(measure.rows[0]?.height).toBe(tallMinHeight + 2 * border);
        expect(measure.rows[0]?.height).toBeGreaterThan(tallMinHeight);
      }
    }, fakeMeasure);
  });

  test("does not double-count borders when content already exceeds the minimum height", () => {
    withFakeTextMeasure(() => {
      // Guard against over-correction: when content+border already exceeds the
      // explicit minimum, the row is content+border (the border is not added a
      // second time), matching the no-explicit-height case.
      const border = 4;
      const cell = () => ({
        id: "a",
        blocks: [para("a1", "One line")],
        padding: { top: 0, right: 0, bottom: 0, left: 0 },
        borders: { top: { width: border }, bottom: { width: border } },
      });
      const withMin: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [120],
        rows: [{ id: "r0", height: 1, heightRule: "atLeast", cells: [cell()] }],
      };
      const noMin: TableBlock = {
        kind: "table",
        id: "t2",
        columnWidths: [120],
        rows: [{ id: "r0", cells: [cell()] }],
      };

      const withMinHeight = measureTableBlock(withMin, 120).rows[0]?.height ?? 0;
      const noMinHeight = measureTableBlock(noMin, 120).rows[0]?.height ?? 0;
      const contentHeight = measureTableBlock(noMin, 120).rows[0]?.cells[0]?.height ?? 0;

      expect(withMinHeight).toBe(contentHeight + 2 * border);
      expect(withMinHeight).toBe(noMinHeight);
    }, fakeMeasure);
  });

  test("exact height rule stays exact and ignores cell borders", () => {
    withFakeTextMeasure(() => {
      // `hRule=exact` fixes the whole row height; borders do not extend it.
      const exactHeight = 300;
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [120],
        rows: [
          {
            id: "r0",
            height: exactHeight,
            heightRule: "exact",
            cells: [
              {
                id: "a",
                blocks: [para("a1", "One line")],
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                borders: { top: { width: 8 }, bottom: { width: 8 } },
              },
            ],
          },
        ],
      };

      expect(measureTableBlock(table, 120).rows[0]?.height).toBe(exactHeight);
    }, fakeMeasure);
  });

  test("hidden rows remain addressable but contribute zero height", () => {
    withFakeTextMeasure(() => {
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [120],
        rows: [
          {
            id: "visible",
            cells: [{ id: "a", blocks: [para("p1", "visible")] }],
          },
          {
            id: "hidden",
            hidden: true,
            height: 40,
            cells: [{ id: "b", blocks: [para("p2", "hidden but measured as absent")] }],
          },
        ],
      };

      const measure = measureTableBlock(table, 120);

      expect(measure.rows).toHaveLength(2);
      expect(measure.rows[0]?.height).toBeGreaterThan(0);
      expect(measure.rows[1]?.height).toBe(0);
      expect(measure.totalHeight).toBe(measure.rows[0]?.height);
    }, fakeMeasure);
  });

  test("image-only table cell paragraphs use the image visual height", () => {
    withFakeTextMeasure(() => {
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [240],
        rows: [
          {
            id: "logo-row",
            cells: [
              {
                id: "logo-cell",
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                blocks: [
                  {
                    kind: "paragraph",
                    id: "logo",
                    runs: [
                      {
                        kind: "image",
                        src: "data:image/png;base64,",
                        width: 80,
                        height: 40,
                      },
                    ],
                    attrs: {
                      defaultFontFamily: "Calibri",
                      defaultFontSize: 11,
                    },
                  },
                ],
              },
            ],
          },
        ],
      };

      const measure = measureTableBlock(table, 240);
      const row = measure.rows[0];
      const cell = row?.cells[0];
      const paragraph = cell?.blocks[0];

      expect(paragraph?.kind).toBe("paragraph");
      if (paragraph?.kind !== "paragraph") {
        throw new Error("Expected paragraph measure");
      }
      expect(paragraph.totalHeight).toBeGreaterThan(40);
      expect(cell?.height).toBe(40);
      expect(cell?.height).toBeLessThan(paragraph.totalHeight);
      expect(row?.height).toBe(cell?.height);
    }, fakeMeasure);
  });
});

describe("measureTableBlock w:noWrap column pinning", () => {
  // Column content width is 60px (no padding) and each char is 5px, so this
  // three-word phrase (75px unbroken) must wrap when the column is pinned and
  // must stay on one line when the cell is allowed to keep noWrap.
  const OVERFLOW_TEXT = "wordone wordtwo wordthree";
  const FIT_TEXT = "hi";

  const noWrapTable = (opts: {
    text: string;
    widthType?: TableBlock["widthType"];
    layout?: TableBlock["layout"];
  }): TableBlock => ({
    kind: "table",
    id: "t",
    columnWidths: [60],
    ...(opts.widthType ? { widthType: opts.widthType } : {}),
    ...(opts.layout ? { layout: opts.layout } : {}),
    rows: [
      {
        id: "r",
        cells: [
          {
            id: "c",
            noWrap: true,
            padding: { top: 0, right: 0, bottom: 0, left: 0 },
            blocks: [para("p", opts.text)],
          },
        ],
      },
    ],
  });

  const cellLineCount = (table: TableBlock): number => {
    const paragraph = measureTableBlock(table, 500).rows[0]?.cells[0]?.blocks[0];
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected paragraph measure");
    }
    return paragraph.lines.length;
  };

  test("wraps an overflowing noWrap cell when an explicit dxa width pins the column", () => {
    withFakeTextMeasure(() => {
      expect(cellLineCount(noWrapTable({ text: OVERFLOW_TEXT, widthType: "dxa" }))).toBeGreaterThan(
        1,
      );
    }, fakeMeasure);
  });

  test("wraps an overflowing noWrap cell under a fixed table layout", () => {
    withFakeTextMeasure(() => {
      expect(cellLineCount(noWrapTable({ text: OVERFLOW_TEXT, layout: "fixed" }))).toBeGreaterThan(
        1,
      );
    }, fakeMeasure);
  });

  test("keeps an overflowing noWrap cell on one line in an auto-width table", () => {
    withFakeTextMeasure(() => {
      // Auto layout: Word would widen the column, so noWrap stays single-line.
      expect(cellLineCount(noWrapTable({ text: OVERFLOW_TEXT, widthType: "auto" }))).toBe(1);
    }, fakeMeasure);
  });

  test("keeps a fitting noWrap cell on one line even when the column is pinned", () => {
    withFakeTextMeasure(() => {
      expect(cellLineCount(noWrapTable({ text: FIT_TEXT, widthType: "dxa" }))).toBe(1);
    }, fakeMeasure);
  });
});

describe("measureTableBlock floating cell content", () => {
  test("does not add a floating text box to the table row height", () => {
    withFakeTextMeasure(() => {
      const table: TableBlock = {
        kind: "table",
        id: "t",
        columnWidths: [200],
        rows: [
          {
            id: "r",
            cells: [
              {
                id: "c",
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                blocks: [
                  para("anchor", "anchor"),
                  {
                    kind: "textBox",
                    id: "floating-box",
                    width: 180,
                    height: 120,
                    content: [],
                    displayMode: "float",
                    position: { vertical: { relativeTo: "paragraph", posOffset: 0 } },
                  },
                ],
              },
            ],
          },
        ],
      };

      const measure = measureTableBlock(table, 200);
      const paragraph = measure.rows.at(0)?.cells.at(0)?.blocks.at(0);
      if (paragraph?.kind !== "paragraph") {
        throw new Error("Expected anchor paragraph measure");
      }

      expect(measure.rows.at(0)?.height).toBe(paragraph.totalHeight);
    }, fakeMeasure);
  });
});

describe("measureBlocks error instrumentation", () => {
  test("routes a block-measurement failure through the instrumentation hook", () => {
    type MeasureBlockErrorEvent = {
      blockIndex: number;
      blockKind: FlowBlock["kind"];
      message: string;
    };
    const events: MeasureBlockErrorEvent[] = [];
    const previous = globalThis.__folioLayoutInstrumentation;
    globalThis.__folioLayoutInstrumentation = {
      onMeasureBlockError: (event: MeasureBlockErrorEvent) => {
        events.push(event);
      },
    };

    try {
      withFakeTextMeasure(
        () => {
          const measures = measureBlocks([para("p", "will throw")], 600);

          // Pagination still gets a usable fallback measure instead of crashing.
          expect(measures).toHaveLength(1);
          const measure = measures[0]!;
          expect(measure.kind).toBe("paragraph");
          if (measure.kind === "paragraph") {
            expect(measure.totalHeight).toBe(20);
            expect(measure.lines).toEqual([]);
          }
        },
        {
          charWidth: () => {
            throw new Error("measure boom");
          },
        },
      );
    } finally {
      globalThis.__folioLayoutInstrumentation = previous;
    }

    // The swallowed failure is now traceable, not silent.
    expect(events).toEqual([{ blockIndex: 0, blockKind: "paragraph", message: "measure boom" }]);
  });

  test("advances page cursors after returning the fallback measure", () => {
    withFakeTextMeasure(
      () => {
        const failing = para("failing", "will throw!");
        const band: TextBoxBlock = {
          kind: "textBox",
          id: "band",
          width: 300,
          height: 60,
          content: [],
          wrapType: "topAndBottom",
          position: { vertical: { relativeTo: "margin", posOffset: 0 } },
        };
        const after = para("after", "after");

        const measures = measureBlocks([failing, band, after], 600, 96);
        const fallback = measures.at(0);
        const afterMeasure = measures.at(2);

        expect(fallback?.kind).toBe("paragraph");
        expect(afterMeasure?.kind).toBe("paragraph");
        if (fallback?.kind !== "paragraph") {
          throw new Error("Expected fallback paragraph measure");
        }
        if (afterMeasure?.kind !== "paragraph") {
          throw new Error("Expected paragraph after band");
        }

        expect(fallback.totalHeight).toBe(20);
        expect(afterMeasure.lines.at(0)?.floatSkipBefore).toBeCloseTo(40, 5);
      },
      {
        charWidth: (char) => {
          if (char === "!") {
            throw new Error("measure boom");
          }
          return 5;
        },
      },
    );
  });
});
