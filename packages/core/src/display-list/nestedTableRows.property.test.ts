import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { Window } from "happy-dom";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { layoutDocument } from "../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../layout-engine/measure/measureBlocks";
import {
  createTableCellFlowState,
  placeTableCellBlock,
} from "../layout-engine/measure/tableCellFlow";
import {
  type Layout,
  type PageMargins,
  type ParagraphBlock,
  resolveTableCellPadding,
  type TableBlock,
  type TableCell,
  type TableFragment,
  type TableMeasure,
  type TableRow,
} from "../layout-engine/types";
import type { BlockLookup } from "../layout-painter/index";
import { renderTableFragment } from "../layout-painter/renderTable";
import type { RenderContext } from "../layout-painter/renderUtils";
import { buildDisplayList } from "./build/buildDisplayList";
import type { DisplayPrimitive } from "./types";

setDefaultTimeout(propertyTestTimeout(20_000));

/**
 * Every cell of a row starts at the row's top with one line height, and a nested
 * table only sits beside a sibling when padding cannot shift its lines, so the
 * lines of a row meet at common bottoms and a clean break always exists.
 */
type CellShape = {
  readonly lineCounts: readonly number[];
  readonly nested: TableShape | null;
};
type RowShape = {
  readonly cells: readonly CellShape[];
  readonly keep: "none" | "cantSplit" | "exact";
};
type TableShape = {
  readonly rows: readonly RowShape[];
  readonly padding: number;
};

const CELL_WIDTH = 240;
const CHAR_WIDTH = 5;
const WORD = "word ";
/** How many words of `WORD` fill one line of the narrowest cell, roughly. */
const WORDS_PER_LINE = 6;

const cellShape = (
  nested: fc.Arbitrary<TableShape | null>,
): fc.Arbitrary<CellShape> =>
  fc.record({
    lineCounts: fc.array(fc.integer({ min: 1, max: 5 }), {
      minLength: 1,
      maxLength: 2,
    }),
    nested,
  });

const tableShape = (
  nested: fc.Arbitrary<TableShape | null>,
  maxRows: number,
  padding: fc.Arbitrary<number>,
  keep: fc.Arbitrary<RowShape["keep"]>,
): fc.Arbitrary<TableShape> =>
  fc.record({
    rows: fc.array(
      fc.record({
        cells: fc.array(cellShape(nested), { minLength: 1, maxLength: 2 }),
        keep,
      }),
      { minLength: 1, maxLength: maxRows },
    ),
    padding,
  });

const treeOf = (
  padding: fc.Arbitrary<number>,
  soleCellWhenNested: boolean,
): fc.Arbitrary<TableShape> => {
  const shaped = (shape: TableShape): TableShape =>
    soleCellWhenNested
      ? {
          ...shape,
          rows: shape.rows.map((row) =>
            row.cells.some((cell) => cell.nested)
              ? { ...row, cells: row.cells.slice(0, 1) }
              : row,
          ),
        }
      : shape;
  // A fixed height is off the line grid, so it can only sit in a lone stack.
  const keep = fc.constantFrom<RowShape["keep"]>(
    "none",
    "none",
    "none",
    "cantSplit",
    ...(soleCellWhenNested ? (["exact"] as const) : []),
  );
  const innermost = tableShape(fc.constant(null), 24, padding, keep);
  const middle = tableShape(
    fc.oneof(
      { weight: 3, arbitrary: fc.constant(null) },
      { weight: 1, arbitrary: innermost.map(shaped) },
    ),
    12,
    padding,
    keep,
  );
  return tableShape(middle.map(shaped), 2, padding, keep).map(shaped);
};

const buildParagraph = (
  id: string,
  token: string,
  lines: number,
): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [
    {
      kind: "text",
      text: `${token} ${WORD.repeat(lines * WORDS_PER_LINE).trimEnd()}`,
    },
  ],
});

/** A token at the start of the first paragraph of every cell of every row. */
const buildTable = (shape: TableShape, path: string): TableBlock => ({
  kind: "table",
  id: path,
  columnWidths: Array.from(
    { length: Math.max(...shape.rows.map((row) => row.cells.length)) },
    () => CELL_WIDTH,
  ),
  rows: shape.rows.map((row, rowIndex): TableRow => {
    const rowPath = `${path}r${String(rowIndex)}`;
    return {
      id: rowPath,
      ...(row.keep === "cantSplit" ? { cantSplit: true } : {}),
      ...(row.keep === "exact"
        ? { height: 60, heightRule: "exact" as const }
        : {}),
      cells: row.cells.map((cell, cellIndex): TableCell => {
        const cellPath = `${rowPath}c${String(cellIndex)}`;
        const paragraphs = cell.lineCounts.map((lines, index) =>
          buildParagraph(
            `${cellPath}p${String(index)}`,
            `T${cellPath}p${String(index)}T`,
            lines,
          ),
        );
        return {
          id: cellPath,
          padding: {
            top: shape.padding,
            right: 0,
            bottom: shape.padding,
            left: 0,
          },
          blocks: cell.nested
            ? [...paragraphs, buildTable(cell.nested, `${cellPath}n`)]
            : paragraphs,
        };
      }),
    };
  }),
});

type Band = { readonly top: number; readonly bottom: number };

/** A painted paragraph and the row that holds it, both as outer-row offsets. */
type PlacedParagraph = {
  readonly token: string;
  readonly paragraphTop: number;
  readonly paragraphBottom: number;
  /** Its lines as the page can show them: a fixed row clips what overflows it. */
  readonly lines: readonly Band[];
  readonly rowTop: number;
  readonly rowBottom: number;
};

const walkCell = (
  cell: TableCell,
  measure: TableMeasure["rows"][number]["cells"][number],
  cellTop: number,
  rowTop: number,
  rowBottom: number,
  fixedBottom: number,
  placed: PlacedParagraph[],
): void => {
  const padTop = resolveTableCellPadding(cell).top;
  const flow = createTableCellFlowState();
  cell.blocks.forEach((block, index) => {
    const blockMeasure = measure.blocks[index];
    if (!blockMeasure) {
      return;
    }
    const placement = placeTableCellBlock(flow, block, blockMeasure);
    const top = cellTop + padTop + placement.contentTop;
    if (block.kind === "paragraph" && blockMeasure.kind === "paragraph") {
      let lineTop = top;
      const lines = blockMeasure.lines.flatMap((line) => {
        const box = {
          top: lineTop,
          bottom: Math.min(lineTop + line.lineHeight, fixedBottom),
        };
        lineTop += line.lineHeight;
        return box.bottom > box.top ? [box] : [];
      });
      const run = block.runs[0];
      placed.push({
        token: run?.kind === "text" ? (run.text.split(" ")[0] ?? "") : "",
        paragraphTop: cellTop + padTop + placement.top,
        paragraphBottom: top + placement.contentHeight,
        lines,
        rowTop,
        rowBottom,
      });
    }
    if (block.kind === "table" && blockMeasure.kind === "table") {
      walkTable(block, blockMeasure, top, fixedBottom, placed);
    }
  });
};

const walkTable = (
  block: TableBlock,
  measure: TableMeasure,
  tableTop: number,
  fixedBottom: number,
  placed: PlacedParagraph[],
): void => {
  let rowTop = tableTop;
  measure.rows.forEach((rowMeasure, rowIndex) => {
    const row = block.rows[rowIndex];
    const rowBottom = rowTop + rowMeasure.height;
    row?.cells.forEach((cell, cellIndex) => {
      const cellMeasure = rowMeasure.cells[cellIndex];
      if (cellMeasure) {
        walkCell(
          cell,
          cellMeasure,
          rowTop,
          rowTop,
          rowBottom,
          row.heightRule === "exact"
            ? Math.min(fixedBottom, rowBottom)
            : fixedBottom,
          placed,
        );
      }
    });
    rowTop = rowBottom;
  });
};

/** What a fragment shows of one outer row, as offsets from that row's top. */
const fragmentBand = (
  fragment: TableFragment,
  rowIndex: number,
  rowHeight: number,
): Band | undefined => {
  if (rowIndex < fragment.fromRow || rowIndex >= fragment.toRow) {
    return undefined;
  }
  return {
    top: rowIndex === fragment.fromRow ? (fragment.topClip ?? 0) : 0,
    bottom:
      rowIndex === fragment.toRow - 1
        ? (fragment.bottomClip ?? rowHeight)
        : rowHeight,
  };
};

type PaintedFragment = {
  readonly fragment: TableFragment;
  readonly pageIndex: number;
  /** Text this painter built for the fragment. */
  readonly text: string;
};

const collectGlyphText = (primitives: readonly DisplayPrimitive[]): string =>
  primitives
    .map((primitive) => {
      switch (primitive.kind) {
        case "glyphRun":
          return primitive.text;
        case "clipGroup":
        case "rotateGroup":
        case "opacityGroup":
          return collectGlyphText(primitive.children);
        default:
          return "";
      }
    })
    .join(" ");

const tableFragments = (layout: Layout): [TableFragment, number][] =>
  layout.pages.flatMap((page, pageIndex) =>
    page.fragments.flatMap((fragment): [TableFragment, number][] =>
      fragment.kind === "table" ? [[fragment, pageIndex]] : [],
    ),
  );

const context: RenderContext = {
  pageNumber: 1,
  totalPages: 1,
  section: "body",
};
const doc = new Window().document as unknown as Document;

const paintLegacy = (
  layout: Layout,
  block: TableBlock,
  measure: TableMeasure,
): PaintedFragment[] =>
  tableFragments(layout).map(([fragment, pageIndex]) => ({
    fragment,
    pageIndex,
    text:
      renderTableFragment(fragment, block, measure, context, { document: doc })
        .textContent ?? "",
  }));

/**
 * One fragment per page, which the outer table in these shapes guarantees for
 * every page holding a clipped row; several rows on one page share its text.
 */
const paintDisplayList = (
  layout: Layout,
  block: TableBlock,
  measure: TableMeasure,
): PaintedFragment[] => {
  const blockLookup: BlockLookup = new Map([
    [String(block.id), { block, measure }],
  ]);
  const list = buildDisplayList({ layout, blockLookup });
  return tableFragments(layout).map(([fragment, pageIndex]) => ({
    fragment,
    pageIndex,
    text: collectGlyphText(list.pages[pageIndex]?.primitives ?? []),
  }));
};

const margins: PageMargins = { top: 20, right: 20, bottom: 20, left: 20 };

const bandHolds = (band: Band, line: Band): boolean =>
  band.top <= line.top + 0.01 && band.bottom >= line.bottom - 0.01;

const assertShape = (outer: TableShape, pageHeight: number): void => {
  withFakeTextMeasure(
    () => {
      const block = buildTable(outer, "t");
      const measures = measureBlocks(
        [block],
        CELL_WIDTH * block.columnWidths.length,
      );
      const measure = measures[0];
      if (measure?.kind !== "table") {
        throw new TypeError("A table block measures as a table.");
      }
      const layout = layoutDocument([block], measures, {
        pageSize: { w: CELL_WIDTH * 3, h: pageHeight },
        margins,
      });
      const pageBottom = pageHeight - margins.bottom;

      measure.rows.forEach((rowMeasure, rowIndex) => {
        const placed: PlacedParagraph[] = [];
        const row = block.rows[rowIndex];
        row?.cells.forEach((cell, cellIndex) => {
          const cellMeasure = rowMeasure.cells[cellIndex];
          if (cellMeasure) {
            walkCell(
              cell,
              cellMeasure,
              0,
              0,
              rowMeasure.height,
              row.heightRule === "exact"
                ? rowMeasure.height
                : Number.POSITIVE_INFINITY,
              placed,
            );
          }
        });
        const bands = tableFragments(layout).flatMap(([fragment]) => {
          const band = fragmentBand(fragment, rowIndex, rowMeasure.height);
          return band && fragment.y + fragment.height <= pageBottom + 0.01
            ? [band]
            : [];
        });
        for (const paragraph of placed) {
          for (const line of paragraph.lines) {
            if (!bands.some((band) => bandHolds(band, line))) {
              throw new Error(
                `${paragraph.token} line at ${String(line.top)} is on no page`,
              );
            }
          }
        }

        for (const paint of [paintLegacy, paintDisplayList]) {
          const painted = paint(layout, block, measure);
          for (const paragraph of placed) {
            const firstLine = paragraph.lines[0];
            if (!firstLine) {
              continue;
            }
            const shown = painted.some(({ fragment, text }) => {
              const band = fragmentBand(fragment, rowIndex, rowMeasure.height);
              return (
                band !== undefined &&
                bandHolds(band, firstLine) &&
                text.includes(paragraph.token)
              );
            });
            if (!shown) {
              throw new Error(
                `${paint.name}: ${paragraph.token} is painted on no page`,
              );
            }
          }
        }

        // A clipped row builds nothing its slice cannot show. The display list
        // reports a page's text, so it can only answer for a page it has alone.
        const fragmentsOnPage = (pageIndex: number): number =>
          tableFragments(layout).filter(([, index]) => index === pageIndex)
            .length;
        for (const paint of [paintLegacy, paintDisplayList]) {
          for (const { fragment, pageIndex, text } of paint(
            layout,
            block,
            measure,
          )) {
            const band = fragmentBand(fragment, rowIndex, rowMeasure.height);
            const clipped =
              fragment.topClip !== undefined ||
              fragment.bottomClip !== undefined;
            if (!band || !clipped || fragmentsOnPage(pageIndex) !== 1) {
              continue;
            }
            for (const paragraph of placed) {
              const outsideBand =
                paragraph.rowBottom <= band.top + 0.01 ||
                paragraph.rowTop >= band.bottom - 0.01;
              if (outsideBand && text.includes(paragraph.token)) {
                throw new Error(
                  `${paint.name}: ${paragraph.token} built on a page that cannot show it`,
                );
              }
            }
          }
        }
      });
    },
    { charWidth: fixedCharWidth(CHAR_WIDTH) },
  );
};

test.each([
  [
    "with padding, a nested table alone in its row",
    treeOf(fc.constantFrom(0, 2, 5), true),
  ],
  ["without padding, beside sibling cells", treeOf(fc.constant(0), false)],
] as const)(
  "every line of every nested row is whole on some page of both painters: %s",
  (_label, shape) => {
    fc.assert(
      fc.property(
        fc.record({
          pageHeight: fc.integer({ min: 240, max: 520 }),
          outer: shape,
        }),
        ({ pageHeight, outer }) => {
          assertShape(outer, pageHeight);
        },
      ),
      propertyConfig({ numRuns: 30 }),
    );
  },
);

test("the first whole nested row of a continuation draws its top rule on both painters", () => {
  withFakeTextMeasure(
    () => {
      const rule = { width: 2, style: "solid" as const, color: "#123456" };
      const inner: TableBlock = {
        kind: "table",
        id: "inner",
        columnWidths: [CELL_WIDTH],
        rows: Array.from({ length: 20 }, (_, index) => ({
          id: `r${String(index)}`,
          cells: [
            {
              id: `c${String(index)}`,
              padding: { top: 0, right: 0, bottom: 0, left: 0 },
              borders: { top: rule, bottom: rule },
              blocks: [
                {
                  kind: "paragraph" as const,
                  id: `p${String(index)}`,
                  runs: [
                    { kind: "text" as const, text: `Row${String(index)}` },
                  ],
                },
              ],
            },
          ],
        })),
      };
      const outer: TableBlock = {
        kind: "table",
        id: "outer",
        columnWidths: [CELL_WIDTH],
        rows: [
          {
            id: "row",
            cells: [
              {
                id: "cell",
                padding: { top: 0, right: 0, bottom: 0, left: 0 },
                blocks: [inner],
              },
            ],
          },
        ],
      };
      const measures = measureBlocks([outer], CELL_WIDTH);
      const measure = measures[0];
      if (measure?.kind !== "table") {
        throw new TypeError("A table block measures as a table.");
      }
      const layout = layoutDocument([outer], measures, {
        pageSize: { w: CELL_WIDTH * 2, h: 240 },
        margins,
      });
      const continuation = tableFragments(layout).find(
        ([fragment]) => fragment.topClip !== undefined,
      );
      if (!continuation) {
        throw new TypeError("A table this tall continues on another page.");
      }
      const [fragment, pageIndex] = continuation;

      const element = renderTableFragment(fragment, outer, measure, context, {
        document: doc,
      });
      const firstCell = element.querySelector<HTMLElement>(
        ".layout-nested-table .layout-table-cell",
      );
      expect(firstCell?.style.borderTop).not.toBe("none");
      expect(firstCell?.style.borderTop).not.toBe("");

      const list = buildDisplayList({
        layout,
        blockLookup: new Map([[String(outer.id), { block: outer, measure }]]),
      });
      const clip = list.pages[pageIndex]?.primitives.find(
        (primitive) => primitive.kind === "clipGroup",
      );
      if (clip?.kind !== "clipGroup") {
        throw new TypeError("A continuation paints inside a clip group.");
      }
      const clipTop = clip.rect.yPx;
      expect(
        clip.children.some(
          (primitive) =>
            primitive.kind === "line" &&
            primitive.y1Px === primitive.y2Px &&
            primitive.y1Px >= clipTop &&
            primitive.y1Px <= clipTop + rule.width,
        ),
      ).toBe(true);
    },
    { charWidth: fixedCharWidth(CHAR_WIDTH) },
  );
});
