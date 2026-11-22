import { expect, test } from "bun:test";

import { layoutDocument } from "./index";
import { withFakeTextMeasure } from "./measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "./measure/measureBlocks";
import type {
  LayoutOptions,
  ParagraphBlock,
  ParagraphMeasure,
  TableBlock,
  TableMeasure,
} from "./types";

const options: LayoutOptions = {
  pageSize: { w: 300, h: 300 },
  margins: { top: 40, right: 40, bottom: 40, left: 40 },
};

const paragraphMeasure: ParagraphMeasure = {
  kind: "paragraph",
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 10,
      ascent: 16,
      descent: 4,
      lineHeight: 20,
    },
  ],
  totalHeight: 20,
};

type AnchorSpacingOptions = {
  contextualSpacing: boolean;
  tableOffset?: number;
  verticalAnchor?: "text" | "margin";
  measurement?: "supplied" | "real";
};

const anchorOffsets = ({
  contextualSpacing,
  tableOffset = 32,
  verticalAnchor = "text",
  measurement = "supplied",
}: AnchorSpacingOptions) => {
  const exactLine = { line: 20, lineRule: "exact", lineUnit: "px" } as const;
  const lead: ParagraphBlock = {
    kind: "paragraph",
    id: "lead",
    runs: [{ kind: "text", text: "L" }],
    attrs: {
      styleId: "Body",
      contextualSpacing,
      spacing: { ...exactLine, after: 4 },
    },
  };
  const anchor: ParagraphBlock = {
    kind: "paragraph",
    id: "anchor",
    runs: [{ kind: "text", text: "A" }],
    attrs: {
      styleId: "Body",
      contextualSpacing,
      spacing: { ...exactLine, before: 8 },
    },
  };
  const table: TableBlock = {
    kind: "table",
    id: "table",
    columnWidths: [60],
    rows: [
      {
        id: "row",
        cells: [
          {
            id: "cell",
            blocks: [
              {
                kind: "paragraph",
                id: "cell-paragraph",
                runs: [{ kind: "text", text: "C" }],
              },
            ],
          },
        ],
      },
    ],
    floating: {
      horzAnchor: "text",
      vertAnchor: verticalAnchor,
      tblpY: tableOffset,
    },
  };
  const tableMeasure: TableMeasure = {
    kind: "table",
    totalWidth: 60,
    totalHeight: 20,
    columnWidths: [60],
    rows: [
      {
        height: 20,
        cells: [{ blocks: [paragraphMeasure], width: 60, height: 20 }],
      },
    ],
  };
  const blocks = [lead, table, anchor];
  const measures =
    measurement === "real"
      ? measureBlocks(blocks, 220, options.margins.top, {
          pageWidth: options.pageSize.w,
          pageHeight: options.pageSize.h,
          marginLeft: options.margins.left,
          marginRight: options.margins.right,
          marginBottom: options.margins.bottom,
        })
      : [paragraphMeasure, tableMeasure, paragraphMeasure];
  const layout = layoutDocument(blocks, measures, options);
  const fragments = layout.pages.at(0)?.fragments;
  const leadFragment = fragments?.find(({ blockId }) => blockId === lead.id);
  const tableFragment = fragments?.find(({ blockId }) => blockId === table.id);
  const anchorFragment = fragments?.find(
    ({ blockId }) => blockId === anchor.id,
  );
  if (!leadFragment || !tableFragment || !anchorFragment) {
    throw new Error("Expected the lead, table, and anchor on one page");
  }
  const leadBottom = leadFragment.y + leadFragment.height;
  const anchorMeasure = measures.at(2);
  return {
    offsets: {
      table: tableFragment.y - leadBottom,
      anchor: anchorFragment.y - leadBottom,
    },
    wrapOffset:
      anchorMeasure?.kind === "paragraph"
        ? anchorMeasure.lines.at(0)?.leftOffset
        : undefined,
  };
};

test("a text-anchored table starts after the lead's trailing space", () => {
  expect(anchorOffsets({ contextualSpacing: false }).offsets).toEqual({
    table: 36,
    anchor: 12,
  });
});

test("same-style paragraphs suppress contextual spacing across a text-anchored table", () => {
  expect(anchorOffsets({ contextualSpacing: true }).offsets).toEqual({
    table: 32,
    anchor: 0,
  });
});

test("a margin-anchored table leaves ordinary spacing at its own position", () => {
  expect(
    anchorOffsets({ contextualSpacing: true, verticalAnchor: "margin" })
      .offsets,
  ).toEqual({ table: 12, anchor: 8 });
});

test("measures text beside the table when the anchor's spacing brings its line into the wrap zone", () => {
  withFakeTextMeasure(() => {
    const result = anchorOffsets({
      contextualSpacing: false,
      tableOffset: 20,
      measurement: "real",
    });
    expect(result.offsets).toEqual({ table: 24, anchor: 12 });
    expect(result.wrapOffset).toBe(72);
  });
});
