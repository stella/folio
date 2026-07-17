import { describe, expect, test } from "bun:test";

import type { FlowBlock } from "../../packages/core/src/layout-engine/types";
import { traceFlowBlocks } from "../sourceTraceFlow";

const TABLE_BLOCKS = [
  {
    kind: "table",
    id: "table-1",
    rows: [
      {
        id: "row-1",
        cells: [
          {
            id: "cell-1",
            blocks: [
              {
                kind: "paragraph",
                id: "paragraph-1",
                runs: [{ kind: "text", text: "Alpha", fontFamily: "Aptos" }],
              },
            ],
          },
          {
            id: "cell-2",
            blocks: [
              {
                kind: "paragraph",
                id: "paragraph-2",
                attrs: { alignment: "center" },
                runs: [{ kind: "text", text: "Beta" }],
              },
            ],
          },
        ],
      },
    ],
  },
] as const satisfies FlowBlock[];

describe("flow source tracing", () => {
  test("matches text spanning multiple table cells and reports cell context", () => {
    const matches = traceFlowBlocks({ blocks: TABLE_BLOCKS, query: "alpha beta", limit: 8 });

    expect(matches).toHaveLength(1);
    expect(matches.at(0)).toMatchObject({
      type: "tableRow",
      path: ["blocks", 0, "rows", 0],
      tableId: "table-1",
      rowId: "row-1",
      rowIndex: 0,
      cells: [
        {
          cellIndex: 0,
          cellId: "cell-1",
          text: "Alpha",
          blocks: [{ type: "paragraph", id: "paragraph-1", text: "Alpha" }],
        },
        {
          cellIndex: 1,
          cellId: "cell-2",
          text: "Beta",
          blocks: [
            {
              type: "paragraph",
              id: "paragraph-2",
              text: "Beta",
              attrs: { alignment: "center" },
            },
          ],
        },
      ],
    });
  });

  test("reports a paragraph inside a table alongside its row context", () => {
    const matches = traceFlowBlocks({ blocks: TABLE_BLOCKS, query: "beta", limit: 8 });

    expect(matches.map(({ type }) => type)).toEqual(["tableRow", "paragraph"]);
    expect(matches.at(1)).toMatchObject({
      type: "paragraph",
      path: ["blocks", 0, "rows", 0, "cells", 1, "blocks", 0],
      id: "paragraph-2",
      attrs: { alignment: "center" },
    });
  });

  test("enforces the match limit across nested blocks", () => {
    expect(traceFlowBlocks({ blocks: TABLE_BLOCKS, query: "beta", limit: 1 })).toHaveLength(1);
  });

  test("searches complete cell text beyond the diagnostic preview", () => {
    const longPrefix = "x".repeat(300);
    const blocks = [
      {
        kind: "table",
        id: "long-table",
        rows: [
          {
            id: "long-row",
            cells: [
              {
                id: "long-cell",
                blocks: [
                  {
                    kind: "paragraph",
                    id: "long-paragraph",
                    runs: [{ kind: "text", text: `${longPrefix} needle` }],
                  },
                ],
              },
            ],
          },
        ],
      },
    ] satisfies FlowBlock[];

    expect(traceFlowBlocks({ blocks, query: "needle", limit: 8 }).at(0)).toMatchObject({
      type: "tableRow",
      tableId: "long-table",
    });
  });
});
