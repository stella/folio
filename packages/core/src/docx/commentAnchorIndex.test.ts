import { expect, test } from "bun:test";

import type { BlockContent } from "../types/document";
import {
  removeCommentMarkers,
  visitCommentMarkers,
} from "./commentAnchorIndex";

test("removeCommentMarkers rebuilds nested marker paths without changing its input", () => {
  const blocks: BlockContent[] = [
    {
      type: "blockSdt",
      properties: {},
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [
                    {
                      type: "paragraph",
                      content: [
                        {
                          type: "insertion",
                          info: { id: 7, author: "Reviewer" },
                          content: [
                            { type: "commentRangeStart", id: 12 },
                            {
                              type: "run",
                              content: [{ type: "text", text: "Clause" }],
                            },
                            { type: "commentRangeEnd", id: 12 },
                            { type: "commentReference", id: 12 },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
    {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "Stable" }] }],
    },
  ];
  const before = structuredClone(blocks);

  const rebuilt = removeCommentMarkers(blocks, new Set([12]));

  expect(blocks).toEqual(before);
  expect(rebuilt).not.toBe(blocks);
  expect(rebuilt.at(1)).toBe(blocks.at(1));
  const markers: number[] = [];
  visitCommentMarkers(rebuilt, ({ item }) => markers.push(item.id));
  expect(markers).toEqual([]);
});

test("removeCommentMarkers removes direct inline SDT range anchors", () => {
  const blocks: BlockContent[] = [
    {
      type: "paragraph",
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "group" },
          content: [
            { type: "commentRangeStart", id: 12 },
            { type: "run", content: [{ type: "text", text: "Clause" }] },
            { type: "commentRangeEnd", id: 12 },
          ],
        },
      ],
    },
  ];

  const rebuilt = removeCommentMarkers(blocks, new Set([12]));

  expect(rebuilt).toEqual([
    {
      type: "paragraph",
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "group" },
          content: [
            { type: "run", content: [{ type: "text", text: "Clause" }] },
          ],
        },
      ],
    },
  ]);
  const remaining: number[] = [];
  visitCommentMarkers(rebuilt, ({ item }) => remaining.push(item.id));
  expect(remaining).toEqual([]);
});
