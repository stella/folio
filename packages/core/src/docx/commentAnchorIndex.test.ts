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

test("removeCommentMarkers visits markers inside hyperlinks and inline controls", () => {
  const blocks: BlockContent[] = [
    {
      type: "paragraph",
      content: [
        {
          type: "hyperlink",
          href: "https://example.org",
          children: [
            { type: "commentRangeStart", id: 21 },
            { type: "run", content: [{ type: "text", text: "linked" }] },
            { type: "commentRangeEnd", id: 21 },
            { type: "commentReference", id: 21 },
          ],
        },
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 2, tag: "clause" },
          content: [
            { type: "commentRangeStart", id: 22 },
            { type: "run", content: [{ type: "text", text: "controlled" }] },
            { type: "commentRangeEnd", id: 22 },
            { type: "commentReference", id: 22 },
          ],
        },
      ],
    },
  ];
  const markerIds: number[] = [];

  visitCommentMarkers(blocks, ({ item }) => markerIds.push(item.id));

  expect(markerIds).toEqual([21, 21, 21, 22, 22, 22]);
  const withoutMarkers = removeCommentMarkers(blocks, new Set([21, 22]));
  const remaining: number[] = [];
  visitCommentMarkers(withoutMarkers, ({ item }) => remaining.push(item.id));
  expect(remaining).toEqual([]);
});
