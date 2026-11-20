import { expect, test } from "bun:test";

import type { BlockContent } from "../types/document";
import { collectBlockTexts } from "./blockPlainText";

test.each(["footnote", "endnote"] as const)(
  "a %s's blocks read accepted as that note, its automatic mark included",
  (kind) => {
    const blocks: BlockContent[] = [
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          { type: "run", content: [{ type: "noteMarker", kind }] },
          { type: "run", content: [{ type: "text", text: "Kept " }] },
          {
            type: "insertion",
            info: { id: 1, author: "A" },
            content: [
              { type: "run", content: [{ type: "text", text: "added" }] },
            ],
          },
          {
            type: "deletion",
            info: { id: 2, author: "A" },
            content: [
              { type: "run", content: [{ type: "text", text: "gone" }] },
            ],
          },
        ],
      },
    ];
    expect(collectBlockTexts(blocks)).toEqual(["Kept added"]);
  },
);
