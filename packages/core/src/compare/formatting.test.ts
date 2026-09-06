import { describe, expect, test } from "bun:test";

import type { FolioAIBlock } from "../ai-edits/types";
import { inlineFormattingSegments } from "./formatting";
import { projectSupportedInlineFormatting } from "./verification";

const block = (previewRuns: FolioAIBlock["previewRuns"]): FolioAIBlock => ({
  id: "block",
  kind: "paragraph",
  text: "Contract",
  previewRuns,
});

describe("inlineFormattingSegments", () => {
  test("treats equivalent formatting across different run splits as equal", () => {
    const split = block([
      { text: "", bold: true },
      { text: "Con", strike: true },
      { text: "tract", strike: true },
      { text: "", italic: true },
    ]);
    const joined = block([{ text: "Contract", strike: true }]);

    expect(
      inlineFormattingSegments({ baseBlock: split, targetBlock: joined, maxSegments: 10 }),
    ).toEqual([]);
    expect(projectSupportedInlineFormatting(split)).toBe(projectSupportedInlineFormatting(joined));
  });

  test("coalesces adjacent strike changes across target run boundaries", () => {
    expect(
      inlineFormattingSegments({
        baseBlock: block([{ text: "Contract" }]),
        targetBlock: block([
          { text: "Con", strike: true },
          { text: "tract", strike: true },
        ]),
        maxSegments: 10,
      }),
    ).toEqual([{ startOffset: 0, endOffset: 8, formatting: { strike: true } }]);
  });
});
