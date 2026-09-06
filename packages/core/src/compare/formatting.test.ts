import { describe, expect, test } from "bun:test";

import type { FolioAIBlock } from "../ai-edits/types";
import { inlineFormattingSegments } from "./formatting";

const block = (previewRuns: FolioAIBlock["previewRuns"]): FolioAIBlock => ({
  id: "block",
  kind: "paragraph",
  text: "Contract",
  previewRuns,
});

describe("inlineFormattingSegments", () => {
  test("treats equivalent formatting across different run splits as equal", () => {
    expect(
      inlineFormattingSegments({
        baseBlock: block([
          { text: "Con", strike: true },
          { text: "tract", strike: true },
        ]),
        targetBlock: block([{ text: "Contract", strike: true }]),
        maxSegments: 10,
      }),
    ).toEqual([]);
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
