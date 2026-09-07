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

  test("reports target font face, half-point size, and normalized RGB color", () => {
    expect(
      inlineFormattingSegments({
        baseBlock: block([{ text: "Contract", fontFamily: "Arial", fontSizePt: 10 }]),
        targetBlock: block([
          {
            text: "Contract",
            fontFamily: "Georgia",
            fontSizePt: 10.5,
            color: "c00000",
            directFormatting: {
              fontFamily: "Georgia",
              fontSizePt: 10.5,
              color: "c00000",
            },
          },
        ]),
        maxSegments: 10,
      }),
    ).toEqual([
      {
        startOffset: 0,
        endOffset: 8,
        formatting: { fontFamily: "Georgia", fontSizePt: 10.5, color: "C00000" },
      },
    ]);
  });

  test("represents cleared direct font properties explicitly", () => {
    expect(
      inlineFormattingSegments({
        baseBlock: block([
          {
            text: "Contract",
            fontFamily: "Georgia",
            fontSizePt: 10.5,
            color: "C00000",
            directFormatting: {
              fontFamily: "Georgia",
              fontSizePt: 10.5,
              color: "C00000",
            },
          },
        ]),
        targetBlock: block([{ text: "Contract" }]),
        maxSegments: 10,
      }),
    ).toEqual([
      {
        startOffset: 0,
        endOffset: 8,
        formatting: { fontFamily: null, fontSizePt: null, color: null },
      },
    ]);
  });

  test("verification distinguishes an inherited value from the same direct value", () => {
    const inherited = block([{ text: "Contract", fontFamily: "Georgia" }]);
    const direct = block([
      {
        text: "Contract",
        fontFamily: "Georgia",
        directFormatting: { fontFamily: "Georgia" },
      },
    ]);

    expect(projectSupportedInlineFormatting(inherited)).not.toBe(
      projectSupportedInlineFormatting(direct),
    );
  });
});
