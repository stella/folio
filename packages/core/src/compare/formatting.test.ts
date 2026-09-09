import { describe, expect, test } from "bun:test";

import type { FolioAIBlock, FolioAIInlineBooleanProperty } from "../ai-edits/types";
import { inlineFormattingSegments } from "./formatting";
import { projectSupportedInlineFormatting } from "./verification";

const block = (previewRuns: FolioAIBlock["previewRuns"]): FolioAIBlock => ({
  id: "block",
  kind: "paragraph",
  text: "Contract",
  previewRuns,
});

const DIRECT_BOOLEAN_STATES = [
  { label: "absent" },
  { label: "on", value: true },
  { label: "off", value: false },
] as const;

const INLINE_BOOLEAN_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "strike",
] as const satisfies readonly FolioAIInlineBooleanProperty[];

const blockWithBooleanState = (
  property: (typeof INLINE_BOOLEAN_PROPERTIES)[number],
  inherited: boolean,
  direct: (typeof DIRECT_BOOLEAN_STATES)[number],
): FolioAIBlock => {
  const effective = "value" in direct ? direct.value : inherited;
  return block([
    {
      text: "Contract",
      ...(effective && { [property]: true }),
      ...("value" in direct && { directFormatting: { [property]: direct.value } }),
    },
  ]);
};

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

  test("preserves direct boolean on, off, and absence across every inherited state", () => {
    for (const property of INLINE_BOOLEAN_PROPERTIES) {
      for (const inherited of [false, true]) {
        for (const baseDirect of DIRECT_BOOLEAN_STATES) {
          for (const targetDirect of DIRECT_BOOLEAN_STATES) {
            const base = blockWithBooleanState(property, inherited, baseDirect);
            const target = blockWithBooleanState(property, inherited, targetDirect);
            const sameDirectState = baseDirect.label === targetDirect.label;
            const expectedSegments = sameDirectState
              ? []
              : [
                  {
                    startOffset: 0,
                    endOffset: 8,
                    formatting: {
                      [property]: "value" in targetDirect ? targetDirect.value : null,
                    },
                  },
                ];

            expect({
              property,
              inherited,
              base: baseDirect.label,
              target: targetDirect.label,
              segments: inlineFormattingSegments({
                baseBlock: base,
                targetBlock: target,
                maxSegments: 10,
              }),
            }).toEqual({
              property,
              inherited,
              base: baseDirect.label,
              target: targetDirect.label,
              segments: expectedSegments,
            });
            expect(
              projectSupportedInlineFormatting(base) === projectSupportedInlineFormatting(target),
            ).toBe(sameDirectState);
          }
        }
      }
    }
  });
});
