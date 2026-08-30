import { describe, expect, test } from "bun:test";
import { panic } from "better-result";

import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import type { FlowBlock, ParagraphBlock, Run } from "../../layout-engine/types";
import type { Footnote } from "../../types/document";
import { applyFootnotePresentation, convertFootnoteToContent } from "./footnoteLayout";

const fakeMeasure = { charWidth: fixedCharWidth(5) };

function presentedFootnoteParagraph(runs: Run[]): ParagraphBlock {
  const paragraph = applyFootnotePresentation(
    [{ kind: "paragraph", id: "footnote-text", runs }],
    8,
  ).at(0);
  if (paragraph?.kind !== "paragraph") {
    panic("Expected a paragraph block");
  }
  return paragraph;
}

function visibleRunSequence(runs: Run[]): string {
  return runs
    .map((run) => {
      switch (run.kind) {
        case "text":
          return run.text;
        case "tab":
          return "\t";
        case "image":
          return "\ufffc";
        case "lineBreak":
          return "\n";
        case "renderedPageBreak":
          return "";
        case "field":
          return run.fallback || "1";
        case "math":
          return run.plainText;
        default: {
          const exhaustiveRun: never = run;
          return exhaustiveRun;
        }
      }
    })
    .join("");
}

const footnoteWithTable: Footnote = {
  type: "footnote",
  id: 7,
  noteType: "normal",
  content: [
    {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "Intro" }] }],
    },
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
                      type: "run",
                      content: [{ type: "text", text: "Cell" }],
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
};

const emptyFootnoteWithTable: Footnote = {
  type: "footnote",
  id: 8,
  noteType: "normal",
  content: [
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [{ type: "paragraph", content: [] }],
            },
          ],
        },
      ],
    },
  ],
};

const footnoteWithBlockSdt: Footnote = {
  type: "footnote",
  id: 10,
  noteType: "normal",
  content: [
    {
      type: "blockSdt",
      properties: {
        tag: "cite",
        alias: "Citation",
      },
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Smith v Jones" }],
            },
          ],
        },
      ],
    },
  ],
};

const footnoteWithRowSpanTable: Footnote = {
  type: "footnote",
  id: 9,
  noteType: "normal",
  content: [
    {
      type: "table",
      columnWidths: [1440, 2880],
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              formatting: { vMerge: "restart" },
              content: [
                {
                  type: "paragraph",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "A" }],
                    },
                  ],
                },
              ],
            },
            {
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "B" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              formatting: { vMerge: "continue" },
              content: [{ type: "paragraph", content: [] }],
            },
            {
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "C" }],
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
};

describe("footnote layout", () => {
  test("applies document line-breaking policy to footnote paragraphs", () => {
    const content = convertFootnoteToContent(footnoteWithTable, 3, 400, {
      measureBlocks: (blocks) =>
        blocks.map(() => ({ kind: "paragraph" as const, lines: [], totalHeight: 12 })),
      automaticHyphenation: { enabled: true, doNotHyphenateCaps: true },
      lineBreakRules: {
        noLineBreaksBefore: { language: "ja-JP", characters: "※" },
      },
    });

    expect(content.blocks.at(0)).toMatchObject({
      kind: "paragraph",
      attrs: {
        automaticHyphenation: { enabled: true, doNotHyphenateCaps: true },
        lineBreakRules: {
          noLineBreaksBefore: { language: "ja-JP", characters: "※" },
        },
      },
    });
  });

  test("routes footnotes through the body pipeline so tables survive", () => {
    const content = convertFootnoteToContent(footnoteWithTable, 3, 400, {
      measureBlocks(blocks) {
        return blocks.map((block) =>
          block.kind === "table"
            ? {
                kind: "table",
                rows: [],
                columnWidths: [400],
                totalWidth: 400,
                totalHeight: 24,
              }
            : { kind: "paragraph", lines: [], totalHeight: 12 },
        );
      },
    });

    expect(content.blocks.map((block) => block.kind)).toEqual(["paragraph", "table"]);
    expect(content.height).toBe(36);
  });

  test("renders paragraphs nested inside footnote block SDTs", () => {
    const content = convertFootnoteToContent(footnoteWithBlockSdt, 10, 400, {
      measureBlocks(blocks) {
        return blocks.map(() => ({
          kind: "paragraph",
          lines: [],
          totalHeight: 12,
        }));
      },
    });

    expect(content.blocks).toHaveLength(1);
    const paragraph = content.blocks.at(0);
    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected footnote SDT to produce a paragraph");
    }

    expect(paragraph.runs.at(1)).toMatchObject({
      kind: "text",
      text: "Smith v Jones",
    });
    expect(paragraph.sdtGroups?.at(0)).toMatchObject({
      tag: "cite",
      alias: "Citation",
    });
  });

  test("measures table footnotes without a caller-provided measurement hook", () => {
    withFakeTextMeasure(() => {
      const content = convertFootnoteToContent(emptyFootnoteWithTable, 3, 400);

      expect(content.blocks.map((block) => block.kind)).toEqual(["paragraph", "table"]);
      expect(Number.isNaN(content.height)).toBe(false);
      expect(content.height).toBeGreaterThan(0);

      const numberBlock = content.blocks.at(0);
      expect(numberBlock?.kind).toBe("paragraph");
      if (numberBlock?.kind !== "paragraph") {
        throw new Error("Expected footnote number paragraph");
      }
      expect(numberBlock.runs.at(0)).toMatchObject({
        kind: "text",
        text: "3  ",
        superscript: true,
      });

      const tableMeasure = content.measures.at(1);
      expect(tableMeasure?.kind).toBe("table");
      if (tableMeasure?.kind !== "table") {
        throw new Error("Expected footnote table to have a table measure");
      }
      expect(tableMeasure.totalHeight).toBeGreaterThan(0);
    }, fakeMeasure);
  });

  test("skips row-spanned columns while measuring footnote table rows", () => {
    withFakeTextMeasure(() => {
      const content = convertFootnoteToContent(footnoteWithRowSpanTable, 4, 400);
      const tableMeasure = content.measures.at(1);

      expect(tableMeasure?.kind).toBe("table");
      if (tableMeasure?.kind !== "table") {
        throw new Error("Expected footnote table to have a table measure");
      }

      expect(tableMeasure.rows.at(0)?.cells.at(0)?.rowSpan).toBe(2);
      expect(tableMeasure.rows.at(0)?.cells.at(0)?.width).toBeCloseTo(96);
      expect(tableMeasure.rows.at(1)?.cells.at(0)?.width).toBeCloseTo(192);
    }, fakeMeasure);
  });

  test("applies footnote font size to nested table paragraphs and field runs", () => {
    const blocks: FlowBlock[] = [
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
                    id: "cell-p-1",
                    runs: [
                      { kind: "text", text: "Cell" },
                      { kind: "field", fieldType: "PAGE", fallback: "1" },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    ];

    const table = applyFootnotePresentation(blocks, 4).at(1);
    expect(table?.kind).toBe("table");
    if (table?.kind !== "table") {
      throw new Error("Expected a table block");
    }

    const paragraph = table.rows.at(0)?.cells.at(0)?.blocks.at(0);
    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected nested paragraph block");
    }

    expect(paragraph.runs.at(0)?.fontSize).toBe(8);
    expect(paragraph.runs.at(1)?.fontSize).toBe(8);
  });

  test("matches footnote number typography to the first footnote text run", () => {
    const blocks: FlowBlock[] = [
      {
        kind: "paragraph",
        id: "footnote-text",
        runs: [
          {
            kind: "text",
            text: " Insert the name of the legal entity.",
            fontFamily: "Times New Roman",
            fontSize: 10,
          },
        ],
        attrs: {
          defaultFontFamily: "Times New Roman",
          defaultFontSize: 10,
        },
      },
    ];

    const paragraph = applyFootnotePresentation(blocks, 8).at(0);
    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      panic("Expected a paragraph block");
    }

    expect(paragraph.runs.at(0)).toMatchObject({
      kind: "text",
      text: "8",
      fontFamily: "Times New Roman",
      fontSize: 10,
      superscript: true,
    });
    expect(visibleRunSequence(paragraph.runs)).toBe("8 Insert the name of the legal entity.");
  });

  test("adds one separator space when footnote text has no leading space", () => {
    const blocks: FlowBlock[] = [
      {
        kind: "paragraph",
        id: "footnote-text",
        runs: [
          {
            kind: "text",
            text: "Footnote text",
            fontFamily: "Times New Roman",
            fontSize: 10,
          },
        ],
      },
    ];

    const paragraph = applyFootnotePresentation(blocks, 8).at(0);
    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected a paragraph block");
    }

    expect(paragraph.runs.at(0)).toMatchObject({
      kind: "text",
      text: "8 ",
      fontFamily: "Times New Roman",
      fontSize: 10,
      superscript: true,
    });
    expect(visibleRunSequence(paragraph.runs)).toBe("8 Footnote text");
  });

  test.each([
    ") Footnote text",
    "）脚注テキスト",
    ". Footnote text",
    "。脚注テキスト",
    ": Label",
    "：ラベル",
  ])("does not insert a space before authored punctuation: %s", (text) => {
    const blocks: FlowBlock[] = [
      {
        kind: "paragraph",
        id: "footnote-text",
        runs: [
          {
            kind: "text",
            text,
            fontFamily: "Times New Roman",
            fontSize: 10,
          },
        ],
      },
    ];

    const paragraph = applyFootnotePresentation(blocks, 8).at(0);
    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected a paragraph block");
    }

    expect(paragraph.runs.at(0)).toMatchObject({
      kind: "text",
      text: "8",
      fontFamily: "Times New Roman",
      fontSize: 10,
      superscript: true,
    });
    expect(visibleRunSequence(paragraph.runs)).toBe(`8${text}`);
  });

  test.each([
    "(See clause 4)",
    "（条項4参照）",
    "“Quoted authority”",
    "— explanatory text",
    "_connector",
  ])("keeps a separator before opening or non-attaching punctuation: %s", (text) => {
    const paragraph = presentedFootnoteParagraph([{ kind: "text", text }]);

    expect(visibleRunSequence(paragraph.runs)).toBe(`8 ${text}`);
  });

  test.each([
    {
      name: "empty run before closing suffix",
      runs: [
        { kind: "text", text: "" },
        { kind: "text", text: ") Footnote text" },
      ],
      expected: "8) Footnote text",
    },
    {
      name: "split authored whitespace",
      runs: [
        { kind: "text", text: "" },
        { kind: "text", text: " " },
        { kind: "text", text: "Footnote text" },
      ],
      expected: "8 Footnote text",
    },
  ] satisfies { name: string; runs: Run[]; expected: string }[])(
    "resolves adjacency from the first visible text across $name",
    ({ runs, expected }) => {
      const paragraph = presentedFootnoteParagraph(runs);

      expect(visibleRunSequence(paragraph.runs)).toBe(expected);
    },
  );

  test.each([
    {
      name: "tab separator",
      runs: [{ kind: "tab" }, { kind: "text", text: "Footnote text" }],
      expected: "8\tFootnote text",
    },
    {
      name: "field closing suffix",
      runs: [
        { kind: "field", fieldType: "OTHER", fallback: ")" },
        { kind: "text", text: " Footnote text" },
      ],
      expected: "8) Footnote text",
    },
    {
      name: "ordinary field text",
      runs: [{ kind: "field", fieldType: "OTHER", fallback: "Authority" }],
      expected: "8 Authority",
    },
    {
      name: "inline image",
      runs: [
        { kind: "image", src: "data:image/png;base64,", width: 10, height: 10 },
        { kind: "text", text: "Caption" },
      ],
      expected: "8 \ufffcCaption",
    },
  ] satisfies { name: string; runs: Run[]; expected: string }[])(
    "uses the first paint-bearing non-text run for $name",
    ({ runs, expected }) => {
      const paragraph = presentedFootnoteParagraph(runs);

      expect(visibleRunSequence(paragraph.runs)).toBe(expected);
    },
  );

  test("keeps the footnote number alternate paired with its primary source", () => {
    const blocks: FlowBlock[] = [
      {
        kind: "paragraph",
        id: "footnote-text",
        runs: [{ kind: "text", text: "Footnote text", fontFamily: "Direct Face" }],
        attrs: {
          defaultFontFamily: "Default Face",
          defaultAlternateFontFamily: "Default Alternate",
        },
      },
    ];

    const paragraph = applyFootnotePresentation(blocks, 8).at(0);
    expect(paragraph?.kind).toBe("paragraph");
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected a paragraph block");
    }

    expect(paragraph.runs.at(0)).toMatchObject({ fontFamily: "Direct Face" });
    expect(paragraph.runs.at(0)?.alternateFontFamily).toBeUndefined();
  });
});
