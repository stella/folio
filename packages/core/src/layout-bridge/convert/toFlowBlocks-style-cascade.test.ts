import { describe, expect, test } from "bun:test";

import type {
  ParagraphBlock,
  TableBlock as LayoutTableBlock,
  TextRun,
} from "../../layout-engine/types";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, StyleDefinitions, Table } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

function makeDoc(paragraph: Paragraph, styles?: StyleDefinitions): Document {
  return {
    package: {
      document: { content: [paragraph] },
      ...(styles ? { styles } : {}),
    },
  };
}

function firstParagraph(blocks: unknown[]): ParagraphBlock {
  return blocks.find(
    (block) => (block as { kind?: string }).kind === "paragraph",
  ) as ParagraphBlock;
}

function firstRun(blocks: unknown[]): TextRun {
  return firstParagraph(blocks).runs[0] as TextRun;
}

function firstTableRun(blocks: unknown[]): TextRun {
  const table = blocks.find(
    (block) => (block as { kind?: string }).kind === "table",
  ) as LayoutTableBlock;
  const paragraph = table.rows[0]?.cells[0]?.blocks[0] as ParagraphBlock;
  return paragraph.runs[0] as TextRun;
}

describe("toFlowBlocks style cascade", () => {
  test("paragraph style rFonts reaches runs without explicit font marks", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "Normal",
          type: "paragraph",
          default: true,
          name: "Normal",
          rPr: { fontFamily: { ascii: "Arial Narrow", hAnsi: "Arial Narrow" } },
        },
        {
          styleId: "Clauses",
          type: "paragraph",
          basedOn: "Normal",
          name: "Clauses",
          rPr: { fontFamily: { ascii: "Arial Narrow", hAnsi: "Arial Narrow" } },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Clauses" },
      content: [
        {
          type: "run",
          content: [{ type: "text", text: "clause one" }],
        },
      ],
    };

    const blocks = toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {});

    expect(firstRun(blocks).fontFamily).toBe("Arial Narrow");
  });

  test("run with partial rFonts inherits ascii font from paragraph defaults", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "Normal",
          type: "paragraph",
          default: true,
          name: "Normal",
          rPr: { fontFamily: { ascii: "Arial Narrow", hAnsi: "Arial Narrow" } },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Normal" },
      content: [
        {
          type: "run",
          formatting: { fontFamily: { eastAsia: "Calibri" } },
          content: [{ type: "text", text: "mixed" }],
        },
      ],
    };

    const blocks = toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {});

    expect(firstRun(blocks).fontFamily).toBe("Arial Narrow");
  });

  test("inherits the East Asian font from the paragraph style default (eigenpal/docx-editor#949)", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "Normal",
          type: "paragraph",
          default: true,
          name: "Normal",
          rPr: { fontFamily: { ascii: "Arial", eastAsia: "MS Mincho" } },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Normal" },
      content: [
        {
          type: "run",
          content: [{ type: "text", text: "和文 mixed" }],
        },
      ],
    };

    const run = firstRun(toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {}));

    expect(run.fontFamily).toBe("Arial");
    expect(run.eastAsiaFontFamily).toBe("MS Mincho");
  });

  test("a direct eastAsia run mark overrides the inherited EA default", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "Normal",
          type: "paragraph",
          default: true,
          name: "Normal",
          rPr: { fontFamily: { ascii: "Arial", eastAsia: "MS Mincho" } },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Normal" },
      content: [
        {
          type: "run",
          formatting: { fontFamily: { eastAsia: "MS Gothic" } },
          content: [{ type: "text", text: "和文" }],
        },
      ],
    };

    const run = firstRun(toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {}));

    expect(run.eastAsiaFontFamily).toBe("MS Gothic");
  });

  test("explicit run formatting toggles override paragraph style defaults", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "Heading",
          type: "paragraph",
          name: "Heading",
          rPr: {
            bold: true,
            italic: true,
            allCaps: true,
          },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Heading" },
      content: [
        {
          type: "run",
          formatting: {
            bold: false,
            italic: false,
            allCaps: false,
          },
          content: [{ type: "text", text: "Keep mixed case" }],
        },
      ],
    };

    const blocks = toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {});
    const run = firstRun(blocks);

    expect(run.bold).toBe(false);
    expect(run.italic).toBe(false);
    expect(run.allCaps).toBe(false);
  });

  test("default character style reaches runs without rStyle", () => {
    const styles: StyleDefinitions = {
      docDefaults: { rPr: { fontSize: 22 } },
      styles: [
        {
          styleId: "Normal",
          type: "paragraph",
          default: true,
          name: "Normal",
        },
        {
          styleId: "FontePadrao",
          type: "character",
          default: true,
          name: "Default Paragraph Font",
          rPr: { fontFamily: { ascii: "Cambria", hAnsi: "Cambria" } },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "plain" }] }],
    };

    const blocks = toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {});

    expect(firstRun(blocks).fontFamily).toBe("Cambria");
  });

  test("table conditionals without rPr do not override paragraph run defaults", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "Normal",
          type: "paragraph",
          default: true,
          name: "Normal",
          rPr: { fontFamily: { ascii: "Arial", hAnsi: "Arial" } },
        },
        {
          styleId: "FontePadrao",
          type: "character",
          default: true,
          name: "Default Paragraph Font",
          rPr: { fontFamily: { ascii: "Cambria", hAnsi: "Cambria" } },
        },
        {
          styleId: "BandedTable",
          type: "table",
          name: "Banded Table",
          tblStylePr: [
            {
              type: "firstRow",
              tcPr: { shading: { fill: { rgb: "EEEEEE" } } },
            },
          ],
        },
      ],
    };
    const table: Table = {
      type: "table",
      formatting: { styleId: "BandedTable", look: { firstRow: true } },
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  formatting: { styleId: "Normal" },
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "first row" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const blocks = toFlowBlocks(
      toProseDoc(
        {
          package: {
            document: { content: [table] },
            styles,
          },
        },
        { styles },
      ),
      {},
    );

    expect(firstTableRun(blocks).fontFamily).toBe("Arial");
  });

  test("default table style supplies cell margins when table has no style ID", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "TableNormal",
          type: "table",
          default: true,
          name: "Normal Table",
          tblPr: {
            cellMargins: {
              left: { value: 144, type: "dxa" },
              right: { value: 288, type: "dxa" },
            },
          },
        },
      ],
    };
    const table: Table = {
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
                      content: [{ type: "text", text: "cell" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };
    const document: Document = {
      package: {
        document: { content: [table] },
        styles,
      },
    };

    const pmDoc = toProseDoc(document, { styles });
    const tableNode = pmDoc.firstChild;

    expect(tableNode?.attrs["cellMargins"]).toEqual({
      left: 144,
      right: 288,
    });

    const tableBlock = toFlowBlocks(pmDoc, {})[0];
    expect(tableBlock?.kind).toBe("table");
    if (tableBlock?.kind === "table") {
      expect(tableBlock.rows[0]?.cells[0]?.padding?.left).toBeCloseTo(9.6, 1);
      expect(tableBlock.rows[0]?.cells[0]?.padding?.right).toBeCloseTo(19.2, 1);
    }
  });

  test("explicit table styles do not fall back to the default table style borders or margins", () => {
    const gridBorder = {
      style: "single" as const,
      size: 4,
      color: { rgb: "000000" },
    };
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "DefaultGrid",
          type: "table",
          default: true,
          name: "Default Grid",
          tblPr: {
            borders: {
              top: gridBorder,
              bottom: gridBorder,
              left: gridBorder,
              right: gridBorder,
              insideH: gridBorder,
              insideV: gridBorder,
            },
            cellMargins: {
              left: { value: 144, type: "dxa" },
              right: { value: 288, type: "dxa" },
            },
          },
        },
        {
          styleId: "Borderless",
          type: "table",
          name: "Borderless",
        },
      ],
    };
    const table: Table = {
      type: "table",
      formatting: { styleId: "Borderless" },
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
                      content: [{ type: "text", text: "cell" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };
    const document: Document = {
      package: {
        document: { content: [table] },
        styles,
      },
    };

    const tableNode = toProseDoc(document, { styles }).firstChild;
    const rowNode = tableNode?.firstChild;
    const cellNode = rowNode?.firstChild;

    expect(tableNode?.attrs["cellMargins"]).toBeNull();
    expect(cellNode?.attrs["borders"]).toBeNull();
  });

  test("default table style supplies conditional formatting when table has no style ID", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "DefaultGrid",
          type: "table",
          default: true,
          name: "Default Grid",
          tblStylePr: [
            {
              type: "wholeTable",
              tcPr: { shading: { fill: { rgb: "D9EAF7" } } },
              rPr: { bold: true },
            },
          ],
        },
      ],
    };
    const table: Table = {
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
                      content: [{ type: "text", text: "default styled" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };
    const document: Document = {
      package: {
        document: { content: [table] },
        styles,
      },
    };

    const tableBlock = toFlowBlocks(toProseDoc(document, { styles }), {})[0];
    expect(tableBlock?.kind).toBe("table");
    if (tableBlock?.kind === "table") {
      const cell = tableBlock.rows.at(0)?.cells.at(0);
      const paragraph = cell?.blocks.at(0) as ParagraphBlock | undefined;
      const run = paragraph?.runs.at(0) as TextRun | undefined;

      expect(cell?.background).toBe("#D9EAF7");
      expect(run?.bold).toBe(true);
    }
  });

  test("style-less tables use built-in TableNormal side padding", () => {
    const table: Table = {
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
                      content: [{ type: "text", text: "cell" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };
    const document: Document = {
      package: {
        document: { content: [table] },
      },
    };

    const tableBlock = toFlowBlocks(toProseDoc(document), {})[0];
    expect(tableBlock?.kind).toBe("table");
    if (tableBlock?.kind === "table") {
      const padding = tableBlock.rows.at(0)?.cells.at(0)?.padding;
      expect(padding?.top).toBe(0);
      expect(padding?.right).toBeCloseTo(7.2, 1);
      expect(padding?.bottom).toBe(0);
      expect(padding?.left).toBeCloseTo(7.2, 1);
    }
  });

  test("explicit zero cell margins fall through to table defaults", () => {
    const table: Table = {
      type: "table",
      formatting: {
        cellMargins: {
          left: { value: 144, type: "dxa" },
          right: { value: 288, type: "dxa" },
        },
      },
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              formatting: {
                margins: {
                  left: { value: 0, type: "dxa" },
                  right: { value: 0, type: "dxa" },
                },
              },
              content: [
                {
                  type: "paragraph",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "cell" }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };
    const document: Document = {
      package: {
        document: { content: [table] },
      },
    };

    const tableBlock = toFlowBlocks(toProseDoc(document), {})[0];
    expect(tableBlock?.kind).toBe("table");
    if (tableBlock?.kind === "table") {
      const padding = tableBlock.rows.at(0)?.cells.at(0)?.padding;
      expect(padding?.left).toBeCloseTo(9.6, 1);
      expect(padding?.right).toBeCloseTo(19.2, 1);
    }
  });
});

describe("table style paragraph spacing cascade (cell paragraphs)", () => {
  // Mirrors the real-world TableNormal -> TableGrid cascade: docDefaults
  // give every paragraph 200-twip space-after and 1.15x (276) line spacing;
  // TableGrid (based on TableNormal) zeroes both out for its cells. Per
  // ECMA-376 §17.7.2 the table style's own `w:pPr` sits between docDefaults
  // and the paragraph's style chain/direct formatting — see
  // `resolveParagraphStyleInTable` in styleResolver.ts.
  const styles: StyleDefinitions = {
    docDefaults: {
      pPr: { spaceAfter: 200, lineSpacing: 276, lineSpacingRule: "auto" },
    },
    styles: [
      {
        styleId: "TableNormal",
        type: "table",
        name: "Table Normal",
      },
      {
        styleId: "TableGrid",
        type: "table",
        name: "Table Grid",
        basedOn: "TableNormal",
        pPr: { spaceAfter: 0, lineSpacing: 240, lineSpacingRule: "auto" },
      },
      {
        styleId: "CellStyle",
        type: "paragraph",
        name: "Cell Style",
        pPr: { spaceAfter: 100 },
      },
    ],
  };

  function buildDocument(): Document {
    const table: Table = {
      type: "table",
      formatting: { styleId: "TableGrid" },
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              // (a) No pStyle, no direct pPr — must resolve to the table
              // style's spacing, not docDefaults. This is the bug.
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  content: [{ type: "run", content: [{ type: "text", text: "no style" }] }],
                },
              ],
            },
            {
              // (b) Explicit paragraph style sets its own space-after — the
              // style must win over the table overlay for that field.
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  formatting: { styleId: "CellStyle" },
                  content: [{ type: "run", content: [{ type: "text", text: "styled" }] }],
                },
              ],
            },
            {
              // (c) Direct pPr spacing — must win over the table overlay
              // (and over any style).
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  formatting: { spaceAfter: 50 },
                  content: [{ type: "run", content: [{ type: "text", text: "direct" }] }],
                },
              ],
            },
          ],
        },
      ],
    };
    const baselineParagraph: Paragraph = {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "outside the table" }] }],
    };
    return {
      package: {
        document: { content: [baselineParagraph, table] },
        styles,
      },
    };
  }

  test("cell paragraph with no style and no direct pPr resolves to the table style's spacing", () => {
    const pmDoc = toProseDoc(buildDocument(), { styles });
    const table = pmDoc.content.child(1);
    const cell = table.content.child(0).content.child(0);
    const paragraph = cell.content.child(0);

    expect(paragraph.attrs["spaceAfter"]).toBe(0);
    expect(paragraph.attrs["lineSpacing"]).toBe(240);
    expect(paragraph.attrs["lineSpacingRule"]).toBe("auto");
  });

  test("an explicit paragraph style still wins over the table overlay", () => {
    const pmDoc = toProseDoc(buildDocument(), { styles });
    const table = pmDoc.content.child(1);
    const cell = table.content.child(0).content.child(1);
    const paragraph = cell.content.child(0);

    // The style only sets spaceAfter — it wins for that field...
    expect(paragraph.attrs["spaceAfter"]).toBe(100);
    // ...but the table overlay still supplies fields the style leaves unset.
    expect(paragraph.attrs["lineSpacing"]).toBe(240);
  });

  test("direct paragraph formatting wins over both the table overlay and any style", () => {
    const pmDoc = toProseDoc(buildDocument(), { styles });
    const table = pmDoc.content.child(1);
    const cell = table.content.child(0).content.child(2);
    const paragraph = cell.content.child(0);

    expect(paragraph.attrs["spaceAfter"]).toBe(50);
  });

  test("a non-table paragraph is unaffected and still resolves to docDefaults", () => {
    const pmDoc = toProseDoc(buildDocument(), { styles });
    const baselineParagraph = pmDoc.content.child(0);

    expect(baselineParagraph.attrs["spaceAfter"]).toBe(200);
    expect(baselineParagraph.attrs["lineSpacing"]).toBe(276);
  });
});
