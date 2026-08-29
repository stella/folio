import { describe, expect, test } from "bun:test";

import type {
  ParagraphBlock,
  TableBlock as LayoutTableBlock,
  TableMeasure,
  TextRun,
} from "../../layout-engine/types";
import { layoutDocument } from "../../layout-engine";
import { parseStyleDefinitions } from "../../docx/styleParser";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { schema } from "../../prosemirror/schema";
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

function firstTableX(block: LayoutTableBlock): number | undefined {
  const measure: TableMeasure = {
    kind: "table",
    rows: [
      {
        cells: [{ blocks: [], width: 100, height: 20 }],
        height: 20,
      },
    ],
    columnWidths: [100],
    totalWidth: 100,
    totalHeight: 20,
  };
  const layout = layoutDocument([block], [measure], {
    pageSize: { w: 300, h: 200 },
    margins: { top: 40, right: 40, bottom: 40, left: 40 },
    pageGap: 20,
  });
  return layout.pages.at(0)?.fragments.find((fragment) => fragment.kind === "table")?.x;
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

  // Word writes the Arabic and Hebrew face into `w:cs`. That slot was parsed and
  // round-tripped but never reached layout, so complex-script text painted and
  // measured in the Latin face instead.
  test("the complex-script slot reaches layout", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "Normal",
          type: "paragraph",
          default: true,
          name: "Normal",
          rPr: { fontFamily: { ascii: "Calibri", cs: "Traditional Arabic" } },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Normal" },
      content: [
        {
          type: "run",
          content: [{ type: "text", text: "مكتب mixed" }],
        },
      ],
    };

    const run = firstRun(toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {}));

    expect(run.fontFamily).toBe("Calibri");
    expect(run.complexScriptFontFamily).toBe("Traditional Arabic");
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

  test("matching paragraph and character style toggles cancel", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "EmphasizedParagraph",
          type: "paragraph",
          name: "Emphasized Paragraph",
          rPr: {
            bold: true,
            boldCs: true,
            italic: true,
            italicCs: true,
            allCaps: true,
            emboss: true,
            hidden: true,
            imprint: true,
            outline: true,
            shadow: true,
            smallCaps: true,
            strike: true,
          },
        },
        {
          styleId: "EmphasizedRun",
          type: "character",
          name: "Emphasized Run",
          rPr: {
            bold: true,
            boldCs: true,
            italic: true,
            italicCs: true,
            allCaps: true,
            emboss: true,
            hidden: true,
            imprint: true,
            outline: true,
            shadow: true,
            smallCaps: true,
            strike: true,
          },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "EmphasizedParagraph" },
      content: [
        {
          type: "run",
          formatting: { styleId: "EmphasizedRun" },
          content: [{ type: "text", text: "regular" }],
        },
      ],
    };

    const proseDoc = toProseDoc(makeDoc(paragraph, styles), { styles });
    const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());
    const blocks = toFlowBlocks(clonedProseDoc, {});

    expect(firstRun(blocks)).toMatchObject({
      bold: false,
      complexScriptBold: false,
      italic: false,
      complexScriptItalic: false,
      allCaps: false,
      emboss: false,
      hidden: false,
      imprint: false,
      textOutline: false,
      textShadow: false,
      smallCaps: false,
      strike: false,
    });
  });

  test.each([
    ["inherited off, character off", false, false, false],
    ["inherited off, character on", false, true, true],
    ["inherited on, character off", true, false, true],
    ["inherited on, character on", true, true, false],
  ] as const)(
    "complex-script style toggles resolve %s after a JSON clone",
    (_label, inherited, character, expected) => {
      const styles: StyleDefinitions = {
        styles: [
          {
            styleId: "ParagraphEmphasis",
            type: "paragraph",
            name: "Paragraph Emphasis",
            rPr: { boldCs: inherited, italicCs: inherited },
          },
          {
            styleId: "CharacterEmphasis",
            type: "character",
            name: "Character Emphasis",
            rPr: { bold: true, boldCs: character, italic: true, italicCs: character },
          },
        ],
      };
      const paragraph: Paragraph = {
        type: "paragraph",
        formatting: { styleId: "ParagraphEmphasis" },
        content: [
          {
            type: "run",
            formatting: { styleId: "CharacterEmphasis" },
            content: [{ type: "text", text: "toggle matrix" }],
          },
        ],
      };

      const proseDoc = toProseDoc(makeDoc(paragraph, styles), { styles });
      const clonedProseDoc = schema.nodeFromJSON(proseDoc.toJSON());
      const run = firstRun(toFlowBlocks(clonedProseDoc, {}));

      expect(run).toMatchObject({
        bold: true,
        complexScriptBold: expected,
        complexScriptItalic: expected,
        italic: true,
      });
    },
  );

  test("a based-on chain contributes one style-level toggle value", () => {
    const styles = parseStyleDefinitions(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Base">
          <w:rPr><w:b/></w:rPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Derived">
          <w:basedOn w:val="Base"/>
          <w:rPr><w:b/></w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Derived" },
      content: [{ type: "run", content: [{ type: "text", text: "bold" }] }],
    };

    const blocks = toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {});

    expect(firstRun(blocks).bold).toBe(true);
  });

  test("false child toggles preserve inherited true values from parsed styles", () => {
    const toggleElements = [
      "b",
      "bCs",
      "i",
      "iCs",
      "caps",
      "emboss",
      "imprint",
      "outline",
      "shadow",
      "smallCaps",
      "strike",
      "vanish",
    ];
    const styles = parseStyleDefinitions(
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Base">
          <w:rPr>${toggleElements.map((name) => `<w:${name}/>`).join("")}</w:rPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Derived">
          <w:basedOn w:val="Base"/>
          <w:rPr>${toggleElements.map((name) => `<w:${name} w:val="0"/>`).join("")}</w:rPr>
        </w:style>
      </w:styles>`,
      null,
    );
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: { styleId: "Derived" },
      content: [{ type: "run", content: [{ type: "text", text: "inherited" }] }],
    };

    const run = firstRun(toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {}));

    expect(run).toMatchObject({
      bold: true,
      complexScriptBold: true,
      italic: true,
      complexScriptItalic: true,
      allCaps: true,
      emboss: true,
      hidden: true,
      imprint: true,
      textOutline: true,
      textShadow: true,
      smallCaps: true,
      strike: true,
    });
  });

  test("paragraph-mark booleans do not replace inherited paragraph-style booleans", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "EmphasizedClause",
          type: "paragraph",
          name: "Emphasized Clause",
          rPr: { bold: true },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      formatting: {
        styleId: "EmphasizedClause",
        runProperties: { bold: false },
      },
      content: [
        {
          type: "run",
          formatting: { bold: false },
          content: [{ type: "text", text: "Plain lead" }],
        },
        {
          type: "run",
          formatting: { fontSize: 22 },
          content: [{ type: "text", text: "Styled remainder" }],
        },
      ],
    };

    const blocks = toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {});
    const runs = firstParagraph(blocks).runs.filter((run) => run.kind === "text");

    expect(runs).toHaveLength(2);
    expect(runs[0]?.bold).toBe(false);
    expect(runs[1]?.bold).toBe(true);
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

  test("the default character style toggles above a paragraph-style off", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "PlainParagraph",
          type: "paragraph",
          default: true,
          name: "Plain Paragraph",
          rPr: { bold: false },
        },
        {
          styleId: "DefaultRun",
          type: "character",
          default: true,
          name: "Default Run",
          rPr: { bold: true },
        },
      ],
    };
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "bold" }] }],
    };

    const blocks = toFlowBlocks(toProseDoc(makeDoc(paragraph, styles), { styles }), {});

    expect(firstRun(blocks).bold).toBe(true);
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

  test("paragraph style toggles apply above an explicit table-style off", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "TableText",
          type: "paragraph",
          name: "Table Text",
          rPr: { bold: true },
        },
        {
          styleId: "PlainTable",
          type: "table",
          name: "Plain Table",
          rPr: { bold: false },
        },
      ],
    };
    const table: Table = {
      type: "table",
      formatting: { styleId: "PlainTable" },
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  formatting: { styleId: "TableText" },
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "bold" }],
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
      toProseDoc({ package: { document: { content: [table] }, styles } }, { styles }),
      {},
    );

    expect(firstTableRun(blocks).bold).toBe(true);
  });

  test("default table style supplies cell margins without authoring its zero indent", () => {
    const styles: StyleDefinitions = {
      styles: [
        {
          styleId: "TableNormal",
          type: "table",
          default: true,
          name: "Normal Table",
          tblPr: {
            indent: { value: 0, type: "dxa" },
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
    expect(tableNode?.attrs["_resolvedIndent"]).toBeNull();
    expect(tableNode?.attrs["_originalFormatting"]?.indent).toBeUndefined();

    const tableBlock = toFlowBlocks(pmDoc, {})[0];
    expect(tableBlock?.kind).toBe("table");
    if (tableBlock?.kind === "table") {
      expect(tableBlock.indent).toBeUndefined();
      expect(tableBlock.rows[0]?.cells[0]?.padding?.left).toBeCloseTo(9.6, 1);
      expect(tableBlock.rows[0]?.cells[0]?.padding?.right).toBeCloseTo(19.2, 1);
      expect(firstTableX(tableBlock)).toBeCloseTo(30.4, 1);
    }
  });

  test.each([
    {
      name: "direct",
      formatting: { indent: { value: 0, type: "dxa" as const } },
      styles: undefined,
    },
    {
      name: "explicit style",
      formatting: { styleId: "AuthoredTable" },
      styles: {
        styles: [
          {
            styleId: "AuthoredTable",
            type: "table" as const,
            name: "Authored Table",
            tblPr: { indent: { value: 0, type: "dxa" as const } },
          },
        ],
      },
    },
  ])("keeps $name zero table indentation authored", ({ formatting, styles }) => {
    const table: Table = {
      type: "table",
      formatting,
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
    };
    const document: Document = {
      package: {
        document: { content: [table] },
        ...(styles ? { styles } : {}),
      },
    };

    const tableBlock = toFlowBlocks(toProseDoc(document, styles ? { styles } : {}), {}).at(0);

    expect(tableBlock?.kind).toBe("table");
    if (tableBlock?.kind === "table") {
      expect(tableBlock.indent).toBe(0);
      expect(firstTableX(tableBlock)).toBe(40);
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

  test("explicit zero cell margins override table defaults", () => {
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
      expect(padding?.left).toBe(0);
      expect(padding?.right).toBe(0);
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
