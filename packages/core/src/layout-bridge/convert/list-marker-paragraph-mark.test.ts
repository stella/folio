import { describe, expect, test } from "bun:test";

import { computeListRendering, parseNumbering } from "../../docx/numberingParser";
import { resolveListMarkerFont } from "../../layout-engine/measure/listMarkerWidth";
import type { ParagraphBlock } from "../../layout-engine/types";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph, StyleDefinitions, TextFormatting } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

const numberingXml = (levelRPr = "") => `<w:numbering
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      ${levelRPr ? `<w:rPr>${levelRPr}</w:rPr>` : ""}
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const styles: StyleDefinitions = {
  styles: [
    {
      styleId: "Normal",
      type: "paragraph",
      default: true,
      name: "Normal",
      rPr: { fontFamily: { ascii: "Arial", hAnsi: "Arial" }, fontSize: 20 },
    },
    {
      styleId: "StrongBody",
      type: "paragraph",
      basedOn: "Normal",
      name: "Strong Body",
      rPr: { fontFamily: { ascii: "Arial", hAnsi: "Arial" }, fontSize: 20, bold: true },
    },
    {
      styleId: "BigMark",
      type: "character",
      name: "Big Mark",
      rPr: { fontSize: 48, color: { rgb: "2E74B5" } },
    },
    {
      styleId: "BoldMark",
      type: "character",
      name: "Bold Mark",
      rPr: { bold: true },
    },
  ],
};

type ListParagraphOptions = {
  styleId?: string;
  mark?: TextFormatting;
  levelRPr?: string;
  runFormatting?: TextFormatting;
  empty?: boolean;
  text?: string;
};

const listDocument = ({
  styleId = "Normal",
  mark,
  levelRPr,
  runFormatting = { bold: true, fontSize: 28 },
  empty = false,
  text = "Item",
}: ListParagraphOptions): Document => {
  const numbering = parseNumbering(numberingXml(levelRPr));
  const rendering = computeListRendering({ numId: 1, ilvl: 0 }, numbering);
  if (!rendering) {
    throw new TypeError("Expected list rendering");
  }
  const paragraph: Paragraph = {
    type: "paragraph",
    formatting: {
      styleId,
      numPr: { kind: "reference", numId: 1, ilvl: 0 },
      indentLeft: 720,
      hangingIndent: true,
      indentFirstLine: -360,
      ...(mark ? { runProperties: mark } : {}),
    },
    listRendering: { ...rendering, marker: "1." },
    content: empty
      ? []
      : [{ type: "run", formatting: runFormatting, content: [{ type: "text", text }] }],
  };
  return {
    package: {
      document: { content: [paragraph] },
      styles,
      numbering: numbering.definitions,
    },
  };
};

const listBlock = (options: ListParagraphOptions): ParagraphBlock => {
  const document = listDocument(options);
  const blocks = toFlowBlocks(toProseDoc(document, { styles }), { styles });
  const block = blocks.find((candidate) => candidate.kind === "paragraph");
  if (block?.kind !== "paragraph") {
    throw new TypeError("Expected a paragraph block");
  }
  return block;
};

describe("list marker formatting follows the paragraph mark (§17.9)", () => {
  test("ignores the first text run's formatting", () => {
    const font = resolveListMarkerFont(listBlock({}));

    expect(font.bold).not.toBe(true);
    expect(font.fontSize).toBe(10);
    expect(font.fontFamily).toBe("Arial");
  });

  test("takes bold, size and colour from direct paragraph-mark run properties", () => {
    const font = resolveListMarkerFont(
      listBlock({ mark: { bold: true, fontSize: 32, color: { rgb: "C00000" } } }),
    );

    expect(font.bold).toBe(true);
    expect(font.fontSize).toBe(16);
    expect(font.color?.toUpperCase()).toContain("C00000");
  });

  test("resolves a character style named by the paragraph mark's w:rStyle", () => {
    const font = resolveListMarkerFont(listBlock({ mark: { styleId: "BigMark" } }));

    expect(font.fontSize).toBe(24);
    expect(font.color?.toUpperCase()).toContain("2E74B5");
  });

  test("toggles bold when both the paragraph style and the mark's character style set it", () => {
    const font = resolveListMarkerFont(
      listBlock({ styleId: "StrongBody", mark: { styleId: "BoldMark" } }),
    );

    expect(font.bold).toBe(false);
  });

  test("inherits bold from the paragraph style", () => {
    const font = resolveListMarkerFont(listBlock({ styleId: "StrongBody" }));

    expect(font.bold).toBe(true);
  });

  test("numbering level w:rPr applies over the paragraph mark", () => {
    const font = resolveListMarkerFont(
      listBlock({
        mark: { bold: true, fontSize: 32 },
        levelRPr: `<w:b w:val="0"/><w:color w:val="00B050"/>`,
      }),
    );

    expect(font.bold).toBe(false);
    expect(font.fontSize).toBe(16);
    expect(font.color?.toUpperCase()).toContain("00B050");
  });

  test("an empty paragraph is sized by the mark's w:rStyle character style", () => {
    expect(listBlock({ mark: { styleId: "BigMark" }, empty: true }).attrs?.defaultFontSize).toBe(
      24,
    );
    const whitespaceOnly = listBlock({
      mark: { styleId: "BigMark" },
      runFormatting: {},
      text: " ",
    });
    expect(whitespaceOnly.attrs?.defaultFontSize).toBe(24);
  });
});
