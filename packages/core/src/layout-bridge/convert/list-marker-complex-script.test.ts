import { describe, expect, test } from "bun:test";

import { computeListRendering, parseNumbering } from "../../docx/numberingParser";
import { expectParagraphAttrs } from "../../prosemirror/attrs";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

const numberingXml = `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="7">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="arabicAlpha"/><w:lvlText w:val="%1."/>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:cs="Traditional Arabic"/>
        <w:b/><w:bCs w:val="0"/><w:i w:val="0"/><w:iCs/>
        <w:sz w:val="20"/><w:szCs w:val="32"/><w:rtl w:val="0"/><w:cs w:val="0"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
</w:numbering>`;

const parsedDocument = (): Document => {
  const numbering = parseNumbering(numberingXml);
  const rendering = computeListRendering({ numId: 3, ilvl: 0 }, numbering);
  if (!rendering) {
    throw new TypeError("Expected list rendering");
  }
  return {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            formatting: { numPr: { numId: 3, ilvl: 0 } },
            listRendering: { ...rendering, marker: "ا." },
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "بند" }],
              },
            ],
          },
        ],
      },
      numbering: numbering.definitions,
    },
  };
};

describe("complex-script list-marker pipeline", () => {
  test("preserves one typed formatting value through PM and the Document round trip", () => {
    const prose = toProseDoc(parsedDocument());
    const paragraph = prose.firstChild;
    if (!paragraph) {
      throw new TypeError("Expected ProseMirror paragraph");
    }
    expect(expectParagraphAttrs(paragraph).listMarkerFormatting).toEqual({
      fontFamily: { ascii: "Arial", cs: "Traditional Arabic" },
      bold: true,
      boldCs: false,
      italic: false,
      italicCs: true,
      fontSize: 20,
      fontSizeCs: 32,
      rtl: false,
      cs: false,
    });

    const roundTripped = fromProseDoc(prose);
    const modelParagraph = roundTripped.package.document.content.at(0);
    if (modelParagraph?.type !== "paragraph") {
      throw new TypeError("Expected round-tripped paragraph");
    }
    expect(modelParagraph.listRendering?.markerFormatting).toEqual(
      expectParagraphAttrs(paragraph).listMarkerFormatting,
    );
  });

  test("converts OOXML units and independent CS values once at the Flow boundary", () => {
    const prose = toProseDoc(parsedDocument());
    const paragraph = toFlowBlocks(prose, {}).find((block) => block.kind === "paragraph");
    if (!paragraph) {
      throw new TypeError("Expected Flow paragraph");
    }

    expect(paragraph.attrs?.listMarkerFormatting).toEqual({
      fontFamily: "Arial",
      complexScriptFontFamily: "Traditional Arabic",
      fontSize: 10,
      complexScriptFontSize: 16,
      bold: true,
      complexScriptBold: false,
      italic: false,
      complexScriptItalic: true,
      rtl: false,
      forceComplexScript: false,
    });
  });
});
