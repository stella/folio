import { describe, expect, test } from "bun:test";

import { computeListRendering, parseNumbering } from "../../docx/numberingParser";
import { parseSettings } from "../../docx/settingsParser";
import {
  withFakeTextMeasure,
  fixedCharWidth,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { getListMarkerInlineWidth } from "../../layout-engine/measure/listMarkerWidth";
import type { ParagraphBlock } from "../../layout-engine/types";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document, DocumentSettings } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Level 0: w:ind left=1440 (96 px) hanging=720 (48 px), so the number starts at
// 48 px. The level also defines a stop at 1080 twips (72 px), inside the slot.
const numberingXml = `<w:numbering xmlns:w="${W}">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:pPr>
        <w:tabs><w:tab w:val="num" w:pos="1080"/></w:tabs>
        <w:ind w:left="1440" w:hanging="720"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const listDocument = (settings?: DocumentSettings): Document => {
  const numbering = parseNumbering(numberingXml);
  const rendering = computeListRendering({ numId: 1, ilvl: 0 }, numbering);
  if (!rendering) {
    throw new TypeError("Expected list rendering");
  }
  return {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            formatting: {
              numPr: { kind: "reference", numId: 1, ilvl: 0 },
              indentLeft: 1440,
              hangingIndent: true,
              indentFirstLine: -720,
            },
            listRendering: { ...rendering, marker: "1." },
            content: [{ type: "run", content: [{ type: "text", text: "Item" }] }],
          },
        ],
      },
      numbering: numbering.definitions,
      ...(settings ? { settings } : {}),
    },
  };
};

const listBlock = (document: Document): ParagraphBlock => {
  const block = toFlowBlocks(toProseDoc(document), {}).find(
    (candidate) => candidate.kind === "paragraph",
  );
  if (block?.kind !== "paragraph") {
    throw new TypeError("Expected a paragraph block");
  }
  return block;
};

const fakeMeasure = { charWidth: fixedCharWidth(10) };

describe("numbering level tab stops (w:lvl/w:pPr/w:tabs)", () => {
  test("reach the paragraph's tab stops", () => {
    const block = listBlock(listDocument());

    expect(block.attrs?.tabs?.map((tab) => tab.pos)).toEqual([1080]);
  });

  test("place the text after the number when the stop lies inside the hanging slot", () => {
    withFakeTextMeasure(() => {
      // Number "1." is 20 px wide: it ends at 68 px, before the 72 px stop.
      expect(getListMarkerInlineWidth(listBlock(listDocument()))).toBeCloseTo(72 - 48, 5);
    }, fakeMeasure);
  });

  test("w:doNotUseIndentAsNumberingTabStop reaches the list paragraph", () => {
    const settings = parseSettings(
      `<w:settings xmlns:w="${W}"><w:compat><w:doNotUseIndentAsNumberingTabStop/></w:compat></w:settings>`,
    );
    const block = listBlock(listDocument(settings));

    expect(block.attrs?.listNumberingTabIgnoresIndent).toBe(true);
    expect(listBlock(listDocument()).attrs?.listNumberingTabIgnoresIndent).toBeUndefined();
  });
});
