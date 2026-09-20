/**
 * Every `ST_TabJc` member survives a parse, a save and the editor projection.
 *
 * A tab stop whose `w:val` fails to narrow is not read as an unaligned stop:
 * the reader requires both a position and an alignment, so the whole
 * `<w:tab/>` leaves the model. `TAB_STOP_ALIGNMENT_VALUES` used to spell seven
 * of the nine, so a Strict-profile `<w:tab w:val="start" w:pos="720"/>` was
 * dropped, and with it the stop the paragraph's text hangs from.
 *
 * The sweep is over the generated list, so a schema refresh widens it.
 */

import { describe, expect, test } from "bun:test";

import type { TabStopAlignment } from "../types/document";
import { TAB_STOP_ALIGNMENT_VALUES } from "../types/documentEnumValues";

import { parseParagraphProperties } from "./paragraphParser";
import { serializeNumberingXml } from "./serializer/numberingSerializer";
import { parseNumbering } from "./numberingParser";
import { parseStyles } from "./styleParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const tabXml = (alignment: string): string =>
  `<w:tabs><w:tab w:val="${alignment}" w:pos="720"/></w:tabs>`;

const parseParagraphTier = (alignment: string): TabStopAlignment | undefined => {
  const pPr = parseXmlDocument(`<w:pPr ${WORD_NAMESPACE}>${tabXml(alignment)}</w:pPr>`);
  if (!pPr) {
    throw new Error("fixture did not parse");
  }
  return parseParagraphProperties(pPr, null)?.tabs?.at(0)?.alignment;
};

const parseStyleTier = (alignment: string): TabStopAlignment | undefined =>
  parseStyles(
    `<w:styles ${WORD_NAMESPACE}>
       <w:style w:type="paragraph" w:styleId="Tabbed">
         <w:name w:val="Tabbed"/>
         <w:pPr>${tabXml(alignment)}</w:pPr>
       </w:style>
     </w:styles>`,
    null,
  )
    .get("Tabbed")
    ?.pPr?.tabs?.at(0)?.alignment;

const numberingXml = (alignment: string): string =>
  `<w:numbering ${WORD_NAMESPACE}>
     <w:abstractNum w:abstractNumId="0">
       <w:lvl w:ilvl="0">
         <w:numFmt w:val="decimal"/>
         <w:lvlText w:val="%1."/>
         <w:pPr>${tabXml(alignment)}</w:pPr>
       </w:lvl>
     </w:abstractNum>
     <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
   </w:numbering>`;

const levelTabAlignment = (numberingXmlText: string): TabStopAlignment | undefined =>
  parseNumbering(numberingXmlText).getLevel(1, 0)?.pPr?.tabs?.at(0)?.alignment;

describe("ST_TabJc", () => {
  // The sweeps below run over the model's own list, so they shrink with it.
  // `scripts/narrowed-enum-schema-types.test.ts` is what holds that list to the
  // enumeration; this names the two members it used to be missing.
  test("start and end are members", () => {
    expect(TAB_STOP_ALIGNMENT_VALUES).toContain("start");
    expect(TAB_STOP_ALIGNMENT_VALUES).toContain("end");
  });

  test.each(TAB_STOP_ALIGNMENT_VALUES)("a paragraph's w:tab reads %s", (alignment) => {
    expect(parseParagraphTier(alignment)).toBe(alignment);
  });

  test.each(TAB_STOP_ALIGNMENT_VALUES)("a style's w:tab reads %s", (alignment) => {
    expect(parseStyleTier(alignment)).toBe(alignment);
  });

  test.each(TAB_STOP_ALIGNMENT_VALUES)(
    "a numbering level's %s tab survives parse, save and parse",
    (alignment) => {
      expect(levelTabAlignment(numberingXml(alignment))).toBe(alignment);

      const saved = serializeNumberingXml(parseNumbering(numberingXml(alignment)).definitions);
      expect(saved).toContain(`<w:tab w:val="${alignment}" w:pos="720"/>`);
      expect(levelTabAlignment(saved)).toBe(alignment);
    },
  );

  test("a stop whose alignment is outside the enumeration leaves no stop behind", () => {
    expect(parseParagraphTier("sideways")).toBeUndefined();
  });
});
