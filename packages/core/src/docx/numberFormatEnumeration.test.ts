/**
 * Every `ST_NumberFormat` member survives a parse, a save and the editor
 * projection, and nothing outside the enumeration is ever written.
 *
 * `NUMBER_FORMAT_VALUES` used to be wrong in both directions. It omitted
 * `bahtText`, `dollarText` and `custom`, so a level, a note or a section
 * written with one lost its format at parse time. And it carried three
 * `decimalZero{3,4,5}` members the format does not declare: the parser minted
 * them from a `custom` format's pad width, and the serializer wrote them back
 * as a `w:val` no consumer can read.
 *
 * Those three are now `CounterFormat`, folio's render vocabulary, which is
 * never serialized. The pad width reaches the renderer through the level's
 * `@w:format` instead.
 */

import { describe, expect, test } from "bun:test";

import { COUNTER_FORMAT_VALUES, NUMBER_FORMAT_VALUES } from "../types/documentEnumValues";

import { computeListRendering, counterFormatOf, parseNumbering } from "./numberingParser";
import { serializeNumberingXml } from "./serializer/numberingSerializer";
import { parseSectionProperties } from "./sectionParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const numberingXml = (numFmt: string, format?: string): string =>
  `<w:numbering ${WORD_NAMESPACE}>
     <w:abstractNum w:abstractNumId="0">
       <w:lvl w:ilvl="0">
         <w:numFmt w:val="${numFmt}"${format === undefined ? "" : ` w:format="${format}"`}/>
         <w:lvlText w:val="%1."/>
       </w:lvl>
     </w:abstractNum>
     <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
   </w:numbering>`;

const levelNumFmt = (xml: string) => parseNumbering(xml).getLevel(1, 0);

const pageNumberFormat = (fmt: string): string | undefined => {
  const sectPr = parseXmlDocument(
    `<w:sectPr ${WORD_NAMESPACE}><w:pgNumType w:fmt="${fmt}"/></w:sectPr>`,
  );
  if (!sectPr) {
    throw new Error("fixture did not parse");
  }
  return parseSectionProperties(sectPr).pageNumbering?.format;
};

describe("ST_NumberFormat", () => {
  // The sweeps below run over the model's own list, so they shrink with it.
  // `scripts/narrowed-enum-schema-types.test.ts` is what holds that list to the
  // enumeration; this names the three members it used to be missing and the
  // three it used to carry that the format does not declare.
  test("bahtText, dollarText and custom are members, and no decimalZeroN is", () => {
    expect(NUMBER_FORMAT_VALUES).toContain("bahtText");
    expect(NUMBER_FORMAT_VALUES).toContain("dollarText");
    expect(NUMBER_FORMAT_VALUES).toContain("custom");
    for (const invented of ["decimalZero3", "decimalZero4", "decimalZero5"] as const) {
      expect(NUMBER_FORMAT_VALUES).not.toContain(invented);
      expect(COUNTER_FORMAT_VALUES).toContain(invented);
    }
  });

  test.each(NUMBER_FORMAT_VALUES)("a level's w:numFmt reads %s", (numFmt) => {
    expect(levelNumFmt(numberingXml(numFmt))?.numFmt).toBe(numFmt);
  });

  test.each(NUMBER_FORMAT_VALUES)("a section's w:pgNumType reads %s", (numFmt) => {
    expect(pageNumberFormat(numFmt)).toBe(numFmt);
  });

  test.each(NUMBER_FORMAT_VALUES)("%s survives parse, save and parse", (numFmt) => {
    const saved = serializeNumberingXml(parseNumbering(numberingXml(numFmt)).definitions);
    expect(saved).toContain(`<w:numFmt w:val="${numFmt}"/>`);
    expect(levelNumFmt(saved)?.numFmt).toBe(numFmt);
  });

  test.each(NUMBER_FORMAT_VALUES)("%s counts in something the renderer knows", (numFmt) => {
    const rendering = computeListRendering(
      { numId: 1, ilvl: 0 },
      parseNumbering(numberingXml(numFmt)),
    );
    expect(COUNTER_FORMAT_VALUES).toContain(rendering?.numFmt);
  });

  test("a custom format carries its w:format through a save", () => {
    const level = levelNumFmt(numberingXml("custom", "0001, 0002, 0003"));
    expect(level?.numFmt).toBe("custom");
    expect(level?.numFmtFormat).toBe("0001, 0002, 0003");

    const saved = serializeNumberingXml(
      parseNumbering(numberingXml("custom", "0001, 0002, 0003")).definitions,
    );
    expect(saved).toContain('<w:numFmt w:val="custom" w:format="0001, 0002, 0003"/>');
    expect(levelNumFmt(saved)?.numFmtFormat).toBe("0001, 0002, 0003");
  });

  test("a custom pad width counts in the decimalZero family, and is never written", () => {
    const level = levelNumFmt(numberingXml("custom", "0001, 0002, 0003"));
    if (!level) {
      throw new Error("expected level 0");
    }
    expect(counterFormatOf(level)).toBe("decimalZero4");

    const saved = serializeNumberingXml(
      parseNumbering(numberingXml("custom", "0001, 0002, 0003")).definitions,
    );
    expect(saved).not.toContain("decimalZero4");
  });

  // The writer takes `w:val` straight from the model, so the model holding only
  // schema tokens is what keeps an unreadable one out of the part.
  test("every w:val a save can write is a member of the enumeration", () => {
    const saved = serializeNumberingXml({
      abstractNums: NUMBER_FORMAT_VALUES.map((numFmt, index) => ({
        abstractNumId: index,
        levels: [{ ilvl: 0, numFmt, lvlText: "%1." }],
      })),
      nums: [],
    });
    for (const [, value] of saved.matchAll(/<w:numFmt w:val="(?<val>[^"]*)"/gu)) {
      expect(NUMBER_FORMAT_VALUES).toContain(value);
    }
  });

  test("a token outside the enumeration does not become a format", () => {
    expect(levelNumFmt(numberingXml("futureFmt"))?.numFmt).toBe("decimal");
    expect(pageNumberFormat("futureFmt")).toBeUndefined();
  });
});
