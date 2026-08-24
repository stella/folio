import { describe, expect, test } from "bun:test";

import { computeListRendering, parseNumbering } from "./numberingParser";
import { serializeNumberingXml } from "./serializer/numberingSerializer";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const numberingXml = `<w:numbering ${W}>
  <w:abstractNum w:abstractNumId="7">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="arabicAlpha"/>
      <w:lvlText w:val="%1."/>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Traditional Arabic"/>
        <w:b/><w:bCs w:val="0"/>
        <w:i w:val="0"/><w:iCs/>
        <w:sz w:val="20"/><w:szCs w:val="32"/>
        <w:rtl w:val="0"/><w:cs w:val="0"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="arabicAbjad"/>
      <w:lvlText w:val="%2)"/>
      <w:rPr><w:rFonts w:cs="Arabic Typesetting"/><w:rtl/><w:cs/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
</w:numbering>`;

const expectedLevelZeroFormatting = {
  fontFamily: { ascii: "Arial", hAnsi: "Arial", cs: "Traditional Arabic" },
  bold: true,
  boldCs: false,
  italic: false,
  italicCs: true,
  fontSize: 20,
  fontSizeCs: 32,
  rtl: false,
  cs: false,
};

describe("numbering-level complex-script marker formatting", () => {
  test("uses the canonical run parser and preserves explicit false values", () => {
    const numbering = parseNumbering(numberingXml);

    expect(numbering.getLevel(3, 0)?.rPr).toMatchObject(expectedLevelZeroFormatting);
    expect(computeListRendering({ numId: 3, ilvl: 0 }, numbering)?.markerFormatting).toEqual(
      expectedLevelZeroFormatting,
    );
    expect(computeListRendering({ numId: 3, ilvl: 1 }, numbering)?.markerFormatting).toEqual({
      fontFamily: { cs: "Arabic Typesetting" },
      rtl: true,
      cs: true,
    });
  });

  test("serializes and reparses every independent CS property", () => {
    const serialized = serializeNumberingXml(parseNumbering(numberingXml).definitions);

    expect(serialized).toContain('w:cs="Traditional Arabic"');
    expect(serialized).toContain('<w:bCs w:val="0"/>');
    expect(serialized).toContain("<w:iCs/>");
    expect(serialized).toContain('<w:szCs w:val="32"/>');
    expect(serialized).toContain('<w:rtl w:val="0"/>');
    expect(serialized).toContain("<w:rtl/>");
    expect(serialized).toContain('<w:cs w:val="0"/>');
    expect(serialized).toContain("<w:cs/>");

    const reparsed = parseNumbering(serialized);
    expect(reparsed.getLevel(3, 0)?.rPr).toMatchObject(expectedLevelZeroFormatting);
    expect(reparsed.getLevel(3, 1)?.rPr).toMatchObject({
      fontFamily: { cs: "Arabic Typesetting" },
      rtl: true,
      cs: true,
    });
  });
});
