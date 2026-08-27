/**
 * Attribute values that reach the serializer from parsed input must be
 * escaped, and OOXML enum attributes must be narrowed where they are read.
 *
 * `w:rFonts` carries four theme-font references. `w:asciiTheme` was narrowed
 * through `FontThemeSchema` at both parse sites while its three siblings were
 * copied verbatim and interpolated into the saved attribute without escaping,
 * so a value carrying a quote reopened the element. Same shape in
 * `w:background` and `w:pgNumType`, whose values are plain parsed strings.
 */
import { describe, expect, test } from "bun:test";

import type { SectionProperties, TextFormatting } from "../../types/document";
import { parseRunProperties } from "../runParser";
import { parseXmlDocument } from "../xmlParser";
import type { XmlElement } from "../xmlParser";
import { serializeTextFormatting } from "./runSerializer";
import { serializeSectionProperties } from "./sectionPropertiesSerializer";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const parseRPr = (rPrXml: string): TextFormatting | undefined => {
  const root = parseXmlDocument(`<w:rPr xmlns:w="${W_NS}">${rPrXml}</w:rPr>`) as XmlElement | null;
  return parseRunProperties(root, null);
};

const THEME_FONT_ATTRIBUTES = [
  "w:asciiTheme",
  "w:hAnsiTheme",
  "w:eastAsiaTheme",
  "w:cstheme",
] as const;

describe("theme font references", () => {
  test("round-trip keeps every recognised reference", () => {
    const formatting = parseRPr(
      '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"' +
        ' w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>',
    );

    expect(formatting?.fontFamily).toMatchObject({
      asciiTheme: "minorHAnsi",
      hAnsiTheme: "minorHAnsi",
      eastAsiaTheme: "minorEastAsia",
      csTheme: "minorBidi",
    });

    const xml = serializeTextFormatting(formatting);
    expect(xml).toContain('w:hAnsiTheme="minorHAnsi"');
    expect(xml).toContain('w:eastAsiaTheme="minorEastAsia"');
    expect(xml).toContain('w:cstheme="minorBidi"');
  });

  test("a reference outside the enum is dropped at parse", () => {
    for (const attribute of THEME_FONT_ATTRIBUTES) {
      const formatting = parseRPr(`<w:rFonts ${attribute}="notATheme"/>`);
      expect(serializeTextFormatting(formatting)).not.toContain("notATheme");
    }
  });

  test("a reference set on the model is escaped on the way out", () => {
    const xml = serializeTextFormatting({
      fontFamily: {
        hAnsiTheme: 'minorHAnsi" w:hAnsi="Injected',
        eastAsiaTheme: 'a"b',
        csTheme: "a&b",
      },
    });

    expect(xml).not.toContain('w:hAnsi="Injected');
    expect(xml).toContain("&quot;");
    expect(xml).toContain("&amp;");
    expect(parseXmlDocument(`<w:rPr xmlns:w="${W_NS}">${xml}</w:rPr>`)).not.toBeNull();
  });
});

describe("section property attributes", () => {
  test("page background and chapter separator are escaped", () => {
    const props = {
      background: {
        color: { rgb: 'FFFFFF" w:themeColor="accent1' },
        themeTint: 'a"b',
        themeShade: "a&b",
      },
      pageNumbering: { chapterSeparator: 'hyphen" w:start="9' },
    } as const satisfies SectionProperties;

    const xml = serializeSectionProperties(props);

    expect(xml).not.toContain('w:themeColor="accent1"');
    expect(xml).not.toContain('w:start="9"');
    expect(xml).toContain("&quot;");
    expect(xml).toContain("&amp;");
    expect(parseXmlDocument(`<w:body xmlns:w="${W_NS}">${xml}</w:body>`)).not.toBeNull();
  });
});
