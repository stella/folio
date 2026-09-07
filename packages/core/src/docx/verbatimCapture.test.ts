import { describe, expect, test } from "bun:test";

import { captureVerbatimXml, UntranslatableStrictNamespaceError } from "./verbatimCapture";
import { parseXml, type XmlElement } from "./xmlParser";

const STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const STRICT_A = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_WP = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

/** Parse one fragment as the parser does, under the root bindings a part declares. */
const fragment = (xml: string): XmlElement => {
  const root = parseXml(xml).elements?.[0];
  if (root === undefined) {
    throw new Error("fixture did not parse");
  }
  return root;
};

const capture = (xml: string): string => captureVerbatimXml(fragment(xml));

describe("verbatim capture", () => {
  test("leaves Transitional markup byte for byte", () => {
    const xml =
      '<w:tcPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:tcW w:w="3117" w:type="dxa"/></w:tcPr>';
    expect(capture(xml)).toBe(xml);
  });

  test("rebinds a Strict namespace the fragment declares", () => {
    expect(capture(`<w:tcPr xmlns:w="${STRICT_W}"/>`)).toBe(
      '<w:tcPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    );
  });

  test("writes a Strict length as the number its Transitional type counts", () => {
    const captured = capture(
      `<w:tcPr xmlns:w="${STRICT_W}"><w:tcW w:w="155.85pt" w:type="dxa"/>` +
        '<w:tcMar><w:top w:w="0.1in" w:type="dxa"/></w:tcMar></w:tcPr>',
    );
    expect(captured).toContain('<w:tcW w:w="3117" w:type="dxa"/>');
    expect(captured).toContain('<w:top w:w="144" w:type="dxa"/>');
  });

  test("writes a Strict percentage in the fraction its Transitional type counts", () => {
    // `w:tcW` counts fiftieths of a percent; DrawingML counts thousandths.
    expect(
      capture(`<w:tcPr xmlns:w="${STRICT_W}"><w:tcW w:w="50%" w:type="pct"/></w:tcPr>`),
    ).toContain('<w:tcW w:w="2500" w:type="pct"/>');
    expect(capture(`<a:alpha xmlns:a="${STRICT_A}" val="60%"/>`)).toBe(
      '<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="60000"/>',
    );
  });

  test("converts a drawing extension the schema graph does not carry", () => {
    // `wp14` has one namespace in both classes; only its Strict ancestry says
    // the producer spelled the value the Strict way.
    const captured = capture(
      `<wp:anchor xmlns:wp="${STRICT_WP}" xmlns:wp14="${WP14}">` +
        '<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>12.5%</wp14:pctWidth></wp14:sizeRelH>' +
        "</wp:anchor>",
    );
    expect(captured).toContain("<wp14:pctWidth>12500</wp14:pctWidth>");
  });

  test("leaves a percentage in running text alone", () => {
    const captured = capture(`<w:p xmlns:w="${STRICT_W}"><w:r><w:t>60%</w:t></w:r></w:p>`);
    expect(captured).toContain("<w:t>60%</w:t>");
  });

  test("leaves a Transitional subtree inside a Strict fragment alone", () => {
    const captured = capture(
      `<w:p xmlns:w="${STRICT_W}">` +
        '<w:tcW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:w="155.85pt" w:type="dxa"/>' +
        "</w:p>",
    );
    expect(captured).toContain('w:w="155.85pt"');
  });

  test("refuses a Strict namespace with no Transitional counterpart", () => {
    expect(() =>
      capture('<x:thing xmlns:x="http://purl.oclc.org/ooxml/invented/vocabulary"/>'),
    ).toThrow(UntranslatableStrictNamespaceError);
  });
});
