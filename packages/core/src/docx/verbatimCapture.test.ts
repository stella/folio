import { describe, expect, test } from "bun:test";

import {
  captureVerbatimXml,
  captureSourceProfileXml,
  UntranslatableStrictNamespaceError,
} from "./verbatimCapture";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";
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

describe("generated source-profile capture", () => {
  test("converts Strict percentage slots and graphic vocabularies without changing user values or frozen input", () => {
    const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const picture = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    const xml = `<w:tbl xmlns:w="${word}" xmlns:a="${drawing}" xmlns:vendor="urn:vendor">
      <w:tblPr><w:tblW w:w="2500" w:type="pct" vendor:w="2500"/><w:tblInd w:w="720" w:type="dxa"/></w:tblPr>
      <w:tr><w:tc><w:p><w:r><w:t label="${word}" vendor:uri="${picture}">2500 ${picture}</w:t></w:r></w:p></w:tc></w:tr>
      <a:graphic><a:graphicData uri="${picture}" vendor:uri="${picture}"><a:alpha val="60000" vendor:val="60000"/></a:graphicData></a:graphic>
      <vendor:graphicData uri="${picture}"/>
    </w:tbl>`;
    const element = fragment(xml);
    const before = structuredClone(element);
    const freeze = (value: unknown): void => {
      if (typeof value !== "object" || value === null || Object.isFrozen(value)) return;
      for (const child of Object.values(value)) freeze(child);
      Object.freeze(value);
    };
    freeze(element);
    const strict = captureSourceProfileXml(element, DOCX_CONFORMANCE_CLASSES.STRICT);
    expect(strict).toContain(`xmlns:w="${STRICT_W}"`);
    expect(strict).toContain(`xmlns:a="${STRICT_A}"`);
    expect(strict).toContain('<w:tblW w:w="50%" w:type="pct" vendor:w="2500"/>');
    expect(strict).toContain('<w:tblInd w:w="720" w:type="dxa"/>');
    expect(strict).toContain('<a:alpha val="60%" vendor:val="60000"/>');
    expect(strict).toContain('uri="http://purl.oclc.org/ooxml/drawingml/picture"');
    expect(strict).toContain(`label="${word}" vendor:uri="${picture}"`);
    expect(strict).toContain(`2500 ${picture}`);
    expect(strict).toContain(`<vendor:graphicData uri="${picture}"/>`);
    expect(strict).toContain(`vendor:uri="${picture}"`);
    const normalized = captureVerbatimXml(fragment(strict));
    expect(normalized).toContain('<w:tblW w:w="2500" w:type="pct" vendor:w="2500"/>');
    const authored = capture(
      `<w:tblW xmlns:w="${STRICT_W}" xmlns:vendor="urn:vendor" w:w="50%" w:type="pct" vendor:w="60%"/>`,
    );
    expect(authored).toContain('w:w="2500"');
    expect(authored).toContain('vendor:w="60%"');
    const authoredAlpha = capture(
      `<a:alpha xmlns:a="${STRICT_A}" xmlns:vendor="urn:vendor" val="60%" vendor:val="50%"/>`,
    );
    expect(authoredAlpha).toContain('val="60000"');
    expect(authoredAlpha).toContain('vendor:val="50%"');
    expect(normalized).toContain('<a:alpha val="60000" vendor:val="60000"/>');
    expect(element).toEqual(before);
    expect(captureSourceProfileXml(element, DOCX_CONFORMANCE_CLASSES.TRANSITIONAL)).toBe(xml);
  });
});
