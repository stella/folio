/**
 * A content control's `w:sdtEndPr` survives a rebuild, not only a replay.
 *
 * The element was captured bytes and nothing else, so every control folio
 * rebuilt — one an edit touched, one a full repack wrote — lost its end mark
 * and the run properties on it. The census recorded
 * `sdt|CT_SdtRun/sdtEndPr` and `sdt|CT_SdtBlock/sdtEndPr` as
 * `serialized-only-via-verbatim-replay` for exactly that.
 *
 * `CT_SdtEndPr` declares `w:rPr` and nothing else, so the record's presence
 * is the element's: `<w:sdtEndPr/>` with no `w:rPr` is what Word writes for
 * most controls and says something an absent element does not.
 */

import { describe, expect, test } from "bun:test";

import { parseSdtProperties } from "./sdtProperties";
import { serializeSdtEndProperties } from "./serializer/sdtPropertiesSerializer";
import { OOXML_NAMESPACE_SCOPE, parseXml, type XmlElement } from "./xmlParser";

const elementOf = (xml: string): XmlElement => {
  const parsed = parseXml(xml, OOXML_NAMESPACE_SCOPE).elements?.[0];
  if (!parsed) {
    throw new Error(`not an element: ${xml}`);
  }
  return parsed;
};

/** The properties as a control folio rebuilds them holds them: no bytes left. */
const withoutCaptures = (xml: string): ReturnType<typeof parseSdtProperties> => {
  const properties = parseSdtProperties(elementOf("<w:sdtPr/>"), elementOf(xml));
  properties.rawEndPropertiesXml = undefined;
  return properties;
};

describe("w:sdtEndPr", () => {
  test("an end mark with run properties is rebuilt from the record", () => {
    expect(
      serializeSdtEndProperties(withoutCaptures("<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>")),
    ).toBe("<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>");
  });

  test("an end mark that declares nothing still comes back", () => {
    expect(serializeSdtEndProperties(withoutCaptures("<w:sdtEndPr/>"))).toBe("<w:sdtEndPr/>");
  });

  test("a control that never had one writes none", () => {
    const properties = parseSdtProperties(elementOf("<w:sdtPr/>"), null);
    expect(properties.endProperties).toBeUndefined();
    expect(serializeSdtEndProperties(properties)).toBe("");
  });

  test("the captured bytes win while they last, so element order is the source's", () => {
    const properties = parseSdtProperties(
      elementOf("<w:sdtPr/>"),
      elementOf('<w:sdtEndPr><w:rPr><w:rStyle w:val="Emphasis"/><w:b/></w:rPr></w:sdtEndPr>'),
    );
    expect(serializeSdtEndProperties(properties)).toBe(
      '<w:sdtEndPr><w:rPr><w:rStyle w:val="Emphasis"/><w:b/></w:rPr></w:sdtEndPr>',
    );
  });
});
