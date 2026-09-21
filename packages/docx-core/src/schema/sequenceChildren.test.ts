import { describe, expect, test } from "bun:test";

import {
  assertSafePreservedMarkup,
  isSafePreservedChildXml,
  isWithinPreservedMarkupBudget,
  serializeSequenceChildren,
} from "./sequenceChildren";

describe("preserved sequence children", () => {
  test("accepts exactly one well-formed element", () => {
    const xml = '<w:mirrorIndents w:val="0"/>';

    expect(isSafePreservedChildXml(xml)).toBe(true);
    expect(
      serializeSequenceChildren({
        container: "paragraph-properties",
        modelled: [],
        preserved: { children: [{ index: 26, xml }] },
      }),
    ).toEqual([xml]);
  });

  test.each([
    "</w:pPr><w:sectPr/>",
    "<w:mirrorIndents/><w:sectPr/>",
    "<w:mirrorIndents/>text",
    '<?xml version="1.0"?><w:mirrorIndents/>',
    "<!DOCTYPE x><w:mirrorIndents/>",
    "<w:mirrorIndents>&evil;</w:mirrorIndents>",
    "<w:mirrorIndents>&#x0;</w:mirrorIndents>",
    "<w:mirrorIndents>&#1114112;</w:mirrorIndents>",
    "<x:mirrorIndents/>",
    '<w:mirrorIndents xmlns:w="urn:evil"/>',
    '<w:mirrorIndents xmlns:x=""/>',
    '<w:mirrorIndents xmlns:xml="urn:evil"/>',
    '<x:extension xmlns:x="urn:a" xmlns:y="urn:a" x:value="1" y:value="2"/>',
  ])("rejects markup that is not one child: %s", (xml) => {
    expect(isSafePreservedChildXml(xml)).toBe(false);
    expect(() =>
      serializeSequenceChildren({
        container: "paragraph-properties",
        modelled: [],
        preserved: { children: [{ index: 26, xml }] },
      }),
    ).toThrow("Preserved container markup must be one bounded, well-formed XML element");
  });

  test.each([
    "<w:mirrorIndents>&amp;&lt;&gt;&apos;&quot;</w:mirrorIndents>",
    "<w:mirrorIndents>&#9;&#999999;&#xA;&#x10ffff;</w:mirrorIndents>",
    '<x:extension xmlns:x="urn:example"><x:child x:value="kept"/></x:extension>',
    '<w:mirrorIndents xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
  ])("accepts safe references and namespace bindings: %s", (xml) => {
    expect(isSafePreservedChildXml(xml)).toBe(true);
  });

  test("bounds the aggregate preserved markup", () => {
    const child = { index: 0, xml: "<w:mirrorIndents/>" };
    const tooMany = Array.from({ length: 4097 }, () => child);
    const largeXml = `<w:mirrorIndents>${"x".repeat(1024 * 1024 - 40)}</w:mirrorIndents>`;
    const tooWide = Array.from({ length: 5 }, () => ({ index: 0, xml: largeXml }));

    expect(isWithinPreservedMarkupBudget(tooMany)).toBe(false);
    expect(isWithinPreservedMarkupBudget(tooWide)).toBe(false);
    expect(() => assertSafePreservedMarkup({ children: tooWide })).toThrow(
      "Preserved container markup exceeds its aggregate resource budget",
    );
    expect(() =>
      serializeSequenceChildren({
        container: "paragraph-properties",
        modelled: [],
        preserved: { children: tooMany },
      }),
    ).toThrow("Preserved container markup exceeds its aggregate resource budget");
  });
});
