/**
 * The escaper's contract, over the whole of UTF-16.
 *
 * Three properties, each stated against something other than the escaper's
 * own implementation:
 *
 *   1. Well-formedness. Whatever goes in, the document that comes out holds
 *      only characters XML 1.0 §2.2 admits, and two independent parsers
 *      (`fast-xml-parser`, which folio reads packages with, and happy-dom's
 *      `DOMParser`) both open it.
 *   2. Value preservation. Attribute and element values come back as the
 *      sanitised input, through both parsers. happy-dom implements §3.3.3
 *      attribute-value normalisation, so it is what catches a literal
 *      newline written into an attribute.
 *   3. Idempotence of the sanitiser: running it twice changes nothing more.
 *
 * The generator deliberately over-samples what a naive escaper gets wrong:
 * the whole UTF-16 code-unit range (so unpaired surrogates appear), the C0
 * controls, `]]>`, the five metacharacters and astral characters.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { XMLParser } from "fast-xml-parser";
import { Window } from "happy-dom";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  escapeXmlAttribute,
  escapeXmlText,
  hasIllegalXmlCharacters,
  sanitizeXmlCharacters,
} from "./xmlEscape";

setDefaultTimeout(propertyTestTimeout(30_000));

const ILLEGAL_XML_CHARACTER =
  // eslint-disable-next-line no-control-regex -- the point of the assertion is the control range.
  /[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u;

const ILLEGAL_XML_CHARACTERS = new RegExp(ILLEGAL_XML_CHARACTER.source, "gu");

/**
 * What XML 1.0 §2.2 alone says the escaper may return: the input minus every
 * character no document can hold. Restated from the spec rather than imported,
 * so the round-trip below is not checked against the code that produced it.
 */
const legalCharactersOf = (value: string): string => value.replace(ILLEGAL_XML_CHARACTERS, "");

const hazards = [
  "]]>",
  "&",
  "<",
  ">",
  '"',
  "'",
  "&amp;",
  "&#0;",
  "\t",
  "\n",
  "\r",
  "\r\n",
  "\u0000",
  "\u0008",
  "\u000B",
  "\u000C",
  "\u001F",
  "\uFFFE",
  "\uFFFF",
  "\uD800",
  "\uDFFF",
  "\u{1F4C4}",
  "\u00A0",
  " ",
];

/** Arbitrary strings over the full UTF-16 range, with the hazards seeded in. */
const hostileString = fc
  .array(
    fc.oneof(
      { weight: 3, arbitrary: fc.constantFrom(...hazards) },
      { weight: 2, arbitrary: fc.string({ unit: "binary", maxLength: 6 }) },
      { weight: 1, arbitrary: fc.string({ unit: "grapheme", maxLength: 6 }) },
    ),
    { maxLength: 12 },
  )
  .map((pieces) => pieces.join(""));

const fxp = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false,
  textNodeName: "#text",
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  htmlEntities: true,
});

type ParsedValues = { attribute: string; text: string };

const readWithFastXmlParser = (xml: string): ParsedValues => {
  const nodes = fxp.parse(xml) as Record<string, unknown>[];
  const root = nodes[0] as { r: { "#text"?: string }[]; ":@": Record<string, string> };
  return {
    attribute: root[":@"]?.["a"] ?? "",
    text: root.r.map((child) => child["#text"] ?? "").join(""),
  };
};

const happyDom = new Window();

const readWithDomParser = (xml: string): ParsedValues => {
  const document = new happyDom.DOMParser().parseFromString(xml, "text/xml");
  const root = document.documentElement;
  expect(root.tagName).toBe("r");
  return { attribute: root.getAttribute("a") ?? "", text: root.textContent ?? "" };
};

const documentFor = (value: string) =>
  `<r a="${escapeXmlAttribute(value)}">${escapeXmlText(value)}</r>`;

describe("the escaper always produces a well-formed document", () => {
  test("the output holds only characters XML 1.0 admits", () => {
    fc.assert(
      fc.property(hostileString, (value) => {
        expect(documentFor(value)).not.toMatch(ILLEGAL_XML_CHARACTER);
      }),
      propertyConfig({ numRuns: 500 }),
    );
  });

  test("both parsers open it", () => {
    fc.assert(
      fc.property(hostileString, (value) => {
        const xml = documentFor(value);
        expect(() => readWithFastXmlParser(xml)).not.toThrow();
        expect(() => readWithDomParser(xml)).not.toThrow();
      }),
      propertyConfig({ numRuns: 500 }),
    );
  });
});

describe("the escaper preserves the value it was given", () => {
  test("attribute and text read back as the sanitised input, in both parsers", () => {
    fc.assert(
      fc.property(hostileString, (value) => {
        const expected = legalCharactersOf(value);
        const xml = documentFor(value);
        expect(readWithFastXmlParser(xml)).toEqual({ attribute: expected, text: expected });
        expect(readWithDomParser(xml)).toEqual({ attribute: expected, text: expected });
      }),
      propertyConfig({ numRuns: 500 }),
    );
  });
});

describe("sanitizeXmlCharacters", () => {
  test("is idempotent and leaves nothing illegal behind", () => {
    fc.assert(
      fc.property(hostileString, (value) => {
        const once = sanitizeXmlCharacters(value);
        expect(sanitizeXmlCharacters(once)).toBe(once);
        expect(hasIllegalXmlCharacters(once)).toBe(false);
      }),
      propertyConfig({ numRuns: 500 }),
    );
  });
});
