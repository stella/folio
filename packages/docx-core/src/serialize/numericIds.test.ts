import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(10_000));

import { assertValidOoxmlNumericIds, InvalidOoxmlNumericIdError } from "./numericIds";

const WORD_NAMESPACES = [
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
];
const decimalCases = [
  (id: string) => `<x:comment x:id="${id}"/>`,
  (id: string) => `<x:commentRangeStart x:id="${id}"/>`,
  (id: string) => `<x:commentRangeEnd x:id="${id}"/>`,
  (id: string) => `<x:commentReference x:id="${id}"/>`,
  (id: string) => `<x:ins x:id="${id}"/>`,
  (id: string) => `<x:bookmarkStart x:id="${id}" x:name="marker"/>`,
  (id: string) => `<x:bookmarkEnd x:id="${id}"/>`,
  (id: string) => `<x:footnote x:id="${id}"/>`,
  (id: string) => `<x:endnoteReference x:id="${id}"/>`,
  (id: string) => `<x:abstractNum x:abstractNumId="${id}"/>`,
  (id: string) => `<x:num x:numId="${id}"/>`,
  (id: string) => `<x:numPicBullet x:numPicBulletId="${id}"/>`,
  (id: string) => `<x:numId x:val="${id}"/>`,
  (id: string) => `<x:abstractNumId x:val="${id}"/>`,
  (id: string) => `<x:lvlPicBulletId x:val="${id}"/>`,
  (id: string) => `<x:sdtPr><x:id x:val="${id}"/></x:sdtPr>`,
];

describe("numeric OOXML identifier writer guard", () => {
  test("rejects the complete decimal-id family beyond the signed 32-bit domain", () => {
    fc.assert(
      fc.property(fc.integer({ min: 2_147_483_648, max: Number.MAX_SAFE_INTEGER }), (id) => {
        for (const namespace of WORD_NAMESPACES) {
          for (const body of decimalCases) {
            expect(() =>
              assertValidOoxmlNumericIds(
                `<x:document xmlns:x="${namespace}">${body(String(id))}</x:document>`,
                "word/document.xml",
              ),
            ).toThrow(InvalidOoxmlNumericIdError);
          }
        }
      }),
      propertyConfig({ numRuns: 20 }),
    );
  });

  test("preserves valid decimal ids including note separators and domain endpoints", () => {
    for (const namespace of WORD_NAMESPACES) {
      for (const id of [-2_147_483_648, -1, 0, 1, 2_147_483_647]) {
        for (const body of decimalCases) {
          expect(() =>
            assertValidOoxmlNumericIds(
              `<x:document xmlns:x="${namespace}">${body(String(id))}</x:document>`,
              "word/document.xml",
            ),
          ).not.toThrow();
        }
      }
    }
  });

  test("rejects fractional, nonfinite and underflow ids", () => {
    for (const id of ["1.5", "NaN", "Infinity", "-2147483649", "", "no-id"]) {
      expect(() =>
        assertValidOoxmlNumericIds(
          `<x:comment xmlns:x="${WORD_NAMESPACES[0]}" x:id="${id}"/>`,
          "word/comments.xml",
        ),
      ).toThrow(InvalidOoxmlNumericIdError);
    }
  });

  test("resolves inner namespace bindings and numeric character references", () => {
    const xml =
      '<root xmlns:x="urn:unrelated"><x:comment x:id="99999999999"/><x:comment xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" x:id="&#50;147483648"/></root>';
    expect(() => assertValidOoxmlNumericIds(xml, "word/comments.xml")).toThrow(
      InvalidOoxmlNumericIdError,
    );
  });

  test("allows permission string ids and unrelated vocabularies", () => {
    const xml = `<x:document xmlns:x="${WORD_NAMESPACES[0]}" xmlns:other="urn:other"><x:permStart x:id="everyone"/><x:permEnd x:id="99999999999"/><other:comment other:id="99999999999"/></x:document>`;
    expect(() => assertValidOoxmlNumericIds(xml, "word/document.xml")).not.toThrow();
  });

  test("drawing identifiers use the unsigned 32-bit domain", () => {
    for (const element of ["docPr", "cNvPr"]) {
      for (const namespace of [
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "http://purl.oclc.org/ooxml/drawingml/picture",
      ]) {
        expect(() =>
          assertValidOoxmlNumericIds(
            `<p:${element} xmlns:p="${namespace}" id="4294967295"/>`,
            "word/document.xml",
          ),
        ).not.toThrow();
        for (const id of ["4294967296", "-1", "1.5"]) {
          expect(() =>
            assertValidOoxmlNumericIds(
              `<p:${element} xmlns:p="${namespace}" id="${id}"/>`,
              "word/document.xml",
            ),
          ).toThrow(InvalidOoxmlNumericIdError);
        }
      }
    }
  });
});
