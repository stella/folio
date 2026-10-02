import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(10_000));

import {
  assertValidOoxmlNumericIds,
  InvalidOoxmlNumericIdError,
  mayContainInvalidOoxmlNumericIds,
} from "./numericIds";

const WORD_NAMESPACES = [
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
] as const;
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
  test("valid numeric ids and lexical relationship ids take the preflight fast path", () => {
    for (const namespace of WORD_NAMESPACES) {
      const xml = `<x:document xmlns:x="${namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:headerReference r:id="rId1"/><x:hyperlink r:id="rId2"/>${decimalCases.map((body) => body("2147483647")).join("")}</x:document>`;
      expect(mayContainInvalidOoxmlNumericIds(xml)).toBe(false);
      expect(mayContainInvalidOoxmlNumericIds(xml, "range")).toBe(false);
    }
  });

  test("range preflight sees identity attributes inside identity-value tags", () => {
    assertProperty(
      fc.property(
        fc.integer({ min: 2_147_483_648, max: 5_000_000_000 }),
        fc.constantFrom("id", "numId", "abstractNumId", "numPicBulletId"),
        fc.constantFrom("numId", "abstractNumId", "id"),
        (id, attribute, element) => {
          const xml = `<x:${element} xmlns:x="${WORD_NAMESPACES[0]}" x:val="1" x:${attribute}="${id}"/>`;
          expect(mayContainInvalidOoxmlNumericIds(xml, "range")).toBe(true);
          expect(() => assertValidOoxmlNumericIds(xml, "word/document.xml")).toThrow(
            InvalidOoxmlNumericIdError,
          );
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("namespace preflight preserves every numeric failure across shadows and encoded bindings", () => {
    const namespace = WORD_NAMESPACES[0];
    for (const declaration of [
      namespace,
      namespace.replace("wordprocessingml", "wordprocessing&#109;l"),
    ]) {
      for (const body of [
        '<r:comment r:id="2147483648"/>',
        '<r:comment r:id="&#50;147483648"/>',
        '<r:numId other:val="ignored" r:val="2147483648"/>',
      ]) {
        const xml = `<root xmlns:r="urn:relationship" xmlns:other="urn:other"><nested xmlns:r="${declaration}">${body}</nested></root>`;
        expect(mayContainInvalidOoxmlNumericIds(xml)).toBe(true);
        expect(() => assertValidOoxmlNumericIds(xml, "word/document.xml")).toThrow(
          InvalidOoxmlNumericIdError,
        );
      }
    }
  });

  test("range preflight preserves malformed lexical ids while schema preflight rejects them", () => {
    const xml = `<x:comment xmlns:x="${WORD_NAMESPACES[0]}" x:id="7invalid"/>`;
    expect(mayContainInvalidOoxmlNumericIds(xml, "range")).toBe(false);
    expect(mayContainInvalidOoxmlNumericIds(xml)).toBe(true);
    expect(() => assertValidOoxmlNumericIds(xml, "word/comments.xml")).toThrow(
      InvalidOoxmlNumericIdError,
    );
  });

  test("preflight and the writer agree over both numeric domains, aliases and encoded values", () => {
    assertProperty(
      fc.property(
        fc.integer({ min: -3_000_000_000, max: 5_000_000_000 }),
        fc.constantFrom("signed32", "unsigned32"),
        fc.constantFrom("x", "r", "alternate"),
        fc.constantFrom("decimal", "entity"),
        (id, domain, prefix, spelling) => {
          const value =
            spelling === "entity"
              ? String(id).replace(/\d/u, (digit) => `&#${digit.charCodeAt(0)};`)
              : String(id);
          const body =
            domain === "signed32"
              ? `<${prefix}:comment xmlns:${prefix}="${WORD_NAMESPACES[0]}" ${prefix}:id="${value}"/>`
              : `<${prefix}:docPr xmlns:${prefix}="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" id="${value}"/>`;
          const valid =
            domain === "signed32"
              ? id >= -2_147_483_648 && id <= 2_147_483_647
              : id >= 0 && id <= 4_294_967_295;
          // Exercise encoded namespace bindings and decimal/hex references
          // without changing the arbitrary that produced the pinned seed.
          for (const xml of [
            body,
            body.replace("wordprocessing", "wordprocess&#105;ng"),
            body.replace(
              /&#(\d+);/gu,
              (_reference, digits) => `&#x${Number(digits).toString(16)};`,
            ),
          ]) {
            if (valid) {
              expect(() => assertValidOoxmlNumericIds(xml, "word/document.xml")).not.toThrow();
              continue;
            }
            expect(mayContainInvalidOoxmlNumericIds(xml)).toBe(true);
            expect(mayContainInvalidOoxmlNumericIds(xml, "range")).toBe(true);
            expect(() => assertValidOoxmlNumericIds(xml, "word/document.xml")).toThrow(
              InvalidOoxmlNumericIdError,
            );
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  test("escaped entity text is not decoded twice into a numeric id or namespace", () => {
    for (const id of ["&amp;#48;", "&#38;#48;", "&#x26;#48;"]) {
      expect(() =>
        assertValidOoxmlNumericIds(
          `<x:comment xmlns:x="${WORD_NAMESPACES[0]}" x:id="${id}"/>`,
          "word/comments.xml",
        ),
      ).toThrow(InvalidOoxmlNumericIdError);
    }
    const namespace = WORD_NAMESPACES[0].replace("wordprocessingml", "wordprocessing&amp;#109;l");
    expect(() =>
      assertValidOoxmlNumericIds(
        `<x:comment xmlns:x="${namespace}" x:id="2147483648"/>`,
        "word/comments.xml",
      ),
    ).not.toThrow();
  });

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
