import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { validateSchemaAttributes } from "./schemaAttributes";

const TRANSITIONAL = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const ID_ELEMENTS = [
  "comment",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "ins",
  "del",
  "moveFrom",
  "moveTo",
  "bookmarkStart",
  "bookmarkEnd",
  "footnote",
  "endnote",
  "footnoteReference",
  "endnoteReference",
] as const;

// Mutate the integer boundary dimension in every annotation family, including
// inherited slots; enum-only checks cannot detect timestamp-sized ids.
describe("schema-derived story attribute oracle", () => {
  test.each([TRANSITIONAL, STRICT])("checks all identity families under %s", (namespace) => {
    for (const tag of ID_ELEMENTS) {
      const part = (id: string) =>
        `<x:${tag} xmlns:x="${namespace}" x:id="${id}" x:author="Fixture" x:name="Bookmark"/>`;
      for (const id of ["-2147483648", "-1", "0", "+2147483647"]) {
        expect(validateSchemaAttributes(part(id))).toBeNull();
      }
      for (const id of ["-2147483649", "2147483648", "1790884800000", "1.5", "1e3", "NaN"]) {
        expect(validateSchemaAttributes(part(id))).toContain("Invalid attribute value");
      }
    }
  });

  test("distinguishes foreign attributes and rebound prefixes", () => {
    expect(
      validateSchemaAttributes(
        `<x:commentReference xmlns:x="${TRANSITIONAL}" xmlns:f="urn:foreign" x:id="1" f:id="1790884800000"/>`,
      ),
    ).toBeNull();
    expect(
      validateSchemaAttributes(
        `<w:document xmlns:w="${TRANSITIONAL}"><w:body><w:commentReference xmlns:w="urn:foreign" w:id="1790884800000"/></w:body></w:document>`,
      ),
    ).toBeNull();
    expect(
      validateSchemaAttributes(
        `<x:commentReference xmlns:x="${TRANSITIONAL}" xmlns:f="urn:foreign" f:id="1"/>`,
      ),
    ).toContain("Missing required attribute");
  });

  test("checks inherited required slots and enum values", () => {
    expect(validateSchemaAttributes(`<w:comment xmlns:w="${TRANSITIONAL}" w:id="1"/>`)).toContain(
      "Missing required attribute",
    );
    expect(validateSchemaAttributes(`<w:jc xmlns:w="${TRANSITIONAL}" w:val="center"/>`)).toBeNull();
    expect(validateSchemaAttributes(`<w:jc xmlns:w="${TRANSITIONAL}" w:val="centre"/>`)).toContain(
      "Invalid attribute value",
    );
    expect(validateSchemaAttributes(`<w:b xmlns:w="${TRANSITIONAL}" w:val="false"/>`)).toBeNull();
    expect(validateSchemaAttributes(`<w:b xmlns:w="${TRANSITIONAL}" w:val="maybe"/>`)).toContain(
      "Invalid attribute value",
    );
  });

  test("checks schema numeric facets inside lexical unions", () => {
    expect(
      validateSchemaAttributes(`<w:rPr xmlns:w="${TRANSITIONAL}"><w:w w:val="600"/></w:rPr>`),
    ).toBeNull();
    expect(
      validateSchemaAttributes(`<w:rPr xmlns:w="${TRANSITIONAL}"><w:w w:val="601"/></w:rPr>`),
    ).toContain("Invalid attribute value");
    expect(
      validateSchemaAttributes(`<w:rPr xmlns:w="${TRANSITIONAL}"><w:w w:val="601%"/></w:rPr>`),
    ).toContain("Invalid attribute value");
    expect(validateSchemaAttributes(`<w:sz xmlns:w="${STRICT}" w:val="12pt"/>`)).toBeNull();
  });

  test("checks builtin integer ranges without floating point rounding", () => {
    expect(
      validateSchemaAttributes(
        `<w:rPr xmlns:w="${TRANSITIONAL}"><w:sz w:val="18446744073709551615"/></w:rPr>`,
      ),
    ).toBeNull();
    expect(
      validateSchemaAttributes(
        `<w:rPr xmlns:w="${TRANSITIONAL}"><w:sz w:val="18446744073709551616"/></w:rPr>`,
      ),
    ).toContain("Invalid attribute value");
    expect(
      validateSchemaAttributes(`<w:rPr xmlns:w="${TRANSITIONAL}"><w:sz w:val="-1"/></w:rPr>`),
    ).toContain("Invalid attribute value");
  });

  test.each([TRANSITIONAL, STRICT])(
    "applies schema whitespace per union member under %s",
    (namespace) => {
      const part = (tag: string, value: string) =>
        tag === "sz"
          ? `<w:rPr xmlns:w="${namespace}"><w:sz w:val="${value}"/></w:rPr>`
          : `<w:${tag} xmlns:w="${namespace}" w:val="${value}"/>`;
      // ST_Jc derives from xs:string: ASCII padding is part of the enum value.
      for (const value of [" center ", "&#x9;center&#xA;", "&#xA0;center&#xA0;"]) {
        expect(validateSchemaAttributes(part("jc", value))).toContain("Invalid attribute value");
      }
      for (const value of [" 1 ", "&#x9;1&#xD;&#xA;"]) {
        expect(
          validateSchemaAttributes(`<w:commentReference xmlns:w="${namespace}" w:id="${value}"/>`),
        ).toBeNull();
      }
      for (const value of ["&#xA0;1&#xA0;", "&#x2003;1&#x2003;"]) {
        expect(
          validateSchemaAttributes(`<w:commentReference xmlns:w="${namespace}" w:id="${value}"/>`),
        ).toContain("Invalid attribute value");
      }
      // ST_OnOff unions a collapsing xs:boolean with a preserving string enum.
      expect(validateSchemaAttributes(part("b", " true "))).toBeNull();
      expect(validateSchemaAttributes(part("b", " on "))).toContain("Invalid attribute value");
      expect(validateSchemaAttributes(part("b", "&#xA0;true&#xA0;"))).toContain(
        "Invalid attribute value",
      );
      expect(validateSchemaAttributes(part("sz", " 12 "))).toBeNull();
      expect(validateSchemaAttributes(part("sz", " 12pt "))).toContain("Invalid attribute value");
    },
  );

  test(
    "annotation ids preserve validity across numeric XML encodings",
    () => {
      fc.assert(
        fc.property(
          fc.integer(),
          fc.constantFrom(TRANSITIONAL, STRICT),
          fc.constantFrom("decimal", "hex"),
          (id, namespace, encoding) => {
            const encoded = [...String(id)]
              .map((character) =>
                encoding === "decimal"
                  ? `&#${character.codePointAt(0)};`
                  : `&#x${character.codePointAt(0)?.toString(16)};`,
              )
              .join("");
            expect(
              validateSchemaAttributes(
                `<x:commentReference xmlns:x="${namespace}" x:id="${encoded}"/>`,
              ),
            ).toBeNull();
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(5_000),
  );

  test("decodes XML entities before checking enums", () => {
    expect(
      validateSchemaAttributes(`<w:jc xmlns:w="${TRANSITIONAL}" w:val="c&#101;nter"/>`),
    ).toBeNull();
  });

  test.each([
    "<w:p>",
    "<p/><q/>",
    "<p x:id='1'/>",
    "<p><q></p>",
    "<!DOCTYPE p [<!ENTITY a 'x'>]><p/>",
  ])("rejects malformed or entity-declaring XML: %s", (xml) => {
    expect(validateSchemaAttributes(xml)).not.toBeNull();
  });

  test("accepts well-formed package XML outside story schemas", () => {
    expect(
      validateSchemaAttributes(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
      ),
    ).toBeNull();
    expect(
      validateSchemaAttributes(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      ),
    ).toBeNull();
  });
});
