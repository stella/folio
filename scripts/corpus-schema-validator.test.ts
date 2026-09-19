import { describe, expect, test } from "bun:test";

import {
  SCHEMA_VIOLATION_KINDS,
  loadSchemaGraph,
  validateOoxmlPart,
} from "./lib/corpus-schema-validator";

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MARKUP_COMPATIBILITY_NAMESPACE =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FOREIGN_NAMESPACE = "urn:example:not-an-ooxml-namespace";

const graph = await loadSchemaGraph();

const validate = (xml: string) => validateOoxmlPart({ graph, xml });

/** A `w:document` part around `body`, in the conventional `w:` spelling. */
const document = (body: string, rootAttributes = ""): string =>
  `<w:document xmlns:w="${WORDPROCESSINGML_NAMESPACE}"${rootAttributes}>` +
  `<w:body>${body}</w:body></w:document>`;

describe("validateOoxmlPart", () => {
  test("accepts a minimal document", () => {
    expect(validate(document("<w:p><w:r><w:t>text</w:t></w:r></w:p>"))).toEqual([]);
  });

  test("accepts a table", () => {
    const table =
      "<w:tbl>" +
      '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4675"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p/></w:tc></w:tr>' +
      "</w:tbl>";
    expect(validate(document(table))).toEqual([]);
  });

  test("reports an element no content model declares", () => {
    const violations = validate(document("<w:p><w:pPr><w:notAProperty/></w:pPr></w:p>"));
    expect(violations).toHaveLength(1);
    expect(violations.at(0)).toMatchObject({
      kind: SCHEMA_VIOLATION_KINDS.unknownElement,
      path: "document/body/p/pPr/notAProperty",
      name: `{${WORDPROCESSINGML_NAMESPACE}}notAProperty`,
    });
  });

  test("reports an attribute no declaration allows, but not a declared one", () => {
    const violations = validate(document('<w:p w:notAnAttribute="1"/>'));
    expect(violations).toHaveLength(1);
    expect(violations.at(0)).toMatchObject({
      kind: SCHEMA_VIOLATION_KINDS.unknownAttribute,
      path: "document/body/p",
      name: `{${WORDPROCESSINGML_NAMESPACE}}notAnAttribute`,
    });
    expect(validate(document('<w:p w:rsidR="00AA00FF"/>'))).toEqual([]);
  });

  test("reports a required attribute that is absent", () => {
    const violations = validate(document("<w:p><w:pPr><w:jc/></w:pPr></w:p>"));
    expect(violations).toHaveLength(1);
    expect(violations.at(0)).toMatchObject({
      kind: SCHEMA_VIOLATION_KINDS.missingRequiredAttribute,
      name: `{${WORDPROCESSINGML_NAMESPACE}}val`,
    });
  });

  test("reports a value outside an enumeration, but not a member of it", () => {
    const violations = validate(document('<w:p><w:pPr><w:jc w:val="sideways"/></w:pPr></w:p>'));
    expect(violations).toHaveLength(1);
    expect(violations.at(0)).toMatchObject({
      kind: SCHEMA_VIOLATION_KINDS.badEnumValue,
      path: "document/body/p/pPr/jc",
      name: `{${WORDPROCESSINGML_NAMESPACE}}val`,
    });
    expect(validate(document('<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>'))).toEqual([]);
  });

  test("accepts every ECMA-376 boolean spelling on an ST_OnOff attribute", () => {
    // `w:b/@w:val` is ST_OnOff: a union of xsd:boolean and ST_OnOff1 (on|off).
    for (const value of ["true", "false", "1", "0", "on", "off"]) {
      const part = document(`<w:p><w:pPr><w:rPr><w:b w:val="${value}"/></w:rPr></w:pPr></w:p>`);
      expect(validate(part)).toEqual([]);
    }
  });

  test("tolerates markup-compatibility and foreign markup", () => {
    const alternateContent =
      "<mc:AlternateContent>" +
      '<mc:Choice Requires="wps"><w:invented/></mc:Choice>' +
      "<mc:Fallback/>" +
      "</mc:AlternateContent>";
    const foreign = `<z:thing xmlns:z="${FOREIGN_NAMESPACE}"><z:inner z:attr="x"/></z:thing>`;
    const part = document(
      `<w:p>${alternateContent}</w:p>${foreign}`,
      ` xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}" mc:Ignorable="w14 wp14"`,
    );
    expect(validate(part)).toEqual([]);
  });

  test("resolves by namespace URI, not by prefix", () => {
    const withX =
      `<x:document xmlns:x="${WORDPROCESSINGML_NAMESPACE}"><x:body>` +
      '<x:p><x:pPr><x:jc x:val="sideways"/><x:notAProperty/></x:pPr></x:p>' +
      "</x:body></x:document>";
    const withW = document('<w:p><w:pPr><w:jc w:val="sideways"/><w:notAProperty/></w:pPr></w:p>');
    expect(validate(withX)).toEqual(validate(withW));
    expect(validate(withX)).toHaveLength(2);
  });

  test("resolves an undecorated default namespace declaration", () => {
    const withDefault =
      `<document xmlns="${WORDPROCESSINGML_NAMESPACE}">` +
      "<body><p><r><t>text</t></r></p></body></document>";
    expect(validate(withDefault)).toEqual([]);
  });

  test("reports a root that is not a global element declaration", () => {
    const violations = validate(`<w:notARoot xmlns:w="${WORDPROCESSINGML_NAMESPACE}"/>`);
    expect(violations).toEqual([
      {
        kind: SCHEMA_VIOLATION_KINDS.unknownRoot,
        path: "notARoot",
        name: `{${WORDPROCESSINGML_NAMESPACE}}notARoot`,
        detail: "root element is not a global element declaration",
      },
    ]);
  });

  test("says nothing about a part rooted in foreign markup", () => {
    expect(validate(`<z:custom xmlns:z="${FOREIGN_NAMESPACE}"><z:field/></z:custom>`)).toEqual([]);
  });

  test("caps the returned violations at the limit", () => {
    const part = document("<w:p><w:pPr><w:one/><w:two/><w:three/><w:four/></w:pPr></w:p>");
    expect(validate(part)).toHaveLength(4);
    expect(validateOoxmlPart({ graph, xml: part, limit: 2 })).toHaveLength(2);
  });

  test("produces identical violations for two parts with the same defect", () => {
    const first = validate(
      document('<w:p w:rsidR="00AA00FF"><w:pPr><w:jc w:val="x"/></w:pPr></w:p>'),
    );
    const second = validate(document('<w:p><w:pPr><w:jc w:val="sideways"/></w:pPr></w:p>'));
    expect(first).toEqual(second);
  });
});

describe("out-of-order children", () => {
  // CT_PPr is a real pure-sequence type: CT_PPrBase declares 33 optional
  // elements in one xsd:sequence, and the extension appends rPr, sectPr and
  // pPrChange. Every name is distinct, so ordinals are unambiguous.
  test("reports the first child that precedes an earlier-declared sibling", () => {
    const violations = validate(
      document('<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:line="240"/></w:pPr></w:p>'),
    );
    expect(violations).toEqual([
      {
        kind: SCHEMA_VIOLATION_KINDS.outOfOrderChild,
        path: "document/body/p/pPr/spacing",
        name: `{${WORDPROCESSINGML_NAMESPACE}}spacing`,
        detail: "child precedes a sibling the sequence declares earlier",
      },
    ]);
  });

  test("accepts the declared order, including the extension's own particles", () => {
    const paragraphProperties =
      "<w:pPr>" +
      '<w:pStyle w:val="Heading1"/>' +
      '<w:spacing w:before="240"/>' +
      '<w:ind w:left="720"/>' +
      '<w:jc w:val="both"/>' +
      "<w:rPr><w:b/></w:rPr>" +
      "</w:pPr>";
    expect(validate(document(`<w:p>${paragraphProperties}</w:p>`))).toEqual([]);
  });

  test("stays silent where a choice compositor makes order a guess", () => {
    // CT_R mixes EG_RPr with the EG_RunInnerContent choice, so a run whose rPr
    // trails its text is not something this check may call a defect.
    expect(
      validate(document("<w:p><w:r><w:t>text</w:t><w:rPr><w:b/></w:rPr></w:r></w:p>")),
    ).toEqual([]);
  });
});

describe("extension attributes", () => {
  const WORDML_2010 = "http://schemas.microsoft.com/office/word/2010/wordml";

  /**
   * Every recent Word decorates `w:p` with `w14:paraId` and `w14:textId`, in a
   * namespace this schema does not describe. Word accepts them, so folio must,
   * and calling them schema violations would report the producer on almost
   * every real package rather than reporting folio.
   */
  test("tolerates an attribute in a namespace the schema does not describe", () => {
    const xml = document(
      `<w:p w14:paraId="12345678" w14:textId="77777777"><w:r><w:t>text</w:t></w:r></w:p>`,
      ` xmlns:w14="${WORDML_2010}"`,
    );
    expect(validate(xml)).toEqual([]);
  });

  test("still reports an undeclared attribute in a namespace the schema does describe", () => {
    const xml = document(`<w:p w:invented="1"><w:r><w:t>text</w:t></w:r></w:p>`);
    expect(validate(xml).map(({ kind }) => kind)).toEqual(["unknown-attribute"]);
  });
});
