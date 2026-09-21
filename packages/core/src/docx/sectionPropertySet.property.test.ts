/**
 * A section's property set survives a rebuild, in the order the schema declares.
 *
 * `w:sectPr` had two answers to the same question and both lost. A child the
 * reader took no typed value from — `<w:cols/>`, `<w:pgSz/>`, an empty
 * `<w:footnotePr/>` — was dropped; a child no reader knew at all made
 * `serializeSectionProperties` return the empty string, which failed the whole
 * save rather than one property. The sink is the single answer: the element is
 * kept and the save goes through.
 *
 * The universe is the generated declared-child list, and the per-child samples
 * are total over it, so a child `CT_SectPr` gains fails the compile here
 * before it can be dropped at run time.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { CONTAINER_CHILDREN, type DeclaredChild } from "./containerChildren.gen";
import { parseSectionProperties } from "./sectionParser";
import { serializeSectionProperties } from "./serializer/sectionPropertiesSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const DECLARED = CONTAINER_CHILDREN["section-properties"];

/** Everything but the two references a package's own parts decide. */
const PROJECTED = DECLARED.filter(
  (name) => name !== "headerReference" && name !== "footerReference",
);

const SAMPLES = {
  headerReference: '<w:headerReference w:type="default" r:id="rId4"/>',
  footerReference: '<w:footerReference w:type="even" r:id="rId5"/>',
  footnotePr: '<w:footnotePr><w:pos w:val="pageBottom"/></w:footnotePr>',
  endnotePr: '<w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr>',
  type: '<w:type w:val="continuous"/>',
  pgSz: '<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>',
  pgMar:
    '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"' +
    ' w:header="708" w:footer="708" w:gutter="0"/>',
  paperSrc: '<w:paperSrc w:first="7" w:other="7"/>',
  pgBorders:
    '<w:pgBorders w:offsetFrom="page">' +
    '<w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/></w:pgBorders>',
  lnNumType: '<w:lnNumType w:countBy="1" w:start="1" w:restart="newPage"/>',
  pgNumType: '<w:pgNumType w:fmt="decimal" w:start="1"/>',
  cols: '<w:cols w:num="2" w:space="708" w:equalWidth="1"/>',
  formProt: '<w:formProt w:val="0"/>',
  vAlign: '<w:vAlign w:val="center"/>',
  noEndnote: "<w:noEndnote/>",
  titlePg: "<w:titlePg/>",
  textDirection: '<w:textDirection w:val="lrTb"/>',
  bidi: "<w:bidi/>",
  rtlGutter: '<w:rtlGutter w:val="0"/>',
  docGrid: '<w:docGrid w:type="lines" w:linePitch="360"/>',
  printerSettings: '<w:printerSettings r:id="rId6"/>',
  sectPrChange:
    '<w:sectPrChange w:id="9" w:author="Reviewer" w:date="2026-05-15T12:00:00Z">' +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:sectPrChange>',
} as const satisfies Record<DeclaredChild<"section-properties">, string>;

/**
 * Children the reader states nothing about.
 *
 * The first five are the shapes the census reported as never parsed: an
 * element every one of whose attributes is optional, written without any.
 * `w:jc` is declared for a table's properties and not for a section's, and
 * the last is a namespace the content model does not name.
 */
const UNREAD_CHILDREN = [
  "<w:cols/>",
  "<w:pgSz/>",
  "<w:paperSrc/>",
  "<w:pgNumType/>",
  "<w:footnotePr/>",
  '<w:jc w:val="center"/>',
  '<x:hint xmlns:x="urn:example:vendor" x:kind="layout"/>',
] as const;

const rebuild = (children: string): string => {
  const element = parseXmlDocument(
    `<w:sectPr xmlns:w="${W}" xmlns:r="${R}">${children}</w:sectPr>`,
  ) as XmlElement | null;
  if (!element) {
    throw new Error("the section fixture did not parse");
  }
  return serializeSectionProperties(parseSectionProperties(element));
};

/** The outer `w:sectPr`'s own children: a `w:sectPrChange` nests a second one. */
const childrenOf = (saved: string): string =>
  saved.slice(saved.indexOf(">") + 1, saved.lastIndexOf("</w:sectPr>"));

describe("a section's property set survives a rebuild", () => {
  test("every declared child comes back, and the save after it is a fixed point", () => {
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        const saved = rebuild(SAMPLES[name]);

        expect(saved).toContain(`<w:${name}`);
        expect(rebuild(childrenOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a child the reader takes nothing from keeps its bytes", () => {
    fc.assert(
      fc.property(fc.constantFrom(...UNREAD_CHILDREN), (child) => {
        const saved = rebuild(`<w:type w:val="continuous"/>${child}`);

        expect(saved).toContain(child);
        expect(rebuild(childrenOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the children come back in the order the content model declares", () => {
    const saved = rebuild(
      [...DECLARED]
        .reverse()
        .map((name) => SAMPLES[name])
        .join(""),
    );

    const positions = DECLARED.map((name) => saved.indexOf(`<w:${name}`));
    expect(positions.every((at) => at > -1)).toBe(true);
    expect(positions).toEqual([...positions].sort((left, right) => left - right));
  });

  test("a section that carries an unread child is still written", () => {
    // The serializer used to answer the empty string here, which the package
    // fidelity guard turns into a refused save: one unread property cost the
    // whole document rather than itself.
    expect(rebuild('<w:pgSz w:w="11906"/><w:cols/>')).toContain("<w:pgSz");
  });

  test("every declared child survives the editor projection", () => {
    // The record travels through ProseMirror whole, so the sink travels with
    // it; the assertion is that nothing on the way back rebuilds it from the
    // fields the editor happens to know.
    //
    // A header or footer reference is left out because it is not decided
    // here: folio drops a reference whose part the package does not carry,
    // and this document carries none. The survival census covers those two
    // against a package that has the part.
    fc.assert(
      fc.property(fc.constantFrom(...PROJECTED), (name) => {
        const element = parseXmlDocument(
          `<w:sectPr xmlns:w="${W}" xmlns:r="${R}">${SAMPLES[name]}</w:sectPr>`,
        ) as XmlElement;
        const finalSectionProperties = parseSectionProperties(element);
        const document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  formatting: {},
                  content: [{ type: "run", formatting: {}, content: [] }],
                },
              ],
              finalSectionProperties,
            },
          },
        } as never;
        const projected = fromProseDoc(toProseDoc(document), document).package.document
          .finalSectionProperties;

        expect(serializeSectionProperties(projected)).toBe(
          serializeSectionProperties(finalSectionProperties),
        );
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a section property change round-trips with its snapshot", () => {
    const saved = rebuild(SAMPLES.sectPrChange);

    expect(saved).toContain('<w:sectPrChange w:id="9" w:author="Reviewer"');
    expect(saved).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(rebuild(childrenOf(saved))).toBe(saved);
  });
});
