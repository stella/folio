/**
 * A table's property set survives a rebuild, in the order the schema declares.
 *
 * `w:tblPr` was walked by a reader per property with no branch for the rest:
 * `w:tblCaption`, `w:tblDescription` and the two band sizes had no model, and
 * `w:tblBorders`, `w:tblCellMar`, `w:tblLayout`, `w:tblLook` and `w:tblpPr`
 * were read and dropped whenever the reader took no typed value from them. A
 * save that rewrote the element — which is every save after an edit — lost
 * them.
 *
 * The universe here is the generated declared-child list rather than a list
 * somebody kept: `Record<DeclaredChild<"table-properties">, string>` is total,
 * so a child the schema gains cannot reach the property test without a sample,
 * and the same list gives the order the assertions check.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Table } from "../types/document";
import { CONTAINER_CHILDREN, type DeclaredChild } from "./containerChildren.gen";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DECLARED = CONTAINER_CHILDREN["table-properties"];

/**
 * One authored instance per declared child, each stating something.
 *
 * `satisfies Record<DeclaredChild<…>, string>` is what makes this total: a
 * child added to `CT_TblPr` fails the compile here before it can be dropped
 * silently at run time.
 */
const SAMPLES = {
  tblStyle: '<w:tblStyle w:val="TableGrid"/>',
  tblpPr: '<w:tblpPr w:leftFromText="180" w:vertAnchor="text" w:tblpX="100"/>',
  tblOverlap: '<w:tblOverlap w:val="never"/>',
  bidiVisual: "<w:bidiVisual/>",
  tblStyleRowBandSize: '<w:tblStyleRowBandSize w:val="2"/>',
  tblStyleColBandSize: '<w:tblStyleColBandSize w:val="3"/>',
  tblW: '<w:tblW w:w="5000" w:type="pct"/>',
  jc: '<w:jc w:val="center"/>',
  tblCellSpacing: '<w:tblCellSpacing w:w="15" w:type="dxa"/>',
  tblInd: '<w:tblInd w:w="120" w:type="dxa"/>',
  tblBorders:
    '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>',
  shd: '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>',
  tblLayout: '<w:tblLayout w:type="fixed"/>',
  tblCellMar: '<w:tblCellMar><w:top w:w="57" w:type="dxa"/></w:tblCellMar>',
  tblLook: '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"/>',
  tblCaption: '<w:tblCaption w:val="Quarterly totals"/>',
  tblDescription: '<w:tblDescription w:val="Totals per quarter"/>',
  tblPrChange:
    '<w:tblPrChange w:id="7" w:author="Reviewer" w:date="2026-05-15T12:00:00Z">' +
    '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange>',
} as const satisfies Record<DeclaredChild<"table-properties">, string>;

/**
 * Children the reader states nothing about, one per way of stating nothing.
 *
 * An element with no attributes at all, one whose value the reader's
 * enumeration does not admit, and one from a namespace the content model does
 * not name. None of the three can be decided by a map keyed on the child's
 * name, which is why the handler answers with what it took.
 */
const UNREAD_CHILDREN = [
  "<w:tblBorders/>",
  "<w:tblLook/>",
  '<w:jc w:val="end"/>',
  '<w:tblLayout w:type="wobble"/>',
  '<x:hint xmlns:x="urn:example:vendor" x:kind="layout"/>',
] as const;

const ROW = "<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>";

const tableXml = (properties: string): string =>
  `<w:tbl xmlns:w="${W}"><w:tblPr>${properties}</w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>${ROW}</w:tbl>`;

/**
 * Parse and save with every replayable capture cleared.
 *
 * `TableFormatting.sourceXml` replays the authored element whenever the model
 * still agrees with it, so a round trip that keeps it exercises the capture
 * machinery instead of the serializer — the same forcing the survival law
 * applies, and the only way this test can see a serializer defect at all.
 */
const rebuild = (properties: string): string => {
  const root = parseXmlDocument(tableXml(properties)) as XmlElement | null;
  if (!root) {
    throw new Error("the table fixture did not parse");
  }
  const table = parseTable(root, null, null, null, null, null);
  if (!table) {
    throw new Error("the table fixture parsed to nothing");
  }
  const { sourceXml: _source, gridSourceXml: _grid, ...formatting } = table.formatting ?? {};
  return serializeTable({ ...table, formatting }, serializeParagraph);
};

/** The same table after a no-op pass through the editor's document model. */
const throughEditor = (properties: string): string => {
  const root = parseXmlDocument(tableXml(properties)) as XmlElement | null;
  if (!root) {
    throw new Error("the table fixture did not parse");
  }
  const table = parseTable(root, null, null, null, null, null);
  if (!table) {
    throw new Error("the table fixture parsed to nothing");
  }
  const document = {
    package: { document: { content: [table], finalSectionProperties: {} } },
  } as never;
  const projected = fromProseDoc(toProseDoc(document), document).package.document
    .content[0] as Table;
  const { sourceXml: _source, gridSourceXml: _grid, ...formatting } = projected.formatting ?? {};
  return serializeTable({ ...projected, formatting }, serializeParagraph);
};

/** The outer `w:tblPr`'s own children: a `w:tblPrChange` nests a second one. */
const propertiesOf = (saved: string): string =>
  saved.slice(saved.indexOf("<w:tblPr>") + "<w:tblPr>".length, saved.lastIndexOf("</w:tblPr>"));

describe("a table's property set survives a rebuild", () => {
  test("every declared child comes back, and the save after it is a fixed point", () => {
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        const saved = rebuild(SAMPLES[name]);

        expect(saved).toContain(`<w:${name}`);
        expect(rebuild(propertiesOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a child the reader takes nothing from keeps its bytes", () => {
    fc.assert(
      fc.property(fc.constantFrom(...UNREAD_CHILDREN), (child) => {
        const saved = rebuild(`<w:tblStyle w:val="TableGrid"/>${child}`);

        expect(saved).toContain(child);
        expect(rebuild(propertiesOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the children come back in the order the content model declares", () => {
    // Authored backwards: a serializer that wrote them in the order it read
    // them, or in the order of its own statements, would come back out of
    // sequence and a validating consumer would refuse the part.
    const saved = rebuild(
      [...DECLARED]
        .reverse()
        .map((name) => SAMPLES[name])
        .join(""),
    );

    const positions = DECLARED.map((name) => saved.indexOf(`<w:${name}`));
    expect(positions).toEqual([...positions].sort((left, right) => left - right));
    expect(positions.every((at) => at > -1)).toBe(true);
  });

  test("a capture comes back between the same two modelled properties", () => {
    const saved = rebuild('<w:tblStyle w:val="TableGrid"/><w:tblLook/><w:tblCaption w:val="c"/>');

    expect(saved.indexOf("<w:tblStyle")).toBeLessThan(saved.indexOf("<w:tblLook/>"));
    expect(saved.indexOf("<w:tblLook/>")).toBeLessThan(saved.indexOf("<w:tblCaption"));
  });

  test("a capture before the first declared property stays first", () => {
    const saved = rebuild(
      '<x:hint xmlns:x="urn:example:vendor" x:kind="layout"/>' + '<w:tblStyle w:val="TableGrid"/>',
    );

    expect(saved.indexOf("<x:hint ")).toBeLessThan(saved.indexOf("<w:tblStyle "));
  });

  test("every declared child survives the editor projection", () => {
    // `TableAttrs._originalFormatting` carries the whole record through
    // ProseMirror, so the sink and the newly modelled properties ride it; the
    // assertion is that the way back does not rebuild the element from the
    // handful of attrs the editor surfaces.
    fc.assert(
      // The comparison is the property set alone: the editor gives every cell
      // an explicit width, which is a decision about cells and not about this.
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        expect(propertiesOf(throughEditor(SAMPLES[name]))).toBe(
          propertiesOf(rebuild(SAMPLES[name])),
        );
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a property change survives a table that states no properties of its own", () => {
    const saved = rebuild(SAMPLES.tblPrChange);

    expect(saved).toContain('<w:tblPrChange w:id="7" w:author="Reviewer"');
    expect(saved).toContain('<w:tblW w:w="0" w:type="auto"/>');
  });
});
