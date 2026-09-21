/**
 * A row's table property exceptions survive a rebuild, in declaration order.
 *
 * `w:tblPrEx` had no reader at all: the row's child walk called it
 * `OWNED_ELSEWHERE` and nothing owned it, so the nine table properties a row
 * may override, and the `w:tblPrExChange` recording that a reviewer changed
 * them, went on every save. Word writes the element whenever two tables are
 * merged, and a consumer reads it in place of the table's own properties for
 * that row, so the loss restyles the row rather than merely thinning the file.
 *
 * `CT_Row` declares `w:tblPrEx` *before* `w:trPr`, which is the other half of
 * the fix: a reader that took the element and a serializer that wrote it after
 * the row's own properties would produce markup a validating consumer refuses.
 *
 * The universe is the generated declared-child list rather than a list
 * somebody kept: `Record<DeclaredChild<"table-property-exceptions">, string>`
 * is total, so a child the schema gains cannot reach this test without a
 * sample, and the same list gives the order the assertions check.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Table } from "../types/document";
import { CONTAINER_CHILDREN, type DeclaredChild } from "./containerChildren.gen";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DECLARED = CONTAINER_CHILDREN["table-property-exceptions"];

/**
 * One authored instance per declared child, each stating something.
 *
 * `satisfies Record<DeclaredChild<…>, string>` is what makes this total: a
 * child added to `CT_TblPrEx` fails the compile here before it can be dropped
 * silently at run time.
 */
const SAMPLES = {
  tblW: '<w:tblW w:w="4000" w:type="pct"/>',
  jc: '<w:jc w:val="center"/>',
  tblCellSpacing: '<w:tblCellSpacing w:w="15" w:type="dxa"/>',
  tblInd: '<w:tblInd w:w="120" w:type="dxa"/>',
  tblBorders:
    '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>',
  shd: '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>',
  tblLayout: '<w:tblLayout w:type="fixed"/>',
  tblCellMar: '<w:tblCellMar><w:top w:w="57" w:type="dxa"/></w:tblCellMar>',
  tblLook: '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"/>',
  tblPrExChange:
    '<w:tblPrExChange w:id="9" w:author="Reviewer" w:date="2026-05-15T12:00:00Z">' +
    '<w:tblPrEx><w:tblW w:w="0" w:type="auto"/></w:tblPrEx></w:tblPrExChange>',
} as const satisfies Record<DeclaredChild<"table-property-exceptions">, string>;

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

const CELL = "<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>";

const tableXml = (exceptions: string, rowProperties = ""): string =>
  `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
  `<w:tr><w:tblPrEx>${exceptions}</w:tblPrEx>${rowProperties}${CELL}</w:tr></w:tbl>`;

const parsed = (xml: string): Table => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("the table fixture did not parse");
  }
  const table = parseTable(root, null, null, null, null, null);
  if (!table) {
    throw new Error("the table fixture parsed to nothing");
  }
  return table;
};

/**
 * The exceptions' own capture cleared, so the serializer runs.
 *
 * `TableFormatting.sourceXml` replays the authored element whenever the model
 * still agrees with it, so a round trip that keeps it exercises the capture
 * machinery instead of the serializer — the same forcing the survival law
 * applies, and the only way this test can see a serializer defect at all.
 */
const withoutReplay = (table: Table): Table => ({
  ...table,
  rows: table.rows.map((row) => {
    if (!row.tablePropertyExceptions) {
      return row;
    }
    const { sourceXml: _source, ...exceptions } = row.tablePropertyExceptions;
    return { ...row, tablePropertyExceptions: exceptions };
  }),
});

const rebuild = (exceptions: string, rowProperties = ""): string =>
  serializeTable(withoutReplay(parsed(tableXml(exceptions, rowProperties))), serializeParagraph);

/** The same table after a no-op pass through the editor's document model. */
const throughEditor = (exceptions: string): string => {
  const table = parsed(tableXml(exceptions));
  const document = {
    package: { document: { content: [table], finalSectionProperties: {} } },
  } as never;
  const projected = fromProseDoc(toProseDoc(document), document).package.document
    .content[0] as Table;
  return serializeTable(withoutReplay(projected), serializeParagraph);
};

/** The outer `w:tblPrEx`'s own children: a `w:tblPrExChange` nests a second one. */
const exceptionsOf = (saved: string): string =>
  saved.slice(
    saved.indexOf("<w:tblPrEx>") + "<w:tblPrEx>".length,
    saved.lastIndexOf("</w:tblPrEx>"),
  );

describe("a row's table property exceptions survive a rebuild", () => {
  test("every declared child comes back, and the save after it is a fixed point", () => {
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        const saved = rebuild(SAMPLES[name]);

        expect(saved).toContain(`<w:${name}`);
        expect(rebuild(exceptionsOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a child the reader takes nothing from keeps its bytes", () => {
    fc.assert(
      fc.property(fc.constantFrom(...UNREAD_CHILDREN), (child) => {
        const saved = rebuild(`<w:tblW w:w="4000" w:type="pct"/>${child}`);

        expect(saved).toContain(child);
        expect(rebuild(exceptionsOf(saved))).toBe(saved);
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
    const saved = rebuild(
      '<w:tblW w:w="4000" w:type="pct"/><w:tblCellMar/><w:tblLook w:val="04A0"/>',
    );

    expect(saved.indexOf("<w:tblW")).toBeLessThan(saved.indexOf("<w:tblCellMar/>"));
    expect(saved.indexOf("<w:tblCellMar/>")).toBeLessThan(saved.indexOf("<w:tblLook "));
  });

  test("the exceptions precede the row's own properties", () => {
    // `CT_Row` is `w:tblPrEx, w:trPr, (cells)*`, and the row's sink counts
    // cells, so a serializer that wrote the exceptions after `w:trPr` — or
    // among the cells — would produce a row nothing opens.
    const saved = rebuild('<w:jc w:val="center"/>', "<w:trPr><w:tblHeader/></w:trPr>");

    expect(saved.indexOf("<w:tblPrEx>")).toBeLessThan(saved.indexOf("<w:trPr>"));
    expect(saved.indexOf("<w:trPr>")).toBeLessThan(saved.indexOf("<w:tc>"));
  });

  test("every declared child survives the editor projection", () => {
    // `TableRowAttrs._tablePropertyExceptions` carries the whole record through
    // ProseMirror, so the sink and the nine modelled properties ride it; the
    // assertion is that the way back does not rebuild the element from the
    // handful of attrs the editor surfaces.
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        expect(exceptionsOf(throughEditor(SAMPLES[name]))).toBe(
          exceptionsOf(rebuild(SAMPLES[name])),
        );
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("an empty element survives, because its presence is the value", () => {
    // `w:tblPrEx` is optional on `CT_Row`, so a row that wrote an empty one
    // said something an absent element does not. Reading it as "no exceptions"
    // deleted it on save.
    expect(rebuild("")).toContain("<w:tblPrEx/>");
  });

  test("a property change survives a row that states no exceptions of its own", () => {
    const saved = rebuild(SAMPLES.tblPrExChange);

    expect(saved).toContain('<w:tblPrExChange w:id="9" w:author="Reviewer"');
    expect(saved).toContain('<w:tblW w:w="0" w:type="auto"/>');
  });
});
