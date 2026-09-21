/**
 * A cell's property set survives a rebuild, in declaration order.
 *
 * `w:tcPr` was read by name, so `w:cnfStyle`, `w:hMerge`, `w:headers` and the
 * structural revision a `w:tcPrChange` snapshot carries went on every save.
 * The set is now dispatched over the generated declared-child list, and
 * `Record<DeclaredChild<"cell-properties">, string>` is total, so a child the
 * schema gains cannot reach this test without a sample.
 *
 * The order assertion carries more weight here than for the row.
 * `CT_TcPrBase` is a sequence and every type extending it is one, so a
 * `w:tcPr` written in another order is markup Word refuses — and the corpus
 * validator cannot say so, because the chain reaches
 * `EG_CellMarkupElements` and it declines to order a model that can reorder
 * itself. This file is the oracle `container-children-order.test.ts` cannot
 * be for this row.
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

const DECLARED = CONTAINER_CHILDREN["cell-properties"];

/** One authored instance per declared child, each stating something. */
const SAMPLES = {
  cnfStyle: '<w:cnfStyle w:val="000100000000"/>',
  tcW: '<w:tcW w:w="2400" w:type="dxa"/>',
  gridSpan: '<w:gridSpan w:val="2"/>',
  hMerge: '<w:hMerge w:val="restart"/>',
  vMerge: '<w:vMerge w:val="restart"/>',
  tcBorders:
    '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>',
  shd: '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>',
  noWrap: "<w:noWrap/>",
  tcMar: '<w:tcMar><w:top w:w="57" w:type="dxa"/></w:tcMar>',
  textDirection: '<w:textDirection w:val="tbRl"/>',
  tcFitText: "<w:tcFitText/>",
  vAlign: '<w:vAlign w:val="center"/>',
  hideMark: "<w:hideMark/>",
  headers: '<w:headers><w:header w:val="TopLeft"/></w:headers>',
  cellIns: '<w:cellIns w:id="11" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  cellDel: '<w:cellDel w:id="12" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  cellMerge:
    '<w:cellMerge w:id="13" w:author="Reviewer" w:date="2026-05-15T12:00:00Z" w:vMerge="cont"/>',
  tcPrChange:
    '<w:tcPrChange w:id="14" w:author="Reviewer" w:date="2026-05-15T12:00:00Z">' +
    '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr></w:tcPrChange>',
} as const satisfies Record<DeclaredChild<"cell-properties">, string>;

/**
 * `EG_CellMarkupElements` is a choice, so the three structural revisions are
 * one slot in the model and the parser takes the insertion. The order
 * assertion authors the insertion and leaves the other two out, rather than
 * asking the model to hold a revision it has no place for.
 */
const STRUCTURAL_ALTERNATIVES: ReadonlySet<string> = new Set(["cellDel", "cellMerge"]);
const ORDERABLE = DECLARED.filter((name) => !STRUCTURAL_ALTERNATIVES.has(name));

/**
 * Children the reader states nothing about, one per way of stating nothing.
 *
 * An element with no attributes at all, one whose value the reader's
 * enumeration does not admit, one whose value it normalises away, and one from
 * a namespace the content model does not name.
 */
const UNREAD_CHILDREN = [
  "<w:cnfStyle/>",
  '<w:vAlign w:val="both"/>',
  '<w:textDirection w:val="wobble"/>',
  '<w:gridSpan w:val="1"/>',
  '<x:hint xmlns:x="urn:example:vendor" x:kind="cell"/>',
] as const;

const tableXml = (cellProperties: string): string =>
  `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr>${cellProperties}</w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p>` +
  "</w:tc></w:tr></w:tbl>";

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

/** The cell's own capture cleared, so the serializer runs; see the row's file. */
const withoutReplay = (table: Table): Table => ({
  ...table,
  rows: table.rows.map((row) => ({
    ...row,
    cells: row.cells.map((cell) => {
      if (!cell.formatting) {
        return cell;
      }
      const { sourceXml: _source, ...formatting } = cell.formatting;
      return { ...cell, formatting };
    }),
  })),
});

const rebuild = (cellProperties: string): string =>
  serializeTable(withoutReplay(parsed(tableXml(cellProperties))), serializeParagraph);

/** The same table after a no-op pass through the editor's document model. */
const throughEditor = (cellProperties: string): string => {
  const table = parsed(tableXml(cellProperties));
  const document = {
    package: { document: { content: [table], finalSectionProperties: {} } },
  } as never;
  const projected = fromProseDoc(toProseDoc(document), document).package.document
    .content[0] as Table;
  return serializeTable(withoutReplay(projected), serializeParagraph);
};

/** The outer `w:tcPr`'s own children: a `w:tcPrChange` nests a second one. */
const cellPropertiesOf = (saved: string): string =>
  saved.slice(saved.indexOf("<w:tcPr>") + "<w:tcPr>".length, saved.lastIndexOf("</w:tcPr>"));

describe("a cell's property set survives a rebuild", () => {
  test("every declared child comes back, and the save after it is a fixed point", () => {
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        const saved = rebuild(SAMPLES[name]);

        expect(saved).toContain(`<w:${name}`);
        expect(rebuild(cellPropertiesOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a child the reader takes nothing from keeps its bytes", () => {
    fc.assert(
      fc.property(fc.constantFrom(...UNREAD_CHILDREN), (child) => {
        const saved = rebuild(`<w:tcW w:w="2400" w:type="dxa"/>${child}`);

        expect(saved).toContain(child);
        expect(rebuild(cellPropertiesOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the children come back in the order the content model declares", () => {
    // Authored backwards. `CT_TcPrBase` is a sequence, so this is the
    // assertion that folio does not hand Word a cell it refuses to open.
    const saved = rebuild(
      [...ORDERABLE]
        .reverse()
        .map((name) => SAMPLES[name])
        .join(""),
    );

    const positions = ORDERABLE.map((name) => saved.indexOf(`<w:${name}`));
    expect(positions.every((at) => at > -1)).toBe(true);
    expect(positions).toEqual([...positions].sort((left, right) => left - right));
  });

  test("a capture comes back between the same two modelled properties", () => {
    const saved = rebuild('<w:tcW w:w="2400" w:type="dxa"/><w:hMerge/><w:shd w:val="clear"/>');

    expect(saved.indexOf("<w:tcW")).toBeLessThan(saved.indexOf("<w:hMerge/>"));
    expect(saved.indexOf("<w:hMerge/>")).toBeLessThan(saved.indexOf("<w:shd "));
  });

  test("every declared child survives the editor projection", () => {
    // `TableCellAttrs._originalFormatting` carries the whole record through
    // ProseMirror, so the sink rides it.
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        expect(cellPropertiesOf(throughEditor(SAMPLES[name]))).toBe(
          cellPropertiesOf(rebuild(SAMPLES[name])),
        );
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("an empty element survives, because its presence is the value", () => {
    // `w:tcPr` is optional on `CT_Tc`, so a cell that wrote an empty one said
    // something an absent element does not.
    expect(rebuild("")).toContain("<w:tcPr/>");
    expect(throughEditor("")).toContain("<w:tcPr/>");
  });

  test("a cell that never wrote one still writes none", () => {
    const table = parsed(
      `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
        "<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>",
    );

    expect(serializeTable(withoutReplay(table), serializeParagraph)).not.toContain("<w:tcPr");
  });

  test("a snapshot's own structural revision rides the change that recorded it", () => {
    // `CT_TcPrInner` declares `EG_CellMarkupElements`, so a `w:tcPrChange` may
    // state "before this change the cell stood inserted". It is not the cell's
    // current revision, and a walk that captured it would write it twice.
    const saved = rebuild(
      '<w:tcPrChange w:id="15" w:author="Reviewer" w:date="2026-05-15T12:00:00Z">' +
        '<w:tcPr><w:cellIns w:id="16" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>' +
        "</w:tcPr></w:tcPrChange>",
    );

    expect(saved).toContain('<w:cellIns w:id="16"');
    expect(saved.match(/<w:cellIns /gu)).toHaveLength(1);
    expect(rebuild(cellPropertiesOf(saved))).toBe(saved);
  });
});
