/**
 * A table keeps the children folio does not model, between the same two rows.
 *
 * `CT_Tbl` declares a permission range, a proofing error, the table-level
 * comment and move ranges and the eight custom-XML revision ranges beside its
 * rows. `parseTableChild` read `w:tr`, unwrapped `w:sdt` and returned from
 * everything else.
 *
 * A table-level child cannot be a row, so the capture is the table's sink
 * rather than a member of a union, and `index` is the count of rows that
 * preceded it. The assertions are about that position: markup written back at
 * the end of the table is markup that has left the row it was authored above.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import { parseTable } from "./tableParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Children `CT_Tbl` declares and folio models nothing of.
 *
 * One per shape: an empty marker, a marker carrying revision attributes, a
 * custom-XML revision range, a transparent wrapper holding a whole row — kept
 * whole rather than unwrapped — and an element from a namespace the content
 * model does not name, which is the sink rather than the handler map.
 */
const UNMODELLED_CHILDREN = [
  '<w:permStart w:id="7" w:edGrp="everyone"/>',
  '<w:proofErr w:type="spellStart"/>',
  '<w:customXmlInsRangeStart w:id="3" w:author="Reviewer"/>',
  '<w:customXml w:element="aside"><w:tr><w:tc><w:p/></w:tc></w:tr></w:customXml>',
  '<x:note xmlns:x="urn:example:vendor" x:kind="aside">kept</x:note>',
] as const;

const ROW = "<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>";

const tableXml = (tableChildren: string): string =>
  `<w:tbl xmlns:w="${W}">${tableChildren}</w:tbl>`;

const roundTrip = (tableChildren: string): string => {
  const root = parseXmlDocument(tableXml(tableChildren)) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse the table fixture");
  }
  const table = parseTable(root, null, null, null, null, null);
  if (!table) {
    throw new Error("The table fixture parsed to nothing");
  }
  return serializeTable(table, serializeParagraph);
};

/** The saved table's own children, so a reopen sees what the first save wrote. */
const savedChildren = (saved: string): string =>
  saved.slice("<w:tbl>".length, saved.lastIndexOf("</w:tbl>"));

describe("a table keeps the children folio does not model", () => {
  test("a table keeps a w:permStart between the two rows it stood between", () => {
    const saved = roundTrip(`${ROW}<w:permStart w:id="7"/>${ROW}`);

    expect(saved).toContain('<w:permStart w:id="7"/>');
    const firstRow = saved.indexOf("<w:tr>");
    const marker = saved.indexOf("<w:permStart");
    const lastRow = saved.lastIndexOf("<w:tr>");
    expect(firstRow).toBeLessThan(marker);
    expect(marker).toBeLessThan(lastRow);
  });

  test("every kind survives a save and the save after it", () => {
    fc.assert(
      fc.property(fc.constantFrom(...UNMODELLED_CHILDREN), (child) => {
        const saved = roundTrip(`${ROW}${child}${ROW}`);
        expect(saved).toContain(child);

        // Save, reopen, save: the second save is where a capture that only
        // replays and does not re-parse stops being a fixed point.
        expect(roundTrip(savedChildren(saved))).toContain(child);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a capture before every row comes back before every row", () => {
    const saved = roundTrip(`<w:permStart w:id="7"/>${ROW}`);
    expect(saved).toContain('<w:permStart w:id="7"/>');
    expect(saved.indexOf("<w:permStart")).toBeLessThan(saved.indexOf("<w:tr>"));
  });

  test("the sink's index counts rows, not the properties that precede them", () => {
    const saved = roundTrip(
      `<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
        `<w:proofErr w:type="spellStart"/>${ROW}`,
    );

    // `w:tblPr` and `w:tblGrid` are written from the model, ahead of the sink;
    // an index that counted them would push the capture past the row.
    expect(saved.indexOf("<w:tblGrid>")).toBeLessThan(saved.indexOf("<w:proofErr"));
    expect(saved.indexOf("<w:proofErr")).toBeLessThan(saved.indexOf("<w:tr>"));
  });

  test("a row-level content control is still unwrapped, and its markup kept", () => {
    const saved = roundTrip(
      `<w:sdt><w:sdtPr/><w:sdtContent>${ROW}<w:proofErr w:type="spellEnd"/></w:sdtContent></w:sdt>`,
    );
    expect(saved).toContain("<w:tr>");
    expect(saved).toContain('<w:proofErr w:type="spellEnd"/>');
  });

  test("a table nested directly in a table is kept whole, not flattened into it", () => {
    const nested = `<w:tbl><w:tblPr/><w:tblGrid/>${ROW}</w:tbl>`;
    const saved = roundTrip(`${ROW}${nested}`);

    expect(saved).toContain(nested);
    // Flattening would have spliced the inner row into the outer table.
    expect(savedChildren(saved).indexOf("<w:tbl>")).toBeGreaterThan(-1);
  });
});
