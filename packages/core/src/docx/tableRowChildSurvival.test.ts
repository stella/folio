/**
 * A row keeps the children folio does not model, between the same two cells.
 *
 * `CT_Row` declares a permission range, a proofing error, the row-level
 * comment and move ranges and the eight custom-XML revision ranges beside its
 * cells. `parseRowChild` read `w:tc`, unwrapped `w:sdt` and carried a bookmark
 * boundary into a neighbouring cell; everything else it returned from.
 *
 * A row-level child cannot be a cell, so the capture is the row's sink rather
 * than a member of a union, and `index` is the count of cells that preceded
 * it. The assertions are about that position: markup written back at the end
 * of the row is markup that has left the column it was authored in.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseTable } from "./tableParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Children `CT_Row` declares and folio models nothing of.
 *
 * One per shape: an empty marker, a marker carrying revision attributes, a
 * wrapper with a whole cell inside it, and an element from a namespace the
 * content model does not name, which is the sink rather than the handler map.
 */
const UNMODELLED_CHILDREN = [
  '<w:permStart w:id="7" w:edGrp="everyone"/>',
  '<w:proofErr w:type="spellStart"/>',
  '<w:customXmlInsRangeStart w:id="3" w:author="Reviewer"/>',
  '<w:ins w:id="4" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:tc><w:p/></w:tc></w:ins>',
  '<x:note xmlns:x="urn:example:vendor" x:kind="aside">kept</x:note>',
] as const;

const CELL = "<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>";

const tableXml = (rowChildren: string): string =>
  `<w:tbl xmlns:w="${W}"><w:tr>${rowChildren}</w:tr></w:tbl>`;

const roundTrip = (rowChildren: string): string => {
  const root = parseXmlDocument(tableXml(rowChildren)) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse the table fixture");
  }
  return serializeTable(parseTable(root, null, null, null, null, null), serializeParagraph);
};

describe("a table row keeps the children folio does not model", () => {
  test("a row keeps a w:permStart between the two cells it stood between", () => {
    const saved = roundTrip(`${CELL}<w:permStart w:id="7"/>${CELL}`);

    expect(saved).toContain('<w:permStart w:id="7"/>');
    const firstCell = saved.indexOf("<w:tc>");
    const marker = saved.indexOf("<w:permStart");
    const lastCell = saved.lastIndexOf("<w:tc>");
    expect(firstCell).toBeLessThan(marker);
    expect(marker).toBeLessThan(lastCell);
  });

  test("every kind survives a save and the save after it", () => {
    fc.assert(
      fc.property(fc.constantFrom(...UNMODELLED_CHILDREN), (child) => {
        const saved = roundTrip(`${CELL}${child}${CELL}`);
        expect(saved).toContain(child);

        // Save, reopen, save: the second save is where a capture that only
        // replays and does not re-parse stops being a fixed point.
        expect(
          roundTrip(saved.slice(saved.indexOf("<w:tr>") + 6, saved.indexOf("</w:tr>"))),
        ).toContain(child);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a capture before every cell comes back before every cell", () => {
    const saved = roundTrip(`<w:permStart w:id="7"/>${CELL}`);
    expect(saved).toContain('<w:permStart w:id="7"/>');
    expect(saved.indexOf("<w:permStart")).toBeLessThan(saved.indexOf("<w:tc>"));
  });

  test("a row-level content control is still unwrapped, and its markup kept", () => {
    const saved = roundTrip(
      `<w:sdt><w:sdtPr/><w:sdtContent>${CELL}<w:proofErr w:type="spellEnd"/></w:sdtContent></w:sdt>`,
    );
    expect(saved).toContain("<w:tc>");
    expect(saved).toContain('<w:proofErr w:type="spellEnd"/>');
  });
});
