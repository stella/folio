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
import { createParseWarningCollector } from "./parseContext";
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

  test("empty customXml wrappers retain their cell position across saves", () => {
    const wrappers = [
      '<w:customXml w:element="empty"/>',
      '<w:customXml w:element="empty"></w:customXml>',
      '<w:customXml w:element="empty"><w:customXmlPr/></w:customXml>',
      '<w:customXml w:element="empty"><w:proofErr w:type="spellStart"/></w:customXml>',
    ] as const;
    fc.assert(
      fc.property(
        fc.constantFrom(...wrappers),
        fc.integer({ min: 0, max: 2 }),
        (wrapper, index) => {
          const expected =
            wrapper === '<w:customXml w:element="empty"></w:customXml>'
              ? '<w:customXml w:element="empty"/>'
              : wrapper;
          const children = [CELL, CELL];
          children.splice(index, 0, wrapper);
          const saved = roundTrip(children.join(""));
          const savedAgain = roundTrip(
            saved.slice(saved.indexOf("<w:tr>") + 6, saved.indexOf("</w:tr>")),
          );
          for (const output of [saved, savedAgain]) {
            expect(output).toContain(expected);
            expect(output.split('<w:customXml w:element="empty"').length - 1).toBe(1);
            const cellsBefore = output.slice(0, output.indexOf(expected)).match(/<w:tc>/gu) ?? [];
            expect(cellsBefore).toHaveLength(index);
          }
        },
      ),
      propertyConfig({ numRuns: 100, seed: 1486206489 }),
    );
  });

  test("a capture before every cell comes back before every cell", () => {
    const saved = roundTrip(`<w:permStart w:id="7"/>${CELL}`);
    expect(saved).toContain('<w:permStart w:id="7"/>');
    expect(saved.indexOf("<w:permStart")).toBeLessThan(saved.indexOf("<w:tc>"));
  });

  test("a row-level content control is still unwrapped, and its markup kept", () => {
    const input = `<w:sdt><w:sdtPr><w:alias w:val="row-control"/></w:sdtPr><w:sdtContent><w:tr>${CELL}</w:tr></w:sdtContent></w:sdt>`;
    const root = parseXmlDocument(`<w:tbl xmlns:w="${W}">${input}</w:tbl>`) as XmlElement;
    const parsed = parseTable(root, null, null, null, null, null);

    expect(parsed?.rows[0]?.cells).toHaveLength(1);
    expect(parsed?.rows[0]?.contentControls?.[0]?.alias).toBe("row-control");
    const saved = serializeTable(parsed!, serializeParagraph);
    expect(saved).toContain("<w:tc>");
    expect(saved).toContain("row-control");
  });

  test("a cell-level content control leaves its cell content readable", () => {
    const cellWithControl =
      '<w:sdt><w:sdtPr><w:alias w:val="cell-control"/></w:sdtPr><w:sdtContent>' +
      "<w:tc><w:p><w:r><w:t>visible sentinel</w:t></w:r></w:p></w:tc>" +
      "</w:sdtContent></w:sdt>";
    const root = parseXmlDocument(tableXml(cellWithControl)) as XmlElement;
    const parsed = parseTable(root, null, null, null, null, null);
    const cell = parsed?.rows[0]?.cells[0];

    expect(cell?.contentControls?.[0]?.alias).toBe("cell-control");
    expect(cell?.content[0]?.type).toBe("paragraph");
    expect(JSON.stringify(cell?.content)).toContain("visible sentinel");
    expect(roundTrip(cellWithControl)).toContain("visible sentinel");
  });

  test("row and table customXml wrappers leave their rows and cells readable", () => {
    const rowCustomXml = '<w:customXml w:uri="urn:row" w:element="row">' + `${CELL}</w:customXml>`;
    const tableCustomXml =
      '<w:customXml w:uri="urn:table" w:element="table">' + `<w:tr>${CELL}</w:tr></w:customXml>`;
    const rowRoot = parseXmlDocument(tableXml(rowCustomXml)) as XmlElement;
    const tableRoot = parseXmlDocument(
      `<w:tbl xmlns:w="${W}">${tableCustomXml}</w:tbl>`,
    ) as XmlElement;

    const rowTable = parseTable(rowRoot, null, null, null, null, null);
    const tableTable = parseTable(tableRoot, null, null, null, null, null);
    expect(rowTable?.rows[0]?.cells).toHaveLength(1);
    expect(tableTable?.rows[0]?.cells).toHaveLength(1);
    expect(serializeTable(rowTable!, serializeParagraph)).toContain('<w:customXml w:uri="urn:row"');
    expect(serializeTable(tableTable!, serializeParagraph)).toContain(
      '<w:customXml w:uri="urn:table"',
    );
  });

  test("a nested row reports that folio retained it as opaque markup", () => {
    const root = parseXmlDocument(
      `<w:tbl xmlns:w="${W}"><w:tr><w:tr><w:tc><w:p><w:r><w:t>hidden row</w:t></w:r></w:p></w:tc></w:tr></w:tr></w:tbl>`,
    ) as XmlElement;
    const collector = createParseWarningCollector("word/document.xml");

    parseTable(root, null, null, null, null, null, { context: collector.context });

    expect(collector.warnings()).toContainEqual({
      code: "nested-row-opaque",
      location: { part: "word/document.xml", element: "w:tr" },
      count: 1,
    });
  });

  test("customXmlPr and bookmark markers stay inside table and row wrappers", () => {
    const property = '<w:customXmlPr><w:attr w:name="key" w:val="value"/></w:customXmlPr>';
    const rowWrapped =
      '<w:customXml w:element="row">' +
      `${property}<w:bookmarkStart w:id="21" w:name="row-boundary"/>${CELL}` +
      '<w:bookmarkEnd w:id="21"/></w:customXml>';
    const tableWrapped =
      '<w:customXml w:element="table">' +
      `${property}<w:bookmarkStart w:id="22" w:name="table-boundary"/>` +
      `<w:tr>${CELL}</w:tr><w:bookmarkEnd w:id="22"/></w:customXml>`;
    const rowSaved = roundTrip(`<w:tr>${rowWrapped}</w:tr>`);
    const tableRoot = parseXmlDocument(
      `<w:tbl xmlns:w="${W}">${tableWrapped}</w:tbl>`,
    ) as XmlElement;
    const tableSaved = serializeTable(
      parseTable(tableRoot, null, null, null, null, null)!,
      serializeParagraph,
    );

    const rowWrapperContent = rowSaved
      .split('<w:customXml w:element="row">')[1]
      ?.split("</w:customXml>")[0];
    const tableWrapperContent = tableSaved
      .split('<w:customXml w:element="table">')[1]
      ?.split("</w:customXml>")[0];
    expect(rowWrapperContent).toContain("<w:customXmlPr>");
    expect(rowWrapperContent).toContain('w:name="row-boundary"');
    expect(rowWrapperContent).toContain("<w:tc>");
    expect(tableWrapperContent).toContain("<w:customXmlPr>");
    expect(tableWrapperContent).toContain('w:name="table-boundary"');
    expect(tableWrapperContent).toContain("<w:tr>");
  });

  test("adjacent identical customXml wrappers stay distinct", () => {
    const duplicateRowWrappers =
      `<w:tr><w:customXml w:element="same">${CELL}</w:customXml>` +
      `<w:customXml w:element="same">${CELL}</w:customXml></w:tr>`;
    const rowSaved = roundTrip(duplicateRowWrappers);
    const duplicateTableWrappers =
      `<w:customXml w:element="same"><w:tr>${CELL}</w:tr></w:customXml>` +
      `<w:customXml w:element="same"><w:tr>${CELL}</w:tr></w:customXml>`;
    const tableRoot = parseXmlDocument(
      `<w:tbl xmlns:w="${W}">${duplicateTableWrappers}</w:tbl>`,
    ) as XmlElement;
    const tableSaved = serializeTable(
      parseTable(tableRoot, null, null, null, null, null)!,
      serializeParagraph,
    );
    const wrapperCount = (xml: string): number =>
      (xml.match(/<w:customXml w:element="same">/gu) ?? []).length;

    expect(wrapperCount(rowSaved)).toBe(2);
    expect(wrapperCount(tableSaved)).toBe(2);
  });

  test("mixed SDT and customXml wrappers keep their authored nesting order", () => {
    const sdtOuter =
      `<w:sdt><w:sdtPr/><w:sdtContent><w:customXml w:element="inner">` +
      `<w:tr>${CELL}</w:tr></w:customXml></w:sdtContent></w:sdt>`;
    const customXmlOuter =
      `<w:customXml w:element="outer"><w:sdt><w:sdtPr/><w:sdtContent>` +
      `<w:tr>${CELL}</w:tr></w:sdtContent></w:sdt></w:customXml>`;
    const parseAndSave = (children: string): string => {
      const root = parseXmlDocument(`<w:tbl xmlns:w="${W}">${children}</w:tbl>`) as XmlElement;
      return serializeTable(parseTable(root, null, null, null, null, null)!, serializeParagraph);
    };
    const sdtOuterSaved = parseAndSave(sdtOuter);
    const customXmlOuterSaved = parseAndSave(customXmlOuter);

    const rowSdtOuter =
      `<w:sdt><w:sdtPr/><w:sdtContent><w:customXml w:element="inner">` +
      `${CELL}</w:customXml></w:sdtContent></w:sdt>`;
    const rowCustomXmlOuter =
      `<w:customXml w:element="outer"><w:sdt><w:sdtPr/><w:sdtContent>` +
      `${CELL}</w:sdtContent></w:sdt></w:customXml>`;
    const rowSdtOuterSaved = roundTrip(rowSdtOuter);
    const rowCustomXmlOuterSaved = roundTrip(rowCustomXmlOuter);

    expect(sdtOuterSaved.indexOf("<w:sdt>")).toBeLessThan(sdtOuterSaved.indexOf("<w:customXml"));
    expect(customXmlOuterSaved.indexOf("<w:customXml")).toBeLessThan(
      customXmlOuterSaved.indexOf("<w:sdt>"),
    );
    expect(rowSdtOuterSaved.indexOf("<w:sdt>")).toBeLessThan(
      rowSdtOuterSaved.indexOf("<w:customXml"),
    );
    expect(rowCustomXmlOuterSaved.indexOf("<w:customXml")).toBeLessThan(
      rowCustomXmlOuterSaved.indexOf("<w:sdt>"),
    );
  });
});
