/**
 * A rebuilt container must keep the tracked-change record it arrived with.
 *
 * `w:tblGridChange`, `w:numberingChange`, and `w:numPr/w:ins` are revisions:
 * they record the replaced grid, replaced numbering, and inserted numbering
 * properties. Nothing in the editable model derives them, so a serializer that
 * rebuilds their container from the model must carry their records explicitly.
 * In a document under review that is not a fidelity detail; it accepts or
 * discards somebody's edit without telling anyone.
 *
 * Both used to survive only by verbatim replay, which is exactly the path an
 * edited document does not take. The tests therefore force the rebuild:
 * resizing a column makes the captured `w:tblGrid` stop matching the model,
 * and the `w:numberingChange` case is checked both through replay and through
 * a rebuild, because refusing the capture used to be what forced a rebuild
 * that could not write it.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Table } from "../../types/document";
import { serializeTable } from "../serializer/tableSerializer";
import { parseTable } from "../tableParser";
import { parseXmlDocument, type XmlElement } from "../xmlParser";
import { parseParagraph } from "../paragraphParser";
import { serializeParagraph } from "../serializer/paragraphSerializer";

setDefaultTimeout(propertyTestTimeout(30_000));

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const GRID_CHANGE =
  '<w:tblGridChange w:id="42"><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="5000"/></w:tblGrid></w:tblGridChange>';

const TABLE_XML =
  `<w:tbl ${W_NS}><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/>${GRID_CHANGE}</w:tblGrid>` +
  "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>";

const NUMBERING_CHANGE =
  '<w:numberingChange w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z" w:original="1)."/>';
const NUMBERING_INSERTION = '<w:ins w:id="8" w:author="Editor" w:date="2024-01-02T00:00:00Z"/>';

const PARAGRAPH_XML =
  `<w:p ${W_NS}><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/>` +
  `${NUMBERING_CHANGE}${NUMBERING_INSERTION}</w:numPr></w:pPr>` +
  "<w:r><w:t>item</w:t></w:r></w:p>";

/**
 * The table with every verbatim capture cleared, which is what the survival
 * law's carrier probe does: whatever survives this, the typed model holds.
 */
const modelOnly = (table: Table | undefined): Table => {
  if (!table) {
    throw new Error("the fixture did not parse into a table");
  }
  const { sourceXml: _source, gridSourceXml: _grid, ...formatting } = table.formatting ?? {};
  return { ...table, formatting };
};

const parseElement = (xml: string): XmlElement => {
  const node = parseXmlDocument(xml) as XmlElement | null;
  if (!node) {
    throw new Error("the fixture did not parse");
  }
  return node;
};

describe("w:tblGridChange survives a grid the serializer has to rebuild", () => {
  test("a resized column keeps the record of the grid it replaced", () => {
    const table = parseTable(parseElement(TABLE_XML), new Map(), null, null, null, null);
    expect(table).toBeDefined();
    if (!table) {
      return;
    }
    expect(table.formatting?.gridChange).toEqual({ id: 42, columnWidths: [1000, 5000] });

    // Resizing a column is what makes the captured grid stop matching the
    // model, so the serializer rebuilds it instead of replaying the capture.
    table.columnWidths = [3000, 3000];
    const xml = serializeTable(table, serializeParagraph);

    expect(xml).toContain('<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/>');
    expect(xml).toContain('<w:tblGridChange w:id="42">');
    expect(xml).toContain('<w:gridCol w:w="5000"/>');
  });

  test("an unchanged grid still replays, record and all", () => {
    const table = parseTable(parseElement(TABLE_XML), new Map(), null, null, null, null);
    expect(table).toBeDefined();
    if (!table) {
      return;
    }
    expect(serializeTable(table, serializeParagraph)).toContain(GRID_CHANGE);
  });

  // A `w:tblGridChange` holds a `w:tblGrid` and that grid holds its own
  // `w:gridCol` children: a container nested in one of its own kind. While the
  // snapshot travelled as bytes, a save with nothing verbatim left wrote the
  // outer grid and nothing under the change, which is what the survival
  // census's carrier probe was reading when it called the pair captured.
  test("the snapshot's own grid and columns survive a save the model alone builds", () => {
    fc.assert(
      fc.property(
        fc.array(fc.option(fc.integer({ min: 1, max: 20_000 }), { nil: undefined }), {
          minLength: 1,
          maxLength: 6,
        }),
        fc.integer({ min: 0, max: 9999 }),
        (snapshotWidths, changeId) => {
          const snapshot = snapshotWidths
            .map((width) => (width === undefined ? "<w:gridCol/>" : `<w:gridCol w:w="${width}"/>`))
            .join("");
          const xml =
            `<w:tbl ${W_NS}><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/>` +
            `<w:tblGridChange w:id="${changeId}"><w:tblGrid>${snapshot}</w:tblGrid></w:tblGridChange>` +
            "</w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>";

          const table = modelOnly(parseTable(parseElement(xml), new Map(), null, null, null, null));
          expect(table.formatting?.gridChange).toEqual({
            id: changeId,
            columnWidths: snapshotWidths,
          });

          const saved = serializeTable(table, serializeParagraph);
          expect(saved).toContain(
            `<w:tblGridChange w:id="${changeId}"><w:tblGrid>${snapshot}</w:tblGrid></w:tblGridChange>`,
          );

          // A fixed point: the second parse reads the same snapshot the first
          // did, so nothing about it depends on the bytes it arrived as. The
          // serializer returns a fragment, and the revision id is resolved
          // against the namespace the prefix is bound to, so the fragment is
          // read back declaring the binding the saved part carries.
          const reparsed = parseTable(
            parseElement(saved.replace("<w:tbl>", `<w:tbl ${W_NS}>`)),
            new Map(),
            null,
            null,
            null,
            null,
          );
          expect(reparsed?.formatting?.gridChange).toEqual(table.formatting?.gridChange);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });
});

describe("numbering revision records survive a w:pPr the serializer has to rebuild", () => {
  test("both records reach the model", () => {
    const paragraph = parseParagraph(parseElement(PARAGRAPH_XML), new Map(), null, null);
    expect(paragraph.formatting?.numberingChangeXml).toContain('w:original="1)."');
    expect(paragraph.formatting?.numberingInsertionXml).toContain('w:author="Editor"');
  });

  test("a rebuilt w:numPr writes both records back in schema order", () => {
    const paragraph = parseParagraph(parseElement(PARAGRAPH_XML), new Map(), null, null);
    // Changing the numbering is what a reviewer's next edit does, and it is
    // what stops the captured `w:pPr` from being replayed.
    if (paragraph.formatting?.numPr) {
      paragraph.formatting.numPr = { kind: "reference", numId: 9 };
    }
    const xml = serializeParagraph(paragraph);
    expect(xml).toContain('<w:numId w:val="9"/>');
    expect(xml).toContain('w:author="Reviewer"');
    expect(xml).toContain('w:original="1)."');
    expect(xml).toContain(`${NUMBERING_CHANGE}${NUMBERING_INSERTION}`);
  });

  test("an insertion record keeps an otherwise empty w:numPr", () => {
    const xml =
      `<w:p ${W_NS}><w:pPr><w:numPr>${NUMBERING_INSERTION}</w:numPr></w:pPr>` +
      "<w:r><w:t>item</w:t></w:r></w:p>";
    const paragraph = parseParagraph(parseElement(xml), new Map(), null, null);

    expect(paragraph.formatting?.numPr).toBeUndefined();
    expect(serializeParagraph(paragraph)).toContain(`<w:numPr>${NUMBERING_INSERTION}</w:numPr>`);
  });
});
