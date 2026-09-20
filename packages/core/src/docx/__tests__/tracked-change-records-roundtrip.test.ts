/**
 * A rebuilt container must keep the tracked-change record it arrived with.
 *
 * `w:tblGridChange` and `w:numberingChange` are revisions: one records the
 * grid a reviewer replaced when resizing a column, the other the numbering a
 * reviewer replaced when changing a list. Nothing in the editable model
 * derives either, so a serializer that rebuilds the container from the model
 * drops the revision, and the document then says the change was always there.
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

import { describe, expect, test } from "bun:test";

import { serializeTable } from "../serializer/tableSerializer";
import { parseTable } from "../tableParser";
import { parseXmlDocument, type XmlElement } from "../xmlParser";
import { parseParagraph } from "../paragraphParser";
import { serializeParagraph } from "../serializer/paragraphSerializer";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const GRID_CHANGE =
  '<w:tblGridChange w:id="42"><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="5000"/></w:tblGrid></w:tblGridChange>';

const TABLE_XML =
  `<w:tbl ${W_NS}><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="3600"/>${GRID_CHANGE}</w:tblGrid>` +
  "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>";

const NUMBERING_CHANGE =
  '<w:numberingChange w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z" w:original="1)."/>';

const PARAGRAPH_XML =
  `<w:p ${W_NS}><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/>${NUMBERING_CHANGE}</w:numPr></w:pPr>` +
  "<w:r><w:t>item</w:t></w:r></w:p>";

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
    expect(table.formatting?.gridChangeXml).toContain("w:tblGridChange");

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
});

describe("w:numberingChange survives a w:pPr the serializer has to rebuild", () => {
  test("the record reaches the model", () => {
    const paragraph = parseParagraph(parseElement(PARAGRAPH_XML), new Map(), null, null);
    expect(paragraph.formatting?.numberingChangeXml).toContain('w:original="1)."');
  });

  test("a rebuilt w:numPr writes the record back", () => {
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
  });
});
