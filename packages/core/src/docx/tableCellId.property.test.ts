/** A table cell's authored identifier survives every editor projection. */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Table } from "../types/document";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable, serializeTableCell } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const cellId = fc
  .array(fc.constantFrom(..."ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-"), {
    minLength: 1,
    maxLength: 32,
  })
  .map((characters) => characters.join(""));

const tableHolding = (id: string): Table => {
  const root = parseXmlDocument(
    `<w:tbl xmlns:w="${WORD_NAMESPACE}"><w:tblGrid><w:gridCol/></w:tblGrid>` +
      `<w:tr><w:tc w:id="${id}"><w:p/></w:tc></w:tr></w:tbl>`,
  );
  if (!root) {
    throw new Error("The table-cell identity fixture XML did not parse.");
  }
  const table = parseTable(root, null, null, null, null, null);
  if (!table) {
    throw new Error("The table-cell identity fixture did not parse.");
  }
  return table;
};

const documentHolding = (table: Table): Document => ({
  package: { document: { content: [table] } },
});

describe("table-cell identifiers", () => {
  test(
    "every authored identifier survives the editor projection and save",
    () => {
      fc.assert(
        fc.property(cellId, (id) => {
          const source = documentHolding(tableHolding(id));
          const prose = toProseDoc(source);
          const cloned = prose.type.schema.nodeFromJSON(prose.toJSON());
          const projectedDocument = fromProseDoc(cloned, source, { reuse: "none" });
          const projected = projectedDocument.package.document.content.at(0);
          if (projected?.type !== "table") {
            throw new Error("The editor projection lost the table.");
          }

          expect(projected.rows.at(0)?.cells.at(0)?.id).toBe(id);
          expect(serializeTable(projected, serializeParagraph)).toContain(`<w:tc w:id="${id}">`);
        }),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(30_000),
  );

  test("escapes an authored identifier at the XML boundary", () => {
    expect(
      serializeTableCell(
        {
          type: "tableCell",
          id: 'cell&"id',
          content: [{ type: "paragraph", content: [] }],
        },
        serializeParagraph,
      ),
    ).toContain('w:id="cell&amp;&quot;id"');
  });
});
