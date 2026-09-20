import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { fixTables, TableMap } from "prosemirror-tables";

import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import type { Document, Table } from "../../types/document";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const paragraph = (text: string) => ({
  type: "paragraph" as const,
  content: [{ type: "run" as const, content: [{ type: "text" as const, text }] }],
});

const documentWithOmittedGridSlots = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          columnWidths: [2400, 2400, 2400],
          rows: [
            {
              type: "tableRow",
              cells: [
                { type: "tableCell", content: [paragraph("A")] },
                { type: "tableCell", content: [paragraph("B")] },
                { type: "tableCell", content: [paragraph("C")] },
              ],
            },
            {
              type: "tableRow",
              formatting: { gridBefore: 1, gridAfter: 1 },
              cells: [{ type: "tableCell", content: [paragraph("middle")] }],
            },
          ],
        },
      ],
    },
  },
});

describe("toProseDoc — omitted table grid slots", () => {
  test("keeps a rectangular PM table without creating authored cells", () => {
    const source = documentWithOmittedGridSlots();
    const pmDoc = toProseDoc(source);
    const table = pmDoc.firstChild;
    if (!table || table.type.name !== "table") {
      throw new Error("Expected table");
    }
    const partialRow = table.child(1);

    expect(partialRow.childCount).toBe(3);
    expect(partialRow.child(0).attrs["_omittedGridSlot"]).toBe("before");
    expect(partialRow.child(1).textContent).toBe("middle");
    expect(partialRow.child(2).attrs["_omittedGridSlot"]).toBe("after");
    expect(TableMap.get(table).width).toBe(3);
    expect(fixTables(EditorState.create({ doc: pmDoc }))).toBeUndefined();
  });

  test("omits PM placeholders from layout and DOCX projection", () => {
    const source = documentWithOmittedGridSlots();
    const pmDoc = toProseDoc(source);
    const flowTable = toFlowBlocks(pmDoc).at(0);
    if (!flowTable || flowTable.kind !== "table") {
      throw new Error("Expected flow table");
    }

    expect(flowTable.rows[1]?.cells).toHaveLength(1);
    expect(flowTable.rows[1]).toMatchObject({ gridBefore: 1, gridAfter: 1 });

    const restored = fromProseDoc(pmDoc, source).package.document.content.at(0) as Table;
    expect(restored.rows[1]).toMatchObject({
      formatting: { gridBefore: 1, gridAfter: 1 },
      cells: [{ content: [paragraph("middle")] }],
    });
  });

  test("lets before and after omissions satisfy an otherwise empty row", () => {
    const source = documentWithOmittedGridSlots();
    const table = source.package.document.content.at(0) as Table;
    table.rows[1] = {
      type: "tableRow",
      formatting: { gridBefore: 1, gridAfter: 2 },
      cells: [],
    };

    const pmDoc = toProseDoc(source);
    const pmTable = pmDoc.firstChild;
    if (!pmTable || pmTable.type.name !== "table") {
      throw new Error("Expected table");
    }

    expect(pmTable.child(1).childCount).toBe(2);
    expect(TableMap.get(pmTable).width).toBe(3);
    expect(fixTables(EditorState.create({ doc: pmDoc }))).toBeUndefined();
    const restored = fromProseDoc(pmDoc, source).package.document.content.at(0) as Table;
    expect(restored.rows[1]?.cells).toEqual([]);
  });
});
