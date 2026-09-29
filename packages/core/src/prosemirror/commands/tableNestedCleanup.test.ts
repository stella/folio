import { expect, test } from "bun:test";
import { EditorState, type Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { schema } from "../schema";
import { rejectAIEditRevision } from "./comments";
import { createRestoredTableCell, tableCellContinuationPayload } from "./tableCellMergeResolution";

const paragraph = (text: string) => schema.node("paragraph", null, [schema.text(text)]);
const insertion = { kind: "ins", info: { revisionId: 99, author: "Reviewer", date: null } };

// The existing structural fixtures lacked a nested deletion followed by payload capture.
// Exercise every removed-column position and both live/payload cleanup boundaries.
for (const capture of [false, true]) {
  for (const column of [0, 1, 2]) {
    test(`revision cleanup closes nested rows: column ${column}, capture ${capture}`, () => {
      const topCells = [
        schema.node("tableCell", { rowspan: 3 }, [paragraph("Left")]),
        schema.node("tableCell", { rowspan: 2 }, [paragraph("Right")]),
      ];
      topCells.splice(
        column,
        0,
        schema.node("tableCell", { cellMarker: insertion }, [paragraph("Remove top")]),
      );
      const inner = schema.node("table", null, [
        schema.node("tableRow", null, topCells),
        schema.node("tableRow", null, [
          schema.node("tableCell", { rowspan: 2, cellMarker: insertion }, [
            paragraph("Remove span"),
          ]),
        ]),
        schema.node("tableRow", null, [schema.node("tableCell", null, [paragraph("Bottom")])]),
      ]);
      expect(TableMap.get(inner).problems).toBeNull();
      const doc = schema.node("doc", null, [
        schema.node("table", null, [
          schema.node("tableRow", null, [schema.node("tableCell", null, [paragraph("Above")])]),
          schema.node("tableRow", null, [
            schema.node(
              "tableCell",
              {
                cellMarker: capture
                  ? {
                      kind: "merge",
                      info: { revisionId: 99, author: "Reviewer", date: null },
                      verticalMergeOriginal: "continue",
                    }
                  : null,
              },
              [paragraph("Below"), inner, paragraph("Tail")],
            ),
          ]),
        ]),
      ]);
      const view = {
        state: EditorState.create({ schema, doc }),
        dispatch(tr: Transaction) {
          view.state = view.state.apply(tr);
        },
      };
      expect(rejectAIEditRevision(99)(view.state, view.dispatch)).toBe(true);
      const table = view.state.doc.firstChild;
      expect(table?.childCount).toBe(capture ? 1 : 2);
      const merged = table?.firstChild?.firstChild;
      const donor = capture ? null : table?.child(1).firstChild;
      if (!merged) throw new Error("Missing outer anchor");
      const payload = tableCellContinuationPayload(merged);
      const continuation = payload?.cells.at(0);
      const restored =
        capture && continuation ? createRestoredTableCell(merged, continuation) : donor;
      expect(Boolean(continuation)).toBe(capture);
      let tables = 0;
      restored?.descendants((node) => {
        if (node.type.name !== "table") return true;
        tables += 1;
        expect(node.childCount).toBe(2);
        expect(TableMap.get(node).problems).toBeNull();
        node.forEach((row) => expect(row.childCount).toBeGreaterThan(0));
        return true;
      });
      expect(tables).toBe(1);
    });
  }
}
