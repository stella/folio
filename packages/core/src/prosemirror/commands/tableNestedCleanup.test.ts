import { expect, test } from "bun:test";
import { EditorState, type Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { applyFolioAIEditOperations } from "../../ai-edits/apply";
import { createFolioAIEditSnapshot } from "../../ai-edits/snapshot";
import { schema } from "../schema";
import { rejectAIEditRevision } from "./comments";
import {
  createRestoredTableCell,
  tableCellContinuationPayload,
  tableCellContinuationFromNode,
} from "./tableCellMergeResolution";

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
      const unaffected = schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { rowspan: 2 }, [paragraph("Keep nested")]),
        ]),
        schema.node("tableRow"),
      ]);
      expect(TableMap.get(unaffected).problems).toBeNull();
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
              [paragraph("Below"), inner, unaffected, paragraph("Tail")],
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
      expect(table?.childCount).toBe(2);
      const merged = table?.firstChild?.firstChild;
      const donor = capture ? null : table?.child(1).firstChild;
      if (!merged) throw new Error("Missing outer anchor");
      const payload = tableCellContinuationPayload(merged);
      const continuation = payload?.cells.at(0);
      const restored =
        capture && continuation ? createRestoredTableCell(merged, continuation) : donor;
      let expectedUnaffected = unaffected;
      if (capture) {
        const originalDonor = doc.firstChild?.child(1).firstChild;
        if (!originalDonor) throw new Error("Missing original donor");
        // Compare through the same continuation conversion, which materializes
        // vertical continuation cells; cleanup must add no changes of its own.
        const baseline = createRestoredTableCell(
          merged,
          tableCellContinuationFromNode(originalDonor),
        );
        let baselineTables = 0;
        baseline?.descendants((node) => {
          if (node.type.name !== "table") return true;
          baselineTables += 1;
          if (baselineTables === 2) expectedUnaffected = node;
          return true;
        });
        expect(baselineTables).toBe(2);
      }
      expect(Boolean(continuation)).toBe(capture);
      let tables = 0;
      restored?.descendants((node) => {
        if (node.type.name !== "table") return true;
        tables += 1;
        if (tables === 2) expect(node.eq(expectedUnaffected)).toBe(true);
        expect(node.childCount).toBe(2);
        expect(TableMap.get(node).problems).toBeNull();
        if (tables === 1) node.forEach((row) => expect(row.childCount).toBeGreaterThan(0));
        node.forEach((row) => expect(row.attrs["_batchRowCleanup"]).toBeNull());
        return true;
      });
      expect(tables).toBe(2);
    });
  }
}

for (const mode of ["revision", "batch"] as const) {
  for (const existingRow of ["separate", "same"] as const) {
    test(`${mode} cleanup preserves an unrelated fully spanned row (${existingRow} table)`, () => {
      const untouched = schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { rowspan: 2 }, [paragraph("Keep")]),
        ]),
        schema.node("tableRow"),
      ]);
      expect(TableMap.get(untouched).problems).toBeNull();
      const changed = schema.node("table", null, [
        schema.node("tableRow", null, [
          schema.node("tableCell", { rowspan: existingRow === "same" ? 3 : 2 }, [
            paragraph("Left"),
          ]),
          schema.node(
            "tableCell",
            { rowspan: existingRow === "same" ? 2 : 1, cellMarker: insertion },
            [paragraph("Remove")],
          ),
        ]),
        ...(existingRow === "same" ? [schema.node("tableRow")] : []),
        schema.node("tableRow", null, [
          schema.node("tableCell", { rowspan: 2, cellMarker: insertion }, [paragraph("Span")]),
        ]),
        schema.node("tableRow", null, [schema.node("tableCell", null, [paragraph("Bottom left")])]),
      ]);
      expect(TableMap.get(changed).problems).toBeNull();
      const doc = schema.node("doc", null, [untouched, changed]);
      const view = {
        state: EditorState.create({ schema, doc }),
        dispatch(tr: Transaction) {
          view.state = view.state.apply(tr);
        },
      };
      if (mode === "revision") {
        expect(rejectAIEditRevision(99)(view.state, view.dispatch)).toBe(true);
      } else {
        const snapshot = createFolioAIEditSnapshot(doc);
        const block = snapshot.blocks.find(({ text }) => text === "Remove");
        if (!block) throw new Error("Missing column target");
        const result = applyFolioAIEditOperations({
          view,
          snapshot,
          mode: "direct",
          operations: [{ id: "delete", type: "deleteTableColumn", blockId: block.id }],
        });
        expect(result.applied).toHaveLength(1);
      }
      expect(view.state.doc.firstChild?.eq(untouched)).toBe(true);
      expect(view.state.doc.lastChild?.childCount).toBe(existingRow === "same" ? 3 : 2);
      if (existingRow === "same") expect(view.state.doc.lastChild?.child(1).childCount).toBe(0);
      view.state.doc.descendants((node) => {
        if (node.type.spec["tableRole"] === "row")
          expect(node.attrs["_batchRowCleanup"]).toBeNull();
      });
    });
  }
}
