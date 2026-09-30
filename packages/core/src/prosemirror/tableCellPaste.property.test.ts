import { expect, test } from "bun:test";
import fc from "fast-check";
import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { buildTableDocx } from "../__tests__/tableOperationDocument";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  placeSelection,
  resolveAllChanges,
  saveHarnessState,
  summarizeState,
} from "../__tests__/editorHarness";
import { acceptChange, rejectChange } from "./commands/comments";

const assertTableGrids = (doc: PMNode): void => {
  doc.descendants((node) => {
    if (node.type.name === "table") expect(TableMap.get(node).problems ?? []).toEqual([]);
  });
};

// The old paste tests generated only unmerged clipboard cells; their
// accept/reject oracle never exercised a pasted merge over nonempty cells.
test(
  "pasted vertical spans resolve like direct replacement across table sizes",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 2, max: 5 }),
        fc.integer({ min: 1, max: 3 }),
        async (rowspan, column) => {
          const base = await parseShapeDocument(
            new Uint8Array(
              await buildTableDocx({
                rows: rowspan + 2,
                columns: column + 2,
                cells: Array.from({ length: rowspan + 2 }, (_, row) =>
                  Array.from({ length: column + 2 }, (_cell, col) => ({
                    row,
                    column: col,
                    rowSpan: 1,
                    columnSpan: 1,
                    text: `R${row}C${col}`,
                  })),
                ).flat(),
              }),
            ),
          );
          const paste = (mode: "editing" | "suggesting") => {
            const state = placeSelection(
              createHarnessState(base, mode),
              `R1C${column}`,
              "caret-end",
            );
            if (!state) throw new Error("No cell selection");
            const { schema } = state;
            const cell = schema.node("tableCell", { rowspan }, [
              schema.node("paragraph", null, schema.text("Merged paste")),
            ]);
            const rows = [schema.node("tableRow", null, [cell])];
            for (let row = 1; row < rowspan; row++) rows.push(schema.node("tableRow"));
            const view = new HeadlessEditorView(state);
            view.paste(new Slice(Fragment.from(schema.node("table", null, rows)), 0, 0));
            return view.state;
          };
          const direct = paste("editing");
          const tracked = paste("suggesting");
          const saved = await saveHarnessState(tracked, base);
          const reopened = createHarnessState(await parseShapeDocument(saved.bytes), "suggesting");
          for (const state of [tracked, reopened]) {
            const accepted = resolveAllChanges(state, "accept");
            const rejected = resolveAllChanges(state, "reject");
            assertTableGrids(accepted.doc);
            assertTableGrids(rejected.doc);
            expect(summarizeState(accepted)).toEqual(summarizeState(direct));
            expect(summarizeState(rejected)).toEqual(
              summarizeState(createHarnessState(base, "editing")),
            );
            for (const mode of ["accept", "reject"] as const) {
              const view = new HeadlessEditorView(state);
              const command = mode === "accept" ? acceptChange : rejectChange;
              expect(command(0, state.doc.content.size)(state, view.dispatch)).toBe(true);
              expect(summarizeState(view.state)).toEqual(
                summarizeState(mode === "accept" ? accepted : rejected),
              );
              assertTableGrids(view.state.doc);
            }
          }
        },
      ),
      { numRuns: 12 },
    );
  },
  propertyTestTimeout(20_000),
);
