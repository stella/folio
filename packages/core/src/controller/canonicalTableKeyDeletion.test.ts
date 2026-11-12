import { panic } from "better-result";
import { afterAll, beforeAll, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import { EditorView } from "prosemirror-view";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { schema } from "../prosemirror/schema";
import type { Document, TableCell } from "../types/document";
import { createCanonicalInputBoundary } from "./canonicalInput";
import { createCanonicalSession } from "./canonicalSession";

let releaseDom: ReturnType<typeof acquireHarnessDom>;
beforeAll(() => {
  releaseDom = acquireHarnessDom();
});
afterAll(() => releaseDom());

const cell = (paraId: string, text: string): TableCell => ({
  type: "tableCell",
  content: [
    {
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

/** Two rows of two cells, and a paragraph after the table. */
const tableDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [cell("21000000", "A"), cell("22000000", "B")],
            },
            {
              type: "tableRow",
              cells: [cell("23000000", "C"), cell("24000000", "D")],
            },
          ],
        },
        {
          type: "paragraph",
          paraId: "30000000",
          content: [
            { type: "run", content: [{ type: "text", text: "After" }] },
          ],
        },
      ],
    },
  },
});

/** What the input boundary asks for when the key goes down over `from`..`to` cells. */
const keyOver = (key: "Backspace" | "Delete", from: string, to: string) => {
  const session = createCanonicalSession(tableDocument()).unwrap();
  const cells = new Map<string, number>();
  const base = EditorState.create({ schema, doc: session.projection.doc });
  base.doc.descendants((node, position) => {
    if (node.type.name === "tableCell") cells.set(node.textContent, position);
    return true;
  });
  const state = base.apply(
    base.tr.setSelection(
      CellSelection.create(
        base.doc,
        cells.get(from) ?? panic(`No cell ${from}`),
        cells.get(to) ?? panic(`No cell ${to}`),
      ),
    ),
  );
  const structure: string[] = [];
  const replaced: string[] = [];
  const boundary = createCanonicalInputBoundary({
    structure: (type) => structure.push(type),
    replace: (input) => replaced.push(input.semantic ?? "replace"),
    refuse: (reason) => {
      throw new TypeError(reason);
    },
    undo: () => false,
    redo: () => false,
  });
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, { state });
  const event = new KeyboardEvent("keydown", { key, cancelable: true });
  expect(boundary.handleKeyDown(view, event)).toBe(true);
  view.destroy();
  return { structure, replaced };
};

for (const key of ["Backspace", "Delete"] as const) {
  test(`${key} over every cell of a table deletes the table`, () => {
    expect(keyOver(key, "A", "D")).toEqual({
      structure: ["deleteTable"],
      replaced: [],
    });
  });

  test(`${key} over some cells clears them and keeps the table`, () => {
    expect(keyOver(key, "A", "B").structure).toEqual([]);
  });
}
