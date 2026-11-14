/**
 * The editor's table repair runs on every transaction that changes a table.
 * A canonical edit inside a table whose rows do not fill its grid, or whose
 * merges do not line up, must publish with the editor's full plugin set: the
 * projection already gives the repair nothing to change.
 */

import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";

import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { schema, singletonManager } from "../prosemirror/schema";
import type {
  Document,
  TableCell,
  TableCellFormatting,
  TableRow,
} from "../types/document";
import { prepareCanonicalCommands } from "./canonicalStructure";

const cell = (
  paraId: string,
  text: string,
  formatting?: TableCellFormatting,
): TableCell => ({
  type: "tableCell",
  ...(formatting === undefined ? {} : { formatting }),
  content: [
    {
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

const row = (...cells: TableCell[]): TableRow => ({ type: "tableRow", cells });

/** Tables the editor's repair would change if the projection left them so. */
const TABLES: Readonly<Record<string, readonly TableRow[]>> = {
  "a row with fewer cells than the grid": [
    row(cell("21000000", "A1"), cell("22000000", "B1"), cell("23000000", "C1")),
    row(cell("24000000", "A2")),
  ],
  "a row with more cells than the grid": [
    row(cell("21000000", "A1"), cell("22000000", "B1")),
    row(cell("23000000", "A2"), cell("24000000", "B2"), cell("25000000", "C2")),
  ],
  "a merge continued with no cell starting it": [
    row(cell("21000000", "", { vMerge: "continue" }), cell("22000000", "B1")),
    row(cell("23000000", "A2"), cell("24000000", "B2")),
  ],
  "a merge started in the last row": [
    row(cell("21000000", "A1"), cell("22000000", "B1")),
    row(cell("23000000", "A2", { vMerge: "restart" }), cell("24000000", "B2")),
  ],
  "a span reaching past the grid": [
    row(cell("21000000", "A1", { gridSpan: 3 }), cell("22000000", "B1")),
    row(cell("23000000", "A2"), cell("24000000", "B2")),
  ],
};

const documentOf = (rows: readonly TableRow[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "10000000",
          content: [
            { type: "run", content: [{ type: "text", text: "Before" }] },
          ],
        },
        { type: "table", columnWidths: [2000, 2000], rows: [...rows] },
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

const fillCell = (
  singletonManager.requireCommand("setCellFillColor" as never) as (
    color: string,
  ) => Command
)("FFCC00");

/** Every cell's first text position, so each cell in turn holds the caret. */
const cellCarets = (state: EditorState) => {
  const carets: number[] = [];
  state.doc.descendants((node, position) => {
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      carets.push(position + 2);
      return false;
    }
    return true;
  });
  return carets;
};

describe("a canonical edit inside an irregular table", () => {
  for (const [name, rows] of Object.entries(TABLES))
    test(`${name} publishes through the editor's table repair`, async () => {
      const source = await parseDocx(
        (
          await prepareCanonicalDocxInput(await createDocx(documentOf(rows)))
        ).unwrap(),
        { preloadFonts: false, detectVariables: false },
      );
      const session = createCanonicalSession(source).unwrap();
      let state = EditorState.create({
        schema,
        doc: session.projection.doc,
        plugins: [...singletonManager.getPlugins()],
      });
      const carets = cellCarets(state);
      expect(carets.length).toBeGreaterThan(0);
      for (const caret of carets) {
        const placed = state.apply(
          state.tr.setSelection(TextSelection.create(state.doc, caret)),
        );
        const intents = getCanonicalCommandIntents(fillCell, placed);
        expect(intents).toBeDefined();
        if (intents === undefined) return;
        const commit = prepareCanonicalCommands(session, placed, intents);
        expect(commit.isOk()).toBe(true);
        if (commit.isErr()) return;
        const published = publishCanonicalProjection({
          state: placed,
          commit: commit.value,
          session,
        });
        expect(published.isErr() ? published.error.message : "published").toBe(
          "published",
        );
        if (published.isErr()) return;
        state = published.value.state;
        expect(state.doc.eq(session.projection.doc)).toBe(true);
      }
    });
});
