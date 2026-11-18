/**
 * The editor's table repair runs on every transaction that changes a table.
 * A canonical edit inside a table whose rows do not fill its grid, or whose
 * merges do not line up, must publish with the editor's full plugin set: the
 * projection already gives the repair nothing to change.
 */

import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

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

const fillCell = singletonManager.requireCommand("setCellFillColor")("FFCC00");

/** Every cell's first text position, so each cell in turn holds the caret. */
const cellCarets = (state: EditorState) => {
  const carets: { at: number; type: "authored" | "omitted" }[] = [];
  state.doc.descendants((node, position) => {
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      carets.push({
        at: position + 2,
        type:
          node.attrs["_omittedGridSlot"] == null
            ? ("authored" as const)
            : ("omitted" as const),
      });
      return false;
    }
    return true;
  });
  return carets;
};

describe("a canonical edit inside an irregular table", () => {
  for (const [name, rows] of Object.entries(TABLES))
    for (const mode of ["direct", "suggesting"] as const)
      test(`${name} publishes through the editor's table repair (${mode})`, async () => {
        const source = await parseDocx(
          (
            await prepareCanonicalDocxInput(await createDocx(documentOf(rows)))
          ).unwrap(),
          { preloadFonts: false, detectVariables: false },
        );
        const session = createCanonicalSession(source).unwrap();
        if (mode === "suggesting")
          session.setMode({ type: "suggesting", author: "Reviewer" });
        let state = EditorState.create({
          schema,
          doc: session.projection.doc,
          plugins: [...singletonManager.getPlugins()],
        });
        const carets = cellCarets(state);
        expect(carets.length).toBeGreaterThan(0);
        for (const caret of carets) {
          const placed = state.apply(
            state.tr.setSelection(TextSelection.create(state.doc, caret.at)),
          );
          const intents = getCanonicalCommandIntents(fillCell, placed);
          expect(intents).toBeDefined();
          if (intents === undefined) return;
          const before = session.document;
          const version = session.version;
          const canUndo = session.canUndo;
          const commit = prepareCanonicalCommands(session, placed, intents);
          if (caret.type === "omitted") {
            expect(commit.isErr()).toBe(true);
            expect(session.document).toBe(before);
            expect(session.version).toBe(version);
            expect(session.canUndo).toBe(canUndo);
            expect(placed.doc.eq(session.projection.doc)).toBe(true);
            state = placed;
            continue;
          }
          expect(commit.isErr() ? commit.error.message : "prepared").toBe(
            "prepared",
          );
          if (commit.isErr()) return;
          const published = publishCanonicalProjection({
            state: placed,
            commit: commit.value,
            session,
          });
          expect(
            published.isErr() ? published.error.message : "published",
          ).toBe("published");
          if (published.isErr()) return;
          state = published.value.state;
          expect(state.doc.eq(session.projection.doc)).toBe(true);
          const edited = session.document;
          const after = state.doc;
          const tokens = (doc: typeof state.doc) => {
            const owners: { id: unknown; token: unknown }[] = [];
            doc.descendants((node) => {
              if (
                node.type.name === "tableCell" &&
                node.attrs["_omittedGridSlot"] != null
              ) {
                node.descendants((child) => {
                  if (child.type.name === "paragraph") {
                    expect(child.attrs["paraId"]).toBeNull();
                    expect(child.attrs["_docxParagraphSourceToken"]).toBeNull();
                  }
                });
                return false;
              }
              if (node.type.name === "paragraph")
                owners.push({
                  id: node.attrs["paraId"],
                  token: node.attrs["_docxParagraphSourceToken"],
                });
              return true;
            });
            for (const { id, token } of owners) {
              expect(typeof id).toBe("string");
              expect(typeof token).toBe("string");
            }
            expect(new Set(owners.map(({ id }) => id)).size).toBe(
              owners.length,
            );
            expect(new Set(owners.map(({ token }) => token)).size).toBe(
              owners.length,
            );
            return owners;
          };
          expect(tokens(after)).toEqual(tokens(placed.doc));
          const undone = publishCanonicalProjection({
            state,
            commit: session.prepareUndo(state).unwrap(),
            session,
          }).unwrap();
          expect(session.document).toStrictEqual(before);
          expect(undone.state.doc.eq(placed.doc)).toBe(true);
          const redone = publishCanonicalProjection({
            state: undone.state,
            commit: session.prepareRedo(undone.state).unwrap(),
            session,
          }).unwrap();
          expect(session.document).toStrictEqual(edited);
          expect(redone.state.doc.eq(after)).toBe(true);
          state = redone.state;
        }
      });
});
