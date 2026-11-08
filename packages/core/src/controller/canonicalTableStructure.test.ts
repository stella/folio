import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import type { Document, TableCell } from "../types/document";
import {
  type CanonicalSession,
  createCanonicalSession,
} from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";

const cell = (paraId: string, text: string, gridSpan?: number): TableCell => ({
  type: "tableCell",
  ...(gridSpan === undefined ? {} : { formatting: { gridSpan } }),
  content: [
    {
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

/** Three columns of uneven width; the first row's last two cells are one. */
const tableDocument = (): Document => ({
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
        {
          type: "table",
          columnWidths: [1000, 3000, 2000],
          rows: [
            {
              type: "tableRow",
              cells: [cell("21000000", "A"), cell("22000000", "BC", 2)],
            },
            {
              type: "tableRow",
              cells: [
                cell("23000000", "D"),
                cell("24000000", "E"),
                cell("25000000", "F"),
              ],
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

const sourceDocument = async () =>
  parseDocx(
    (
      await prepareCanonicalDocxInput(await createDocx(tableDocument()))
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );

const cellStart = (state: EditorState, text: string) => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (
      (node.type.name === "tableCell" || node.type.name === "tableHeader") &&
      node.textContent === text
    )
      found = position;
    return true;
  });
  return found ?? panic(`No cell ${text}`);
};

type Selection =
  { type: "caret"; in: string } | { type: "cells"; from: string; to: string };

const select = (state: EditorState, selection: Selection) =>
  state.apply(
    state.tr.setSelection(
      selection.type === "caret"
        ? TextSelection.create(state.doc, cellStart(state, selection.in) + 2)
        : CellSelection.create(
            state.doc,
            cellStart(state, selection.from),
            cellStart(state, selection.to),
          ),
    ),
  );

/** Whether suggesting mode records the edit, or the engine refuses it. */
type Case = {
  name: string;
  command: string;
  selection: Selection;
  tracked: boolean;
};

const CASES: readonly Case[] = [
  {
    name: "insert a column left",
    command: "addColumnLeft",
    selection: { type: "caret", in: "E" },
    tracked: true,
  },
  {
    name: "insert a column right",
    command: "addColumnRight",
    selection: { type: "caret", in: "F" },
    tracked: true,
  },
  {
    name: "delete a column",
    command: "deleteColumn",
    selection: { type: "caret", in: "D" },
    tracked: true,
  },
  {
    name: "delete two columns",
    command: "deleteColumn",
    selection: { type: "cells", from: "E", to: "F" },
    tracked: false,
  },
  {
    name: "merge a row's cells",
    command: "mergeCells",
    selection: { type: "cells", from: "D", to: "E" },
    tracked: false,
  },
  {
    name: "merge a column's cells",
    command: "mergeCells",
    selection: { type: "cells", from: "A", to: "D" },
    tracked: false,
  },
  {
    name: "split a spanning cell",
    command: "splitCell",
    selection: { type: "caret", in: "BC" },
    tracked: true,
  },
  {
    name: "distribute columns",
    command: "distributeColumns",
    selection: { type: "caret", in: "A" },
    tracked: true,
  },
];

const commandFor = (name: string) =>
  (singletonManager.requireCommand(name as never) as () => Command)();

/** What the table says, without identities the two paths allocate apart. */
const shapeOf = (document: Document) => {
  const table = document.package.document.content.find(
    (block) => block.type === "table",
  );
  if (table?.type !== "table") return panic("No table.");
  return {
    grid: table.columnWidths,
    width: table.formatting?.width,
    layout: table.formatting?.layout,
    rows: table.rows.map((row) =>
      row.cells.map((each) => ({
        text: each.content
          .map((block) =>
            block.type === "paragraph"
              ? block.content
                  .flatMap((item) =>
                    item.type === "run"
                      ? item.content.flatMap((leaf) =>
                          leaf.type === "text" ? [leaf.text] : [],
                        )
                      : [],
                  )
                  .join("")
              : block.type,
          )
          .join("|"),
        gridSpan: each.formatting?.gridSpan ?? 1,
        vMerge: each.formatting?.vMerge,
        width: each.formatting?.width,
      })),
    ),
  };
};

/** The editor's own saved table after running the command directly. */
const editorOutcome = (state: EditorState, source: Document, name: string) => {
  let next = state;
  commandFor(name)(state, (transaction) => {
    next = state.apply(transaction);
  });
  return shapeOf(
    fromProseDoc(next.doc, source, { stylesheetSource: { type: "package" } }),
  );
};

const run = (session: CanonicalSession, state: EditorState, name: string) => {
  const intents = getCanonicalCommandIntents(commandFor(name), state);
  if (intents === undefined) return panic(`${name} has no lowering.`);
  return prepareCanonicalCommands(session, state, intents);
};

describe("table structure commands in the canonical session", () => {
  for (const entry of CASES)
    test(`${entry.name} applies what the editor would save, and undoes`, async () => {
      const session = createCanonicalSession(await sourceDocument()).unwrap();
      const state = select(
        EditorState.create({ schema, doc: session.projection.doc }),
        entry.selection,
      );
      const expected = editorOutcome(state, session.document, entry.command);
      const before = session.document;
      const commit = run(session, state, entry.command).unwrap();
      expect(commit.publish().isOk()).toBe(true);
      expect(shapeOf(session.document)).toEqual(expected);
      const afterState = state.apply(commit.transaction);
      expect(afterState.doc.eq(session.projection.doc)).toBe(true);
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
        featureFlags: { selectiveSave: true },
      });
      expect(
        shapeOf(await parseDocx(saved.buffer, { preloadFonts: false })),
      ).toEqual(expected);
      const undone = session.prepareUndo(afterState).unwrap();
      expect(undone.publish().isOk()).toBe(true);
      expect(session.document).toEqual(before);
    });

  for (const entry of CASES)
    test(`${entry.name} in suggesting mode ${entry.tracked ? "is tracked" : "refuses and changes nothing"}`, async () => {
      const session = createCanonicalSession(await sourceDocument()).unwrap();
      session.setMode({ type: "suggesting", author: "Reviewer" });
      const state = select(
        EditorState.create({ schema, doc: session.projection.doc }),
        entry.selection,
      );
      const before = session.document;
      const prepared = run(session, state, entry.command);
      expect(prepared.isOk()).toBe(entry.tracked);
      if (prepared.isErr()) {
        expect(session.document).toBe(before);
        return;
      }
      expect(prepared.value.publish().isOk()).toBe(true);
      expect(JSON.stringify(session.document.package.document)).toMatch(
        /"insertion"|"deletion"|PropertyChange|propertyChange|cellIns|cellDel|tcPrChange|tblGridChange/u,
      );
    });
});
