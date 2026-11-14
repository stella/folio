import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import {
  type CanonicalSession,
  createCanonicalSession,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { lowerTableAttributeCommand } from "../prosemirror/canonicalTableCommands";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import {
  columnResizeCommand,
  rightEdgeResizeCommand,
  rowResizeCommand,
} from "../prosemirror/tableResize";
import type { Document } from "../types/document";
import { prepareCanonicalCommands } from "./canonicalStructure";

const cell = (paraId: string, text: string) => ({
  type: "tableCell" as const,
  content: [
    {
      type: "paragraph" as const,
      paraId,
      content: [
        { type: "run" as const, content: [{ type: "text" as const, text }] },
      ],
    },
  ],
});

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

/** The table as a saved package, so a save has a source to splice into. */
const sourceDocument = async () =>
  parseDocx(
    (
      await prepareCanonicalDocxInput(await createDocx(tableDocument()))
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );

const positionAfter = (state: EditorState, text: string) => {
  let found: number | undefined;
  state.doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (node.isText && node.text === text) found = position + text.length;
    return true;
  });
  return found ?? panic(`No text ${text}`);
};

type Case = {
  name: string;
  command: string;
  args: readonly unknown[];
  selection: "caret" | "cells";
};

const CASES: readonly Case[] = [
  {
    name: "cell fill",
    command: "setCellFillColor",
    args: ["FFCC00"],
    selection: "cells",
  },
  {
    name: "vertical alignment",
    command: "setCellVerticalAlign",
    args: ["center"],
    selection: "caret",
  },
  {
    name: "cell margins",
    command: "setCellMargins",
    args: [{ top: 120, left: 200 }],
    selection: "caret",
  },
  {
    name: "text direction",
    command: "setCellTextDirection",
    args: ["btLr"],
    selection: "caret",
  },
  { name: "no wrap", command: "toggleNoWrap", args: [], selection: "cells" },
  {
    name: "row height",
    command: "setRowHeight",
    args: [600, "exact"],
    selection: "caret",
  },
  {
    name: "table justification",
    command: "setTableProperties",
    args: [{ justification: "center" }],
    selection: "caret",
  },
  {
    name: "cell border",
    command: "setCellBorder",
    args: ["all", { style: "single", size: 8, color: { rgb: "FF0000" } }],
    selection: "cells",
  },
  {
    name: "table borders",
    command: "setTableBorders",
    args: ["all"],
    selection: "caret",
  },
];

const select = (state: EditorState, selection: Case["selection"]) => {
  if (selection === "caret")
    return state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, positionAfter(state, "A")),
      ),
    );
  // From the first cell to the second cell of the first row.
  const cells: number[] = [];
  state.doc.descendants((node, position) => {
    if (node.type.name === "tableCell") cells.push(position);
    return true;
  });
  const [anchor, head] = cells;
  if (anchor === undefined || head === undefined)
    panic("The table has no cells.");
  return state.apply(
    state.tr.setSelection(CellSelection.create(state.doc, anchor, head)),
  );
};

/** Parse keeps each element's source XML; the comparison is about values. */
const withoutSource = (value: unknown): unknown =>
  JSON.parse(
    JSON.stringify(value, (key, field: unknown) =>
      /sourceXml$/iu.test(key) ? undefined : field,
    ),
  );

/** The editor's own saved document after running the command directly. */
const editorDocument = (state: EditorState, source: Document, run: Case) => {
  const command = singletonManager.requireCommand(run.command as never) as (
    ...args: readonly unknown[]
  ) => Parameters<typeof lowerTableAttributeCommand>[0];
  let next = state;
  command(...run.args)(state, (transaction) => {
    next = state.apply(transaction);
  });
  return fromProseDoc(next.doc, source, {
    stylesheetSource: { type: "package" },
  });
};

/** The editor's own saved table formatting after running the command directly. */
const editorOutcome = (state: EditorState, source: Document, run: Case) => {
  const command = singletonManager.requireCommand(run.command as never) as (
    ...args: readonly unknown[]
  ) => Parameters<typeof lowerTableAttributeCommand>[0];
  let next = state;
  command(...run.args)(state, (transaction) => {
    next = state.apply(transaction);
  });
  return tableOf(
    fromProseDoc(next.doc, source, { stylesheetSource: { type: "package" } }),
  );
};

const tableOf = (document: Document) => {
  const table = document.package.document.content.find(
    (block) => block.type === "table",
  );
  if (table?.type !== "table") return panic("No table.");
  return {
    table: table.formatting ?? {},
    rows: table.rows.map((row) => row.formatting ?? {}),
    cells: table.rows.map((row) =>
      row.cells.map((each) => each.formatting ?? {}),
    ),
  };
};

const run = (session: CanonicalSession, state: EditorState, entry: Case) => {
  const command = singletonManager.requireCommand(entry.command as never) as (
    ...args: readonly unknown[]
  ) => Parameters<typeof lowerTableAttributeCommand>[0];
  const intents = getCanonicalCommandIntents(command(...entry.args), state);
  if (intents === undefined) return panic(`${entry.name} has no lowering.`);
  expect(intents.length).toBe(1);
  return prepareCanonicalCommands(session, state, intents).unwrap();
};

describe("table attribute commands in the canonical session", () => {
  for (const entry of CASES) {
    test(`${entry.name} applies what the editor would save, tracks in suggesting mode, and undoes`, async () => {
      const editing = createCanonicalSession(await sourceDocument()).unwrap();
      const state = select(
        EditorState.create({ schema, doc: editing.projection.doc }),
        entry.selection,
      );
      const expected = editorOutcome(state, editing.document, entry);
      const before = editing.document;
      const commit = run(editing, state, entry);
      expect(commit.publish().isOk()).toBe(true);
      expect(tableOf(editing.document)).toEqual(expected);
      expect(editing.canUndo).toBe(true);
      const afterState = state.apply(commit.transaction);
      const undone = editing.prepareUndo(afterState).unwrap();
      expect(undone.publish().isOk()).toBe(true);
      expect(editing.document).toEqual(before);
      const redone = editing
        .prepareRedo(afterState.apply(undone.transaction))
        .unwrap();
      expect(redone.publish().isOk()).toBe(true);
      expect(tableOf(editing.document)).toEqual(expected);

      // The editor's save, which splices only what changed into the source.
      const saved = await serializeCanonicalSave({
        snapshot: editing.captureSaveSnapshot(),
        featureFlags: { selectiveSave: true },
      });
      // The file the session writes reads back as the file the editor writes.
      const editorSaved = await createDocx(
        editorDocument(state, editing.document, entry),
      );
      expect(
        withoutSource(
          tableOf(await parseDocx(saved.buffer, { preloadFonts: false })),
        ),
      ).toEqual(
        withoutSource(
          tableOf(await parseDocx(editorSaved, { preloadFonts: false })),
        ),
      );

      const suggesting = createCanonicalSession(tableDocument()).unwrap();
      suggesting.setMode({ type: "suggesting", author: "Reviewer" });
      const tracked = select(
        EditorState.create({ schema, doc: suggesting.projection.doc }),
        entry.selection,
      );
      expect(run(suggesting, tracked, entry).publish().isOk()).toBe(true);
      const changes = JSON.stringify(suggesting.document.package.document);
      expect(changes).toMatch(
        /PropertyChange|propertyChange|tcPrChange|trPrChange|tblPrChange/u,
      );
    });
  }

  test("a command that changes more than table attributes has no lowering", () => {
    const session = createCanonicalSession(tableDocument()).unwrap();
    const state = select(
      EditorState.create({ schema, doc: session.projection.doc }),
      "caret",
    );
    const addRow = singletonManager.requireCommand(
      "addRowBelow" as never,
    ) as () => Parameters<typeof lowerTableAttributeCommand>[0];
    expect(lowerTableAttributeCommand(addRow(), state)).toBeUndefined();
  });

  test("a cell that opens with a nested table addresses the outer cell", async () => {
    const document = tableDocument();
    const table = document.package.document.content[1];
    if (table?.type !== "table") return panic("No table.");
    const outer = table.rows[0]?.cells[0] ?? panic("No cell.");
    // The cell's first block is a table of its own; the outer cell's own
    // paragraph comes after it.
    outer.content.unshift({
      type: "table",
      rows: [{ type: "tableRow", cells: [cell("25000000", "Inner")] }],
    });
    const session = createCanonicalSession(document).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    state = state.apply(
      state.tr.setSelection(
        TextSelection.create(state.doc, positionAfter(state, "A")),
      ),
    );
    const fill = singletonManager.requireCommand(
      "setCellFillColor" as never,
    ) as (color: string) => Parameters<typeof lowerTableAttributeCommand>[0];
    const intents = getCanonicalCommandIntents(fill("FF0000"), state);
    if (intents === undefined) return panic("The fill has no lowering.");
    expect(
      prepareCanonicalCommands(session, state, intents)
        .unwrap()
        .publish()
        .isOk(),
    ).toBe(true);
    const saved = session.document.package.document.content[1];
    if (saved?.type !== "table") return panic("No table.");
    const savedOuter = saved.rows[0]?.cells[0] ?? panic("No cell.");
    expect(savedOuter.formatting?.shading).toBeDefined();
    const inner = savedOuter.content[0];
    if (inner?.type !== "table") return panic("No nested table.");
    expect(inner.rows[0]?.cells[0]?.formatting?.shading).toBeUndefined();
  });
});

describe("table resizes in the canonical session", () => {
  // A drag on a column border, the table's right edge or a row border runs a
  // resize command; the session applies what the editor would save for it.
  const sized = (): Document => {
    const document = tableDocument();
    const table = document.package.document.content[1];
    if (table?.type !== "table") return panic("No table.");
    table.columnWidths = [2400, 2400];
    return document;
  };
  const tablePosition = (state: EditorState) => {
    let found: number | undefined;
    state.doc.descendants((node, position) => {
      if (found !== undefined) return false;
      if (node.type.name === "table") found = position;
      return found === undefined;
    });
    return found ?? panic("No table.");
  };
  const grid = (document: Document) => {
    const table = document.package.document.content.find(
      (block) => block.type === "table",
    );
    return table?.type === "table" ? table.columnWidths : undefined;
  };
  const RESIZES = {
    "column border": (pmStart: number) =>
      columnResizeCommand({
        pmStart,
        colIdx: 0,
        newLeft: 3000,
        newRight: 1800,
      }),
    "right edge": (pmStart: number) =>
      rightEdgeResizeCommand({ pmStart, colIdx: 1, newWidth: 3600 }),
    "row border": (pmStart: number) =>
      rowResizeCommand({ pmStart, rowIdx: 1, newHeight: 900 }),
  } as const;
  for (const [name, resize] of Object.entries(RESIZES))
    test(`a ${name} drag applies what the editor would save, and undoes`, async () => {
      const source = await parseDocx(
        (await prepareCanonicalDocxInput(await createDocx(sized()))).unwrap(),
        { preloadFonts: false, detectVariables: false },
      );
      const session = createCanonicalSession(source).unwrap();
      const state = EditorState.create({
        schema,
        doc: session.projection.doc,
      });
      const command = resize(tablePosition(state));
      let direct = state;
      expect(
        command(state, (transaction) => {
          direct = state.apply(transaction);
        }),
      ).toBe(true);
      const expected = fromProseDoc(direct.doc, session.document, {
        stylesheetSource: { type: "package" },
      });
      const before = session.document;
      const intents = getCanonicalCommandIntents(command, state);
      if (intents === undefined) return panic(`${name} has no lowering.`);
      const commit = prepareCanonicalCommands(session, state, intents).unwrap();
      expect(commit.publish().isOk()).toBe(true);
      expect([tableOf(session.document), grid(session.document)]).toEqual([
        tableOf(expected),
        grid(expected),
      ]);
      const undone = session
        .prepareUndo(state.apply(commit.transaction))
        .unwrap();
      expect(undone.publish().isOk()).toBe(true);
      expect(session.document).toEqual(before);

      // Suggesting, the resize is recorded as a change to review.
      const suggesting = createCanonicalSession(source).unwrap();
      suggesting.setMode({ type: "suggesting", author: "Reviewer" });
      const tracked = EditorState.create({
        schema,
        doc: suggesting.projection.doc,
      });
      const trackedIntents = getCanonicalCommandIntents(
        resize(tablePosition(tracked)),
        tracked,
      );
      if (trackedIntents === undefined)
        return panic(`${name} has no lowering.`);
      expect(
        prepareCanonicalCommands(suggesting, tracked, trackedIntents)
          .unwrap()
          .publish()
          .isOk(),
      ).toBe(true);
      expect(JSON.stringify(suggesting.document.package.document)).toMatch(
        /gridChange|PropertyChange|propertyChange/u,
      );
    });
});
