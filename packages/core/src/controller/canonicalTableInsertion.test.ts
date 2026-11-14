import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { AllSelection, EditorState, TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import type { Document } from "../types/document";
import { prepareCanonicalCommands } from "./canonicalStructure";

const source = async () =>
  parseDocx(
    (
      await prepareCanonicalDocxInput(
        await createDocx({
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  content: [
                    {
                      type: "run",
                      content: [{ type: "text", text: "Before" }],
                    },
                  ],
                },
                {
                  type: "paragraph",
                  content: [
                    { type: "run", content: [{ type: "text", text: "After" }] },
                  ],
                },
              ],
            },
          },
        }),
      )
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );

const insertTable = () =>
  (
    singletonManager.requireCommand("insertTable" as never) as (
      rows: number,
      cols: number,
    ) => Command
  )(2, 3);

/** The body as block kinds, a table as its row and cell counts. */
const shapeOf = (document: Document) =>
  document.package.document.content.map((block) => {
    if (block.type === "table")
      return {
        table: block.rows.map((row) => row.cells.length),
        grid: block.columnWidths,
      };
    if (block.type !== "paragraph") return block.type;
    return block.content.length === 0 ? "empty" : "paragraph";
  });

describe("insert table in the canonical session", () => {
  for (const mode of ["editing", "suggesting"] as const)
    test(`inserts what the editor would save, in ${mode} mode, and undoes`, async () => {
      const session = createCanonicalSession(await source()).unwrap();
      if (mode === "suggesting")
        session.setMode({ type: "suggesting", author: "Reviewer" });
      let state = EditorState.create({ schema, doc: session.projection.doc });
      state = state.apply(
        state.tr.setSelection(TextSelection.create(state.doc, 3)),
      );
      let edited = state;
      insertTable()(state, (transaction) => {
        edited = state.apply(transaction);
      });
      const expected = shapeOf(
        fromProseDoc(edited.doc, session.document, {
          stylesheetSource: { type: "package" },
        }),
      );
      const before = session.document;
      const intents = getCanonicalCommandIntents(insertTable(), state);
      if (intents === undefined) return panic("insertTable has no lowering.");
      const commit = prepareCanonicalCommands(session, state, intents).unwrap();
      expect(commit.publish().isOk()).toBe(true);
      expect(shapeOf(session.document)).toEqual(expected);
      if (mode === "suggesting")
        expect(JSON.stringify(session.document.package.document)).toMatch(
          /tableRowInsertion/u,
        );
      const saved = await serializeCanonicalSave({
        snapshot: session.captureSaveSnapshot(),
        featureFlags: { selectiveSave: true },
      });
      expect(
        shapeOf(await parseDocx(saved.buffer, { preloadFonts: false })),
      ).toEqual(expected);
      const afterState = state.apply(commit.transaction);
      expect(session.prepareUndo(afterState).unwrap().publish().isOk()).toBe(
        true,
      );
      expect(session.document).toEqual(before);
    });
});

describe("insert table with the caret in an inline content control", () => {
  test("places the table after the control's paragraph, as the editor does", async () => {
    const document = await parseDocx(
      (
        await prepareCanonicalDocxInput(
          await createDocx({
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    content: [
                      {
                        type: "run",
                        content: [{ type: "text", text: "Before " }],
                      },
                      {
                        type: "inlineSdt",
                        properties: { sdtType: "richText", id: 7 },
                        content: [
                          {
                            type: "run",
                            content: [{ type: "text", text: "Clause" }],
                          },
                        ],
                      },
                    ],
                  },
                  {
                    type: "paragraph",
                    content: [
                      {
                        type: "run",
                        content: [{ type: "text", text: "After" }],
                      },
                    ],
                  },
                ],
              },
            },
          }),
        )
      ).unwrap(),
      { preloadFonts: false, detectVariables: false },
    );
    const session = createCanonicalSession(document).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    let inside: number | undefined;
    state.doc.descendants((node, position) => {
      if (node.isText && node.text === "Clause") inside = position + 3;
      return inside === undefined;
    });
    if (inside === undefined) return panic("No control text.");
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, inside)),
    );
    let edited = state;
    insertTable()(state, (transaction) => {
      edited = state.apply(transaction);
    });
    const expected = shapeOf(
      fromProseDoc(edited.doc, session.document, {
        stylesheetSource: { type: "package" },
      }),
    );
    const intents = getCanonicalCommandIntents(insertTable(), state);
    if (intents === undefined) return panic("insertTable has no lowering.");
    const commit = prepareCanonicalCommands(session, state, intents).unwrap();
    expect(commit.publish().isOk()).toBe(true);
    expect(shapeOf(session.document)).toEqual(expected);
  });
});

describe("insert table from a selection with no caret paragraph", () => {
  /** A paragraph, a one-cell table, and the paragraph after it. */
  const withTable = async () =>
    parseDocx(
      (
        await prepareCanonicalDocxInput(
          await createDocx({
            package: {
              document: {
                content: [
                  {
                    type: "paragraph",
                    content: [
                      {
                        type: "run",
                        content: [{ type: "text", text: "Before" }],
                      },
                    ],
                  },
                  {
                    type: "table",
                    rows: [
                      {
                        type: "tableRow",
                        cells: [
                          {
                            type: "tableCell",
                            content: [
                              {
                                type: "paragraph",
                                content: [
                                  {
                                    type: "run",
                                    content: [{ type: "text", text: "A" }],
                                  },
                                ],
                              },
                            ],
                          },
                        ],
                      },
                    ],
                  },
                  {
                    type: "paragraph",
                    content: [
                      {
                        type: "run",
                        content: [{ type: "text", text: "After" }],
                      },
                    ],
                  },
                ],
              },
            },
          }),
        )
      ).unwrap(),
      { preloadFonts: false, detectVariables: false },
    );

  for (const mode of ["editing", "suggesting"] as const)
    for (const selection of ["whole document", "selected cells"] as const)
      test(`over the ${selection} it inserts what the editor would save, in ${mode} mode`, async () => {
        const session = createCanonicalSession(await withTable()).unwrap();
        if (mode === "suggesting")
          session.setMode({ type: "suggesting", author: "Reviewer" });
        let state = EditorState.create({ schema, doc: session.projection.doc });
        if (selection === "whole document")
          state = state.apply(
            state.tr.setSelection(new AllSelection(state.doc)),
          );
        else {
          let cell = -1;
          state.doc.descendants((node, position) => {
            if (cell < 0 && node.type.name === "tableCell") cell = position;
            return cell < 0;
          });
          state = state.apply(
            state.tr.setSelection(CellSelection.create(state.doc, cell)),
          );
        }
        let edited = state;
        insertTable()(state, (transaction) => {
          edited = state.apply(transaction);
        });
        const expected = shapeOf(
          fromProseDoc(edited.doc, session.document, {
            stylesheetSource: { type: "package" },
          }),
        );
        const intents = getCanonicalCommandIntents(insertTable(), state);
        if (intents === undefined) return panic("insertTable has no lowering.");
        const commit = prepareCanonicalCommands(
          session,
          state,
          intents,
        ).unwrap();
        expect(commit.publish().isOk()).toBe(true);
        expect(shapeOf(session.document)).toEqual(expected);
      });
});
