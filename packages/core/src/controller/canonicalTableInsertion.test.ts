import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import type { Document } from "../types/document";
import { createCanonicalSession } from "./canonicalSession";
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
