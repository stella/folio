import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import type { Document, TableCellFormatting } from "../../types/document";
import { schema, singletonManager } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const documentWith = (formatting?: TableCellFormatting): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  ...(formatting === undefined ? {} : { formatting }),
                  content: [
                    {
                      type: "paragraph",
                      content: [
                        { type: "run", content: [{ type: "text", text: "A" }] },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
        { type: "paragraph", content: [] },
      ],
    },
  },
});

/** Toggle no-wrap in the cell the way the toolbar does, and read what saves. */
const toggledAndSaved = (source: Document) => {
  let state = EditorState.create({ schema, doc: toProseDoc(source) });
  state = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, 4)),
  );
  const toggle = singletonManager.requireCommand(
    "toggleNoWrap" as never,
  ) as () => (
    state: EditorState,
    dispatch: (transaction: EditorState["tr"]) => void,
  ) => boolean;
  expect(
    toggle()(state, (transaction) => {
      state = state.apply(transaction);
    }),
  ).toBe(true);
  const saved = fromProseDoc(state.doc, source, {
    stylesheetSource: { type: "package" },
  });
  const table = saved.package.document.content.at(0);
  if (table?.type !== "table") return panic("The table did not save.");
  return table.rows.at(0)?.cells.at(0)?.formatting;
};

test("a cell set to stop wrapping saves it", () => {
  expect(toggledAndSaved(documentWith())?.noWrap).toBe(true);
  expect(
    toggledAndSaved(documentWith({ verticalAlign: "center" }))?.noWrap,
  ).toBe(true);
});

test("a cell that wraps again drops what it stated", () => {
  const saved = toggledAndSaved(
    documentWith({ noWrap: true, verticalAlign: "center" }),
  );
  expect(saved?.noWrap).toBeUndefined();
  expect(saved?.verticalAlign).toBe("center");
});
