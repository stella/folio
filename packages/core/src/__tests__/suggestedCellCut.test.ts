/**
 * Cutting a selection of table cells while suggesting leaves the caret where
 * a direct cut leaves it, at the end of the last cell. Content pasted next
 * then lands there once, so accepting the suggestions gives what editing
 * gives, not a copy in every cut cell.
 */
import { describe, expect, test } from "bun:test";
import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import { CellSelection } from "prosemirror-tables";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { deleteSelectionAsSuggestion } from "../prosemirror/plugins/suggestionMode";
import { DOCUMENT_SHAPES } from "./documentShapes";
import {
  createHarnessState,
  type EditorMode,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
  summarizeReviewer,
} from "./editorHarness";

const pastedList = (doc: PMNode) => {
  const { schema } = doc.type;
  const item = (text: string, level: number) =>
    schema.node(
      "paragraph",
      { _pastedHtmlList: { group: 1, kind: "bullet", level } },
      schema.text(text),
    );
  return new Slice(
    Fragment.from([item("Alpha", 0), item("Nested", 1), item("Omega", 0)]),
    1,
    1,
  );
};

/** The position of the cell whose first paragraph reads `text`. */
const cellAt = (doc: PMNode, text: string): number => {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found >= 0) return false;
    if (
      (node.type.name === "tableCell" || node.type.name === "tableHeader") &&
      node.firstChild?.textContent === text
    ) {
      found = pos;
      return false;
    }
    return true;
  });
  if (found < 0) throw new Error(`no cell reads ${text}`);
  return found;
};

const SELECTIONS = [
  ["two cells in a row", "Cell B3", "Cell D3"],
  ["a block of cells", "Cell B3", "Cell D4"],
] as const;

const blocksAfter = async (
  mode: EditorMode,
  anchor: string,
  head: string,
  resolution: "accept" | "reject" = "accept",
): Promise<unknown> => {
  const shape = DOCUMENT_SHAPES.find(({ id }) => id === "tables");
  if (!shape) throw new Error("no tables shape");
  const base = await parseShapeDocument(await shape.build());
  const view = new HeadlessEditorView(createHarnessState(base, mode));
  view.dispatch(
    view.state.tr.setSelection(
      CellSelection.create(
        view.state.doc,
        cellAt(view.state.doc, anchor),
        cellAt(view.state.doc, head),
      ),
    ),
  );
  // The host's cut: the suggesting deletion when tracking, else the editor's.
  if (!deleteSelectionAsSuggestion(view.state, view.dispatch))
    view.dispatch(view.state.tr.deleteSelection());
  view.paste(pastedList(view.state.doc));
  const saved = await saveHarnessState(view.state, base);
  const reviewer = await FolioDocxReviewer.fromBuffer(
    saved.bytes.buffer as ArrayBuffer,
  );
  if (resolution === "accept") reviewer.acceptAll();
  else reviewer.rejectAll();
  return summarizeReviewer(
    await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()),
  ).map(({ kind, text }) => ({ kind, text }));
};

describe("a list pasted after cutting table cells", () => {
  for (const [name, anchor, head] of SELECTIONS)
    test(`over ${name}: accepting the suggestions gives what editing gives`, async () => {
      expect(await blocksAfter("suggesting", anchor, head)).toEqual(
        await blocksAfter("editing", anchor, head),
      );
    });
  test("rejecting the cut and the paste gives back the table", async () => {
    const shape = DOCUMENT_SHAPES.find(({ id }) => id === "tables");
    if (!shape) throw new Error("no tables shape");
    const source = summarizeReviewer(
      await FolioDocxReviewer.fromBuffer(
        (await shape.build()).buffer as ArrayBuffer,
      ),
    ).map(({ kind, text }) => ({ kind, text }));
    for (const [, anchor, head] of SELECTIONS)
      expect(await blocksAfter("suggesting", anchor, head, "reject")).toEqual(
        source,
      );
  });
});
