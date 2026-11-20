import { expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { AllSelection, Selection } from "prosemirror-state";

import { shapeArrayBuffer } from "../../__tests__/documentShapes";
import { PASTED_TABLE } from "../../__tests__/editorCommandConformance";
import {
  createHarnessState,
  type EditorMode,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../../__tests__/editorHarness";
import { FolioDocxReviewer } from "../../ai-edits/headless";

type Edit = (view: HeadlessEditorView) => void;

const blocks = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text, displayLabel }) => ({
    kind,
    text,
    displayLabel,
  }));

const editIn = async (shape: string, mode: EditorMode, edits: Edit[]) => {
  const base = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer(shape)),
  );
  const view = new HeadlessEditorView(createHarnessState(base, mode));
  for (const edit of edits) edit(view);
  const { bytes } = await saveHarnessState(view.state, base);
  return bytes;
};

const atDocumentEnd: Edit = (view) => {
  view.dispatch(view.state.tr.setSelection(Selection.atEnd(view.state.doc)));
};
const selectAll: Edit = (view) => {
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
};
const pastePlain: Edit = (view) => {
  view.paste(new Slice(Fragment.from(view.state.schema.text("shall é")), 0, 0));
};
const pasteTable: Edit = (view) => {
  view.paste(PASTED_TABLE(view.state.doc));
};

// The default editor's browser finding of #2302: a table pasted at the end of
// the document's final paragraph, which leaves it ending with the table.
test.each([
  ["comments", "at the end", [atDocumentEnd, pasteTable]],
  ["single-decimal-list", "at the end", [atDocumentEnd, pasteTable]],
  [
    "comments",
    "after text over everything",
    [selectAll, pastePlain, pasteTable],
  ],
  ["notes", "after text over everything", [selectAll, pastePlain, pasteTable]],
] as const)(
  "a table pasted in %s %s while suggesting reviews as the direct paste",
  async (shape, _placement, edits) => {
    const source = await FolioDocxReviewer.fromBuffer(
      await shapeArrayBuffer(shape),
      { session: "legacy" },
    );
    const direct = await FolioDocxReviewer.fromBuffer(
      (await editIn(shape, "editing", [...edits])).slice().buffer,
      { session: "legacy" },
    );
    const suggested = await editIn(shape, "suggesting", [...edits]);
    const accepting = await FolioDocxReviewer.fromBuffer(
      suggested.slice().buffer,
      { session: "legacy" },
    );
    accepting.acceptAll();
    expect(blocks(accepting)).toEqual(blocks(direct));
    const rejecting = await FolioDocxReviewer.fromBuffer(
      suggested.slice().buffer,
      { session: "legacy" },
    );
    rejecting.rejectAll();
    expect(blocks(rejecting)).toEqual(blocks(source));
  },
);
