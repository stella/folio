import { expect, test } from "bun:test";
import { Slice } from "prosemirror-model";
import { AllSelection } from "prosemirror-state";

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

const blocks = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text }) => ({ kind, text }));

const pasteTableOverEverythingAndType = async (
  shape: string,
  mode: EditorMode,
  edges: "closed" | "open",
) => {
  const base = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer(shape)),
  );
  const view = new HeadlessEditorView(createHarnessState(base, mode));
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
  const table = PASTED_TABLE(view.state.doc);
  // A clipboard table arrives closed, or open into its first and last cells.
  view.paste(edges === "closed" ? table : Slice.maxOpen(table.content));
  view.typeText("alpha");
  const { bytes } = await saveHarnessState(view.state, base);
  return bytes;
};

// The default editor's browser finding of #2300: a table pasted over
// everything, then typing, lands in the pasted table's last cell.
test.each([
  ["invisible-markers", "closed"],
  ["tables", "closed"],
  ["comments", "open"],
  ["plain-markdown", "open"],
] as const)(
  "typing after a suggested table paste over everything in %s (%s edges) reviews as the direct edit",
  async (shape, edges) => {
    const source = await FolioDocxReviewer.fromBuffer(
      await shapeArrayBuffer(shape),
      { session: "legacy" },
    );
    const direct = await FolioDocxReviewer.fromBuffer(
      (await pasteTableOverEverythingAndType(shape, "editing", edges)).slice()
        .buffer,
      { session: "legacy" },
    );
    const suggested = await pasteTableOverEverythingAndType(
      shape,
      "suggesting",
      edges,
    );
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
