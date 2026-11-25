import { panic } from "better-result";
import { expect, test } from "bun:test";
import { AllSelection } from "prosemirror-state";

import { shapeArrayBuffer } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  type EditorMode,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../../__tests__/editorHarness";
import { suggestionModeKey } from "../../__tests__/reference-editor/suggestionMode";
import { FolioDocxReviewer } from "../../ai-edits/headless";

const blocks = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text }) => ({ kind, text }));

/** Select everything, then commit a composition over it as the browser does. */
const composeOverEverything = async (
  shape: string,
  mode: EditorMode,
  thenBackspace: boolean,
) => {
  const base = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer(shape)),
  );
  const view = new HeadlessEditorView(createHarnessState(base, mode));
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
  const plugin =
    suggestionModeKey.get(view.state) ?? panic("Missing suggestion plugin");
  plugin.props.handleDOMEvents?.["compositionstart"]?.call(
    plugin,
    view as never,
    new Event("compositionstart"),
  );
  // The browser's native replacement of the selection with the composed text.
  view.dispatch(view.state.tr.insertText("alpha"));
  plugin.props.handleDOMEvents?.["compositionend"]?.call(
    plugin,
    view as never,
    new Event("compositionend"),
  );
  await Promise.resolve();
  await Promise.resolve();
  if (thenBackspace) view.pressKey("Backspace");
  const { bytes } = await saveHarnessState(view.state, base);
  return bytes;
};

// The default editor's browser findings of #2302: a composition over
// everything, alone or followed by Backspace.
test.each([
  ["rtl-cjk", false],
  ["tables", true],
  ["comments", false],
  ["mixed-lists", false],
] as const)(
  "a suggested composition over everything in %s reviews as the direct edit (Backspace after: %p)",
  async (shape, thenBackspace) => {
    const source = await FolioDocxReviewer.fromBuffer(
      await shapeArrayBuffer(shape),
    );
    const direct = await FolioDocxReviewer.fromBuffer(
      (await composeOverEverything(shape, "editing", thenBackspace)).slice()
        .buffer,
    );
    const suggested = await composeOverEverything(
      shape,
      "suggesting",
      thenBackspace,
    );
    const accepting = await FolioDocxReviewer.fromBuffer(
      suggested.slice().buffer,
    );
    accepting.acceptAll();
    expect(blocks(accepting)).toEqual(blocks(direct));
    const rejecting = await FolioDocxReviewer.fromBuffer(
      suggested.slice().buffer,
    );
    rejecting.rejectAll();
    expect(blocks(rejecting)).toEqual(blocks(source));
  },
);
