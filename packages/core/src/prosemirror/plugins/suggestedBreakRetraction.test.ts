import { describe, expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { AllSelection, TextSelection } from "prosemirror-state";

import { shapeArrayBuffer } from "../../__tests__/documentShapes";
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
  reviewer.snapshot().blocks.map(({ kind, text }) => ({ kind, text }));

const editIn = async (shape: string, mode: EditorMode, edits: Edit[]) => {
  const base = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer(shape)),
  );
  const view = new HeadlessEditorView(createHarnessState(base, mode));
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)),
  );
  for (const edit of edits) edit(view);
  const { bytes } = await saveHarnessState(view.state, base);
  return bytes;
};

/** Accepting the suggested edits reads as the direct edit; rejecting them, as the source. */
const expectReviewEquivalence = async (shape: string, edits: Edit[]) => {
  const source = await FolioDocxReviewer.fromBuffer(
    await shapeArrayBuffer(shape),
    { session: "legacy" },
  );
  const direct = await FolioDocxReviewer.fromBuffer(
    (await editIn(shape, "editing", edits)).slice().buffer,
    { session: "legacy" },
  );
  const suggested = await editIn(shape, "suggesting", edits);
  const accepting = await FolioDocxReviewer.fromBuffer(
    suggested.slice().buffer,
    { session: "legacy" },
  );
  accepting.acceptAll();
  const rejecting = await FolioDocxReviewer.fromBuffer(
    suggested.slice().buffer,
    { session: "legacy" },
  );
  rejecting.rejectAll();
  expect(blocks(accepting)).toEqual(blocks(direct));
  expect(blocks(rejecting)).toEqual(blocks(source));
};

const typeTitleAndBreak: Edit = (view) => {
  view.typeText("Title");
  expect(view.pressKey("Enter")).toBe(true);
};

/** From inside the first paragraph to inside the second. */
const selectAcrossBreak: Edit = (view) => {
  const second = view.state.doc.child(0).nodeSize;
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.create(view.state.doc, 3, second + 5),
    ),
  );
};

const pasteTwoParagraphs: Edit = (view) => {
  const { schema } = view.state;
  view.paste(
    new Slice(
      Fragment.from([
        schema.node("paragraph", null, schema.text("Title")),
        schema.node("paragraph", null, schema.text("Body")),
      ]),
      1,
      1,
    ),
  );
};

const selectAll: Edit = (view) => {
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
};

describe("a deletion over a break the current author suggested", () => {
  for (const key of ["Delete", "Backspace"])
    test(`${key} over the break withdraws it`, async () => {
      await expectReviewEquivalence("plain-markdown", [
        typeTitleAndBreak,
        selectAcrossBreak,
        (view) => {
          expect(view.pressKey(key)).toBe(true);
        },
      ]);
    });

  test("typing over the break withdraws it", async () => {
    await expectReviewEquivalence("plain-markdown", [
      typeTitleAndBreak,
      selectAcrossBreak,
      (view) => view.typeText("x"),
    ]);
  });

  // The default editor's browser finding of #2300: pasted paragraphs, then
  // everything typed over.
  for (const shape of ["comments", "notes", "invisible-markers"])
    test(`typing over everything after a paragraph paste in ${shape}`, async () => {
      await expectReviewEquivalence(shape, [
        pasteTwoParagraphs,
        selectAll,
        (view) => view.typeText("alpha"),
      ]);
    });
});
