/**
 * Default-editor findings of #2300: input while suggesting over content that
 * already carries pending revisions reviews as the input made directly.
 * Accepting reads as the direct edit, and rejecting restores the source.
 */
import { expect, test } from "bun:test";
import { DOMParser } from "prosemirror-model";
import { AllSelection, TextSelection } from "prosemirror-state";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { flattenPastedHtmlLists } from "../prosemirror/extensions/features/pastedHtmlLists";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { shapeArrayBuffer } from "./documentShapes";
import { PASTED_TABLE } from "./editorCommandConformance";
import {
  createHarnessState,
  type EditorMode,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "./editorHarness";

const PASTED_PARAGRAPHS = "<p>First <strong>bold</strong></p><p>Second</p>";

const pasteHtml = (view: HeadlessEditorView, html: string) => {
  const host = document.createElement("div");
  host.innerHTML = flattenPastedHtmlLists(html);
  view.paste(DOMParser.fromSchema(view.state.schema).parseSlice(host));
};

const selectAll = (view: HeadlessEditorView) =>
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));

/** Where `text` starts in the document. */
const textPosition = (view: HeadlessEditorView, text: string) => {
  let found = -1;
  view.state.doc.descendants((node, pos) => {
    const offset = node.isText ? (node.text?.indexOf(text) ?? -1) : -1;
    if (found < 0 && offset >= 0) found = pos + offset;
    return found < 0;
  });
  if (found < 0) throw new Error(`no ${text} in the document`);
  return found;
};

const TRACES = {
  "a paste from inside a pasted list item into the next": {
    shape: "plain-markdown",
    run: (view: HeadlessEditorView) => {
      pasteHtml(
        view,
        "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
      );
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(
            view.state.doc,
            textPosition(view, "Alpha") + 1,
            textPosition(view, "Nested") + 2,
          ),
        ),
      );
      pasteHtml(view, PASTED_PARAGRAPHS);
    },
  },
  "a paste over everything from a pasted table": {
    shape: "notes",
    run: (view: HeadlessEditorView) => {
      view.paste(PASTED_TABLE(view.state.doc));
      selectAll(view);
      pasteHtml(view, PASTED_PARAGRAPHS);
    },
  },
  "Enter after typing into a paragraph whose break is struck": {
    shape: "fields-links-bookmarks",
    run: (view: HeadlessEditorView) => {
      selectAll(view);
      view.pressKey("Backspace");
      view.typeText("shall");
      view.pressKey("Enter");
      view.typeText("buyer");
    },
  },
} as const;

const blocks = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text, table, listReference }) => ({
    kind,
    text,
    inTable: table !== undefined,
    level: listReference?.level,
  }));

type Trace = (typeof TRACES)[keyof typeof TRACES];
const runTrace = async ({ shape, run }: Trace, mode: EditorMode) => {
  const base = await parseShapeDocument(
    new Uint8Array(await shapeArrayBuffer(shape)),
  );
  const state = createHarnessState(base, mode);
  const view = new HeadlessEditorView(
    state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1))),
  );
  const releaseDom = acquireHarnessDom();
  try {
    run(view);
  } finally {
    releaseDom();
  }
  const { bytes } = await saveHarnessState(view.state, base);
  return bytes;
};

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

for (const [name, trace] of Object.entries(TRACES))
  for (const session of SESSIONS)
    test(`${name} while suggesting reviews as the direct edit (${session} review)`, async () => {
      const read = (bytes: Uint8Array) =>
        FolioDocxReviewer.fromBuffer(bytes.slice().buffer, { session });
      const source = blocks(
        await read(new Uint8Array(await shapeArrayBuffer(trace.shape))),
      );
      const direct = blocks(await read(await runTrace(trace, "editing")));
      const suggested = await runTrace(trace, "suggesting");
      const accepting = await read(suggested);
      accepting.acceptAll();
      expect(
        blocks(await read(new Uint8Array(await accepting.toBuffer()))),
      ).toEqual(direct);
      const rejecting = await read(suggested);
      rejecting.rejectAll();
      expect(
        blocks(await read(new Uint8Array(await rejecting.toBuffer()))),
      ).toEqual(source);
    });
