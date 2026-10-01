import { expect, test } from "bun:test";
import { panic } from "better-result";
import { Schema } from "prosemirror-model";
import { EditorState, Plugin, NodeSelection, TextSelection } from "prosemirror-state";

import { documentShape } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
  summarizeState,
} from "../../__tests__/editorHarness";
import { FolioDocxReviewer } from "../../ai-edits/headless";
import { schema } from "../schema";
import { createDocumentStylesPlugin } from "./documentStyles";
import { createSuggestionModePlugin, deleteSelectionAsSuggestion } from "./suggestionMode";

const project = (state: EditorState) => {
  let images = 0;
  state.doc.descendants((node) => {
    if (node.type.name === "image") images += 1;
  });
  return { blocks: summarizeState(state), images };
};

// The input conformance gate's Enter targets did not cross paragraph boundaries
// or inline objects. Exercise those ranges in both selection directions.
for (const kind of ["node", "inline-range", "cross-paragraph", "paragraph-start"] as const) {
  for (const direction of ["forward", "backward"] as const) {
    test(`Enter and cut resolve equally for ${direction} ${kind} selections`, async () => {
      const base = await parseShapeDocument(await documentShape("image").build());
      const run = (mode: "suggesting" | "editing") => {
        let state = createHarnessState(base, mode);
        const positions = { image: -1, tail: -1 };
        state.doc.descendants((node, pos) => {
          if (node.type.name === "image") {
            positions.image = pos;
          }
          if (node.isTextblock && node.textContent === "Tail.") positions.tail = pos + 1;
        });
        if (positions.image < 0 || positions.tail < 0) panic("Missing selection fixture positions");
        const from =
          kind === "paragraph-start" ? state.doc.resolve(positions.image).start() : positions.image;
        const to =
          kind === "cross-paragraph" || kind === "paragraph-start"
            ? positions.tail + 2
            : positions.image + 3;
        const selection =
          kind === "node"
            ? NodeSelection.create(state.doc, positions.image)
            : TextSelection.create(
                state.doc,
                direction === "forward" ? from : to,
                direction === "forward" ? to : from,
              );
        state = state.apply(state.tr.setSelection(selection));
        const view = new HeadlessEditorView(state);
        expect(view.pressKey("Enter")).toBe(true);
        expect(view.transactions).toHaveLength(1);
        const undoView = new HeadlessEditorView(view.state);
        expect(undoView.pressKey("Mod-z")).toBe(true);
        expect(project(undoView.state)).toEqual(project(state));
        if (mode === "suggesting") {
          deleteSelectionAsSuggestion(view.state, view.dispatch);
        } else if (!view.state.selection.empty) {
          view.dispatch(view.state.tr.deleteSelection());
        }
        return view.state;
      };
      const editing = run("editing");
      const suggesting = run("suggesting");
      const baseline = project(createHarnessState(base, "editing"));
      expect(project(resolveAllChanges(suggesting, "accept"))).toEqual(project(editing));
      expect(project(resolveAllChanges(suggesting, "reject"))).toEqual(baseline);
      const saved = await saveHarnessState(suggesting, base);
      for (const resolution of ["accept", "reject"] as const) {
        const reviewer = await FolioDocxReviewer.fromBuffer(saved.bytes.slice().buffer);
        if (resolution === "accept") reviewer.acceptAll();
        else reviewer.rejectAll();
        const resolved = await parseShapeDocument(new Uint8Array(await reviewer.toBuffer()));
        expect(project(createHarnessState(resolved, "editing"))).toEqual(
          resolution === "accept" ? project(editing) : baseline,
        );
      }
    });
  }
}

test("Enter replacement preserves the live next style and invokes host hooks once", () => {
  let appends = 0;
  let filters = 0;
  const doc = schema.node("doc", null, [
    schema.node("paragraph", { styleId: "Title" }, [schema.text("Title")]),
  ]);
  let state = EditorState.create({
    doc,
    plugins: [
      createSuggestionModePlugin(true, "Reviewer"),
      createDocumentStylesPlugin({
        styles: [
          { styleId: "Title", type: "paragraph", next: "Body" },
          { styleId: "Body", type: "paragraph" },
        ],
      }),
      new Plugin({
        appendTransaction: () => {
          appends += 1;
          return null;
        },
        filterTransaction: () => {
          filters += 1;
          return true;
        },
      }),
    ],
  });
  state = state.apply(state.tr.setSelection(TextSelection.create(doc, 2, 6)));
  appends = 0;
  filters = 0;
  const view = new HeadlessEditorView(state);
  expect(view.pressKey("Enter")).toBe(true);
  expect(appends).toBe(1);
  expect(filters).toBe(1);
  expect(view.transactions).toHaveLength(1);
  expect(view.state.selection.$from.parent.attrs["styleId"]).toBe("Body");
  expect(resolveAllChanges(view.state, "reject").doc.textContent).toBe(doc.textContent);
});

test("an unsplittable selection remains unchanged", () => {
  const constrainedSchema = new Schema({
    nodes: {
      doc: { content: "paragraph" },
      paragraph: { content: "text*", attrs: { pPrMark: { default: null } } },
      text: {},
    },
    marks: {
      insertion: { attrs: { revisionId: {}, author: {}, date: {} } },
      deletion: { attrs: { revisionId: {}, author: {}, date: {} } },
    },
  });
  const doc = constrainedSchema.node("doc", null, [
    constrainedSchema.node("paragraph", null, [constrainedSchema.text("abc")]),
  ]);
  let state = EditorState.create({ doc, plugins: [createSuggestionModePlugin(true, "Reviewer")] });
  state = state.apply(state.tr.setSelection(TextSelection.create(doc, 1, 3)));
  const view = new HeadlessEditorView(state);
  expect(view.pressKey("Enter")).toBe(false);
  expect(view.transactions).toHaveLength(0);
  expect(view.state).toBe(state);
});
