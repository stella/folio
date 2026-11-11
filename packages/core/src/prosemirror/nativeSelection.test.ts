/**
 * After an update that changes the document, the browser caret sits in the
 * text before the editor's caret, set again by the editor, so an input
 * method composes into the run the caret follows. With no text beside the
 * caret, it sits right after the separator image before an inline anchor.
 */
import { afterAll, afterEach, beforeAll, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import { EditorState, type Plugin, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { createNativeSelectionRefreshPlugin } from "./nativeSelection";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "inline*", group: "block", toDOM: () => ["p", 0] },
    text: { group: "inline" },
    // An inline node that shows nothing, drawn the way bookmark boundaries are.
    anchor: {
      inline: true,
      group: "inline",
      atom: true,
      selectable: false,
      toDOM: () => ["span", { contenteditable: "false" }],
    },
  },
  marks: { strong: { toDOM: () => ["strong", 0] } },
});

const views: EditorView[] = [];
let releaseDom: ReturnType<typeof acquireHarnessDom>;
beforeAll(() => {
  releaseDom = acquireHarnessDom();
});
afterEach(() => {
  for (const view of views.splice(0)) {
    const mount = view.dom.parentElement;
    view.destroy();
    mount?.remove();
  }
});
afterAll(() => releaseDom());

/** "First" in bold, then "Second", with the caret between them. */
const rig = (plugins: readonly Plugin[]) => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text("First", [schema.mark("strong")]),
      schema.text("Second"),
    ]),
    schema.node("paragraph", null, schema.text("Other")),
  ]);
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, {
    state: EditorState.create({
      doc,
      plugins: [...plugins],
      selection: TextSelection.create(doc, 6),
    }),
  });
  views.push(view);
  view.focus();
  const second = view.domAtPos(6, 1);
  const native = document.getSelection();
  if (!native || second.node.nodeType !== second.node.TEXT_NODE)
    throw new TypeError("The rig needs a caret at the second run's start.");
  // The caret the browser holds is the start of the next run, which reads
  // as the same position as the end of the bold one.
  native.removeAllRanges();
  native.collapse(second.node, second.offset);
  return { view, native };
};

/** Change the document away from the caret, keeping the selection. */
const editElsewhere = (view: EditorView) =>
  view.updateState(view.state.apply(view.state.tr.insertText("!", 15)));

test("an update that changes the document puts the caret in the text before it", () => {
  const { view, native } = rig([createNativeSelectionRefreshPlugin()]);
  editElsewhere(view);
  expect(view.state.selection.from).toBe(6);
  expect(native.anchorNode?.textContent).toBe("First");
  expect(native.anchorOffset).toBe(5);
});

test("without the plugin the browser keeps the caret where it was", () => {
  const { view, native } = rig([]);
  editElsewhere(view);
  expect(native.anchorNode?.textContent).toBe("Second");
  expect(native.anchorOffset).toBe(0);
});

test("an update that changes only the selection leaves the browser's caret alone", () => {
  const { view, native } = rig([createNativeSelectionRefreshPlugin()]);
  view.updateState(view.state.apply(view.state.tr.setMeta("noop", true)));
  expect(native.anchorNode?.textContent).toBe("Second");
  expect(native.anchorOffset).toBe(0);
});

/**
 * An anchor and then "Text", with the caret before the anchor, after a
 * paragraph "First". The browser holds the caret where ProseMirror puts it:
 * before the separator image it draws ahead of the anchor.
 */
const anchorRig = () => {
  const anchored = schema.node("paragraph", null, [
    schema.node("anchor"),
    schema.text("Text"),
  ]);
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, schema.text("First")),
    anchored,
  ]);
  const caret = 8;
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, {
    state: EditorState.create({
      doc,
      plugins: [createNativeSelectionRefreshPlugin()],
      selection: TextSelection.create(doc, caret),
    }),
  });
  views.push(view);
  view.focus();
  const paragraph = view.dom.children.item(1);
  const native = document.getSelection();
  if (!native || !paragraph)
    throw new TypeError("The rig needs the anchored paragraph.");
  native.removeAllRanges();
  native.collapse(paragraph, 0);
  return { view, native, caret };
};

test("a caret with no text beside it goes after the separator before the anchor", () => {
  const { view, native, caret } = anchorRig();
  // A paste before the paragraph moves it, as a pasted list does.
  view.updateState(view.state.apply(view.state.tr.insertText("!", 1)));
  expect(view.state.selection.from).toBe(caret + 1);
  const { anchorNode, anchorOffset } = native;
  const before = anchorNode?.childNodes.item(anchorOffset - 1);
  const after = anchorNode?.childNodes.item(anchorOffset);
  expect(anchorNode?.nodeName).toBe("P");
  expect(before?.nodeName).toBe("IMG");
  expect(before instanceof Element && before.className).toBe(
    "ProseMirror-separator",
  );
  expect(
    after instanceof Element && after.getAttribute("contenteditable"),
  ).toBe("false");
  // The browser's caret still reads as the editor's.
  if (anchorNode)
    expect(view.posAtDOM(anchorNode, anchorOffset)).toBe(caret + 1);
});
