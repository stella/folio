import { afterAll, afterEach, beforeAll, describe, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { Schema } from "prosemirror-model";
import { closeHistory, history, redo, undo } from "prosemirror-history";
import { EditorState, Plugin } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { createHistoryBridgePlugin } from "./historyBridge";
import type { HistoryChange } from "./historyBridge";

const schema = new Schema({
  nodes: {
    doc: { content: "paragraph+" },
    paragraph: { content: "text*", toDOM: () => ["p", 0] },
    text: {},
  },
});

/** Like folio's paraId allocator: a follow-up change kept out of history. */
const appendOutsideHistory = new Plugin({
  appendTransaction: (transactions, _old, state) =>
    transactions.some((transaction) => transaction.getMeta("append") === true)
      ? state.tr.insertText("+", 1).setMeta("addToHistory", false)
      : null,
});

const views: EditorView[] = [];

const createEditor = (depth = 100) => {
  const changes: HistoryChange[] = [];
  const plugins = [
    history({ depth }),
    appendOutsideHistory,
    createHistoryBridgePlugin((change) => changes.push(change)),
  ];
  const emptyState = () =>
    EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("paragraph")]),
      plugins,
    });
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, { state: emptyState() });
  views.push(view);
  const run = (command: Command) => command(view.state, view.dispatch);

  return {
    view,
    changes,
    /** Type into a new undo step, as after a pause. */
    typeStep: (text: string) => view.dispatch(closeHistory(view.state.tr.insertText(text))),
    /** Type into the open undo step, as while typing on. */
    typeOn: (text: string) => view.dispatch(view.state.tr.insertText(text)),
    undo: () => run(undo),
    redo: () => run(redo),
    reload: () => view.updateState(emptyState()),
  };
};

beforeAll(() => GlobalRegistrator.register());

afterEach(() => {
  for (const view of views.splice(0)) {
    const mount = view.dom.parentElement;
    view.destroy();
    mount?.remove();
  }
});

afterAll(() => GlobalRegistrator.unregister());

describe("createHistoryBridgePlugin", () => {
  test("reports one new step per undo step, and undo / redo apart", () => {
    const editor = createEditor();
    editor.typeStep("a");
    editor.typeOn("b");
    editor.typeStep("c");
    expect(editor.undo()).toBe(true);
    expect(editor.redo()).toBe(true);
    expect(editor.changes).toEqual(["newStep", "sameStep", "newStep", "undoRedo", "undoRedo"]);
  });

  test("a change another plugin appends joins the step that caused it", () => {
    const editor = createEditor();
    editor.view.dispatch(
      closeHistory(editor.view.state.tr.insertText("a").setMeta("append", true)),
    );
    expect(editor.view.state.doc.textContent).toBe("+a");
    expect(editor.changes).toEqual(["newStep"]);
  });

  test("a change kept out of history opens no step", () => {
    const editor = createEditor();
    editor.view.dispatch(editor.view.state.tr.insertText("a").setMeta("addToHistory", false));
    expect(editor.changes).toEqual(["sameStep"]);
  });

  test("steps the history keeps line up with the host's stack past the depth limit", () => {
    const depth = 3;
    const editor = createEditor(depth);
    const steps = 40;
    for (let step = 0; step < steps; step += 1) {
      editor.typeStep("x");
    }
    expect(editor.changes).toEqual(Array.from({ length: steps }, () => "newStep"));

    // The host's stack is `steps` long. The editor trims its own to `depth`
    // once it overflows, so it can walk back only part of it; an undo past
    // that finds nothing to undo and changes nothing.
    let undone = 0;
    while (editor.undo()) {
      undone += 1;
    }
    expect(undone).toBeGreaterThanOrEqual(depth);
    expect(undone).toBeLessThan(steps);
    expect(editor.changes.slice(steps)).toEqual(Array.from({ length: undone }, () => "undoRedo"));
  });

  test("loading a new document is not an edit", () => {
    const editor = createEditor();
    editor.typeStep("a");
    editor.reload();
    editor.typeStep("b");
    expect(editor.changes).toEqual(["newStep", "newStep"]);
  });
});
