import { afterAll, afterEach, beforeAll, describe, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

// The schema singleton builds a starter kit of its own. Loading it first
// settles the import cycle between the two before the kit below is built.
import "../../schema";
import { ExtensionManager } from "../ExtensionManager";
import { createStarterKit } from "../StarterKit";
import type { HistoryShortcutOwner } from "./HistoryExtension";

// Read before the DOM is registered, as ProseMirror's keymap reads it at import.
const MAC = /Mac|iP(?:hone|[oa]d)/u.test(globalThis.navigator?.platform ?? "");

type Press = { key: string; shift?: boolean };

const UNDO: Press = { key: "z" };
const REDO_KEYS: readonly Press[] = [{ key: "y" }, { key: "Z", shift: true }];

const views: EditorView[] = [];

const createEditor = (historyShortcuts: HistoryShortcutOwner) => {
  const manager = new ExtensionManager(createStarterKit({ historyShortcuts }));
  manager.buildSchema();
  manager.initializeRuntime();
  const schema = manager.getSchema();
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, {
    state: EditorState.create({
      schema,
      doc: schema.node("doc", null, [schema.node("paragraph")]),
      plugins: manager.getPlugins(),
    }),
  });
  views.push(view);

  return {
    text: () => view.state.doc.textContent,
    type: (text: string) => view.dispatch(view.state.tr.insertText(text)),
    /**
     * Press a Mod chord in the editor. ProseMirror cancels Mod-y and Mod-z
     * whatever binds them, to keep the browser's own undo out, so the document
     * is what tells whether anything ran.
     */
    press: ({ key, shift = false }: Press) => {
      const event = new KeyboardEvent("keydown", {
        key,
        // The keymap falls back to the key code when Shift changes `key`.
        keyCode: key.toUpperCase().charCodeAt(0),
        metaKey: MAC,
        ctrlKey: !MAC,
        shiftKey: shift,
        bubbles: true,
        cancelable: true,
      });
      view.dom.dispatchEvent(event);
    },
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

describe("HistoryExtension shortcuts", () => {
  test("the editor answers undo and redo by default", () => {
    for (const redo of REDO_KEYS) {
      const editor = createEditor("editor");
      editor.type("typed");
      editor.press(UNDO);
      expect(editor.text()).toBe("");
      editor.press(redo);
      expect(editor.text()).toBe("typed");
    }
  });

  test("a host that owns the history keys gets every press untouched", () => {
    const editor = createEditor("host");
    editor.type("typed");
    for (const press of [UNDO, ...REDO_KEYS]) {
      editor.press(press);
      expect(editor.text()).toBe("typed");
    }
  });
});
