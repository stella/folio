import { afterAll, afterEach, beforeAll, describe, expect, setSystemTime, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

// The schema singleton builds a starter kit of its own. Loading it first
// settles the import cycle between the two before the kit below is built.
import "../../schema";
import { ExtensionManager } from "../ExtensionManager";
import { createStarterKit } from "../StarterKit";
import type { HistoryShortcutOwner } from "./HistoryExtension";
import { createSuggestionModePlugin } from "../../plugins/suggestionMode";
import { resolveAllChangesInHeadlessState } from "../../commands/comments";

// Read before the DOM is registered, as ProseMirror's keymap reads it at import.
const MAC = /Mac|iP(?:hone|[oa]d)/u.test(globalThis.navigator?.platform ?? "");

type Press = { key: string; shift?: boolean };

const UNDO: Press = { key: "z" };
const REDO_KEYS: readonly Press[] = [{ key: "y" }, { key: "Z", shift: true }];

const views: EditorView[] = [];

type EditorOptions = { mode?: "editing" | "suggesting"; initialText?: string };

const createEditor = (
  historyShortcuts: HistoryShortcutOwner,
  { mode = "editing", initialText = "" }: EditorOptions = {},
) => {
  const manager = new ExtensionManager(createStarterKit({ historyShortcuts }));
  manager.buildSchema();
  manager.initializeRuntime();
  const schema = manager.getSchema();
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, {
    state: EditorState.create({
      schema,
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, initialText ? schema.text(initialText) : undefined),
      ]),
      plugins: [
        createSuggestionModePlugin(mode === "suggesting", "Reviewer"),
        ...manager.getPlugins(),
      ],
    }),
  });
  views.push(view);

  return {
    view,
    text: () => view.state.doc.textContent,
    type: (text: string) => view.dispatch(view.state.tr.insertText(text)),
    /**
     * Press a Mod chord in the editor and expose whether native history was
     * cancelled as well as the resulting document.
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
      return event;
    },
  };
};

beforeAll(() => GlobalRegistrator.register());

afterEach(() => {
  setSystemTime();
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

  for (const mode of ["editing", "suggesting"] as const) {
    test(`${mode}: empty history shortcuts suppress native history after composition`, async () => {
      for (const redo of REDO_KEYS) {
        const editor = createEditor("editor", { mode });
        editor.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
        editor.view.dispatch(editor.view.state.tr.insertText("alpha").setMeta("composition", 1));
        editor.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
        await Promise.resolve();
        const composed = editor.view.state.doc;
        const caret = editor.view.state.selection;
        expect(editor.press(redo).defaultPrevented).toBe(true);
        expect(editor.view.state.doc.eq(composed)).toBe(true);
        expect(editor.view.state.selection.eq(caret)).toBe(true);
        editor.press(UNDO);
        expect(editor.text()).toBe("");
        expect(editor.press(UNDO).defaultPrevented).toBe(true);
      }
    });
  }

  test("composition replacement, empty redo, and delete have the same accepted result", async () => {
    const results: string[] = [];
    for (const mode of ["editing", "suggesting"] as const) {
      const editor = createEditor("editor", { mode, initialText: "before target after" });
      editor.view.dispatch(
        editor.view.state.tr.setSelection(TextSelection.create(editor.view.state.doc, 8, 14)),
      );
      editor.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      editor.view.dispatch(editor.view.state.tr.insertText("alpha").setMeta("composition", 2));
      editor.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
      await Promise.resolve();
      const composed = editor.view.state.doc;
      const caret = editor.view.state.selection;
      expect(editor.press({ key: "Z", shift: true }).defaultPrevented).toBe(true);
      expect(editor.view.state.doc.eq(composed)).toBe(true);
      expect(editor.view.state.selection.eq(caret)).toBe(true);
      const event = new KeyboardEvent("keydown", {
        key: "Delete",
        keyCode: 46,
        bubbles: true,
        cancelable: true,
      });
      editor.view.dom.dispatchEvent(event);
      // Native editing performs the unhandled one-character deletion.
      if (!event.defaultPrevented) {
        const from = editor.view.state.selection.from;
        editor.view.dispatch(editor.view.state.tr.delete(from, from + 1));
      }
      results.push(resolveAllChangesInHeadlessState(editor.view.state, "accept").doc.textContent);
      editor.press(UNDO);
      expect(editor.view.state.doc.eq(composed)).toBe(true);
      editor.press(UNDO);
      expect(editor.text()).toBe("before target after");
      editor.press({ key: "Z", shift: true });
      expect(editor.view.state.doc.eq(composed)).toBe(true);
    }
    expect(results).toEqual(["before alphaafter", "before alphaafter"]);
  });

  test("a native commit pending at compositionend stays in the same undo event", async () => {
    const editor = createEditor("editor", { mode: "suggesting" });
    editor.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
    editor.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
    editor.view.dispatch(editor.view.state.tr.insertText("alpha").setMeta("composition", 3));
    await Promise.resolve();
    const composed = editor.view.state.doc;
    editor.press(UNDO);
    expect(editor.text()).toBe("");
    editor.press({ key: "Z", shift: true });
    expect(editor.view.state.doc.eq(composed)).toBe(true);
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

describe("HistoryExtension input grouping", () => {
  const input = (view: EditorView, inputType: string) => {
    view.dom.dispatchEvent(new InputEvent("beforeinput", { inputType, bubbles: true }));
  };

  // Mark-only deletes have empty step maps; replacement deletes have ranges.
  // Exercise both representations so grouping cannot depend on those maps.
  for (const representation of ["replacement", "mark"] as const) {
    for (const deleteInput of [
      "deleteContentBackward",
      "deleteContentForward",
      "deleteWordBackward",
    ]) {
      test(`${representation}: typing and ${deleteInput} undo separately`, () => {
        const editor = createEditor("editor");
        input(editor.view, "insertText");
        editor.type("alpha");
        const typed = editor.view.state.doc;
        input(editor.view, deleteInput);
        const tr = editor.view.state.tr;
        if (representation === "replacement") tr.delete(1, 2);
        else tr.addMark(1, 2, editor.view.state.schema.mark("bold"));
        editor.view.dispatch(tr);
        editor.press(UNDO);
        expect(editor.view.state.doc.eq(typed)).toBe(true);
        editor.press(UNDO);
        expect(editor.text()).toBe("");
      });
    }

    for (const direction of ["Backward", "Forward"] as const) {
      test(`${representation}: consecutive ${direction} deletions undo together`, () => {
        const editor = createEditor("editor");
        input(editor.view, "insertText");
        editor.type("alpha");
        const typed = editor.view.state.doc;
        for (let index = 0; index < 3; index++) {
          editor.view.dom.dispatchEvent(
            new KeyboardEvent("keydown", {
              key: direction === "Backward" ? "Backspace" : "Delete",
              bubbles: true,
            }),
          );
          input(editor.view, `deleteContent${direction}`);
          const tr = editor.view.state.tr;
          const from = direction === "Backward" ? 5 - index : 1;
          if (representation === "replacement") tr.delete(from, from + 1);
          else {
            const markFrom = direction === "Backward" ? from : 1 + index;
            tr.addMark(markFrom, markFrom + 1, editor.view.state.schema.mark("bold"));
          }
          editor.view.dispatch(tr);
        }
        editor.press(UNDO);
        expect(editor.view.state.doc.eq(typed)).toBe(true);
        editor.press(UNDO);
        expect(editor.text()).toBe("");
      });
    }

    test(`${representation}: a pause ends the deletion group`, () => {
      setSystemTime(new Date("2026-01-01T00:00:00Z"));
      const editor = createEditor("editor");
      input(editor.view, "insertText");
      editor.type("alpha");
      input(editor.view, "deleteContentForward");
      const first = editor.view.state.tr;
      if (representation === "replacement") first.delete(1, 2);
      else first.addMark(1, 2, editor.view.state.schema.mark("bold"));
      editor.view.dispatch(first);
      const firstDelete = editor.view.state.doc;
      setSystemTime(new Date("2026-01-01T00:00:01Z"));
      input(editor.view, "deleteContentForward");
      const second = editor.view.state.tr;
      if (representation === "replacement") second.delete(1, 2);
      else second.addMark(2, 3, editor.view.state.schema.mark("bold"));
      editor.view.dispatch(second);
      editor.press(UNDO);
      expect(editor.view.state.doc.eq(firstDelete)).toBe(true);
    });

    test(`${representation}: changing deletion direction starts a new event`, () => {
      const editor = createEditor("editor");
      input(editor.view, "insertText");
      editor.type("alpha");
      input(editor.view, "deleteContentBackward");
      const backward = editor.view.state.tr;
      if (representation === "replacement") backward.delete(5, 6);
      else backward.addMark(5, 6, editor.view.state.schema.mark("bold"));
      editor.view.dispatch(backward);
      const firstDelete = editor.view.state.doc;
      input(editor.view, "deleteContentForward");
      const forward = editor.view.state.tr;
      if (representation === "replacement") forward.delete(1, 2);
      else forward.addMark(1, 2, editor.view.state.schema.mark("bold"));
      editor.view.dispatch(forward);
      editor.press(UNDO);
      expect(editor.view.state.doc.eq(firstDelete)).toBe(true);
    });

    for (const direction of ["Backward", "Forward"] as const) {
      test(`${representation}: toolbar edits do not inherit ${direction} deletion grouping`, () => {
        const editor = createEditor("editor");
        input(editor.view, "insertText");
        editor.type("alpha");
        const typed = editor.view.state.doc;
        input(editor.view, `deleteContent${direction}`);
        const deletion = editor.view.state.tr;
        const from = direction === "Backward" ? 5 : 1;
        if (representation === "replacement") deletion.delete(from, from + 1);
        else deletion.addMark(from, from + 1, editor.view.state.schema.mark("bold"));
        editor.view.dispatch(deletion);
        const deleted = editor.view.state.doc;
        const toolbar = editor.view.state.tr.addMark(
          1,
          editor.view.state.doc.content.size - 1,
          editor.view.state.schema.mark("italic"),
        );
        editor.view.dispatch(toolbar);
        expect(toolbar.getMeta("composition")).toBeUndefined();
        const styled = editor.view.state.doc;
        input(editor.view, `deleteContent${direction}`);
        const nextDeletion = editor.view.state.tr;
        if (representation === "replacement") {
          const nextFrom = direction === "Backward" ? 4 : 1;
          nextDeletion.delete(nextFrom, nextFrom + 1);
        } else {
          const nextFrom = direction === "Backward" ? 4 : 2;
          nextDeletion.addMark(nextFrom, nextFrom + 1, editor.view.state.schema.mark("bold"));
        }
        editor.view.dispatch(nextDeletion);
        editor.press(UNDO);
        expect(editor.view.state.doc.eq(styled)).toBe(true);
        editor.press(UNDO);
        expect(editor.view.state.doc.eq(deleted)).toBe(true);
        editor.press(UNDO);
        expect(editor.view.state.doc.eq(typed)).toBe(true);
      });
    }
  }

  test("keyboard boundaries run before a key handler consumes deletion", () => {
    const editor = createEditor("editor");
    input(editor.view, "insertText");
    editor.type("alpha");
    const typed = editor.view.state.doc;
    editor.view.setProps({
      handleKeyDown(view, event) {
        if (event.key !== "Delete") return false;
        view.dispatch(view.state.tr.delete(1, 2));
        return true;
      },
    });
    editor.view.dom.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "Delete",
        bubbles: true,
        cancelable: true,
      }),
    );
    editor.press(UNDO);
    expect(editor.view.state.doc.eq(typed)).toBe(true);
  });

  test("each composition groups its updates and separates surrounding typing", () => {
    const editor = createEditor("editor");
    input(editor.view, "insertText");
    editor.type("a");
    for (const character of ["b", "c"]) {
      editor.view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
      input(editor.view, "insertCompositionText");
      editor.type(character);
      input(editor.view, "insertCompositionText");
      editor.type(character);
      editor.view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
    }
    input(editor.view, "insertText");
    editor.type("d");
    for (const text of ["abbcc", "abb", "a", ""]) {
      editor.press(UNDO);
      expect(editor.text()).toBe(text);
    }
  });

  test("continuous typing stays in one undo event", () => {
    const editor = createEditor("editor");
    for (const text of ["a", "l", "p", "h", "a"]) {
      input(editor.view, "insertText");
      editor.type(text);
    }
    editor.press(UNDO);
    expect(editor.text()).toBe("");
  });

  for (const eventType of ["paste", "drop"]) {
    test(`${eventType} separates edits on both sides`, () => {
      const editor = createEditor("editor");
      input(editor.view, "insertText");
      editor.type("a");
      editor.view.dom.dispatchEvent(new Event(eventType, { bubbles: true }));
      editor.type("b");
      input(editor.view, "insertText");
      editor.type("c");
      editor.press(UNDO);
      expect(editor.text()).toBe("ab");
      editor.press(UNDO);
      expect(editor.text()).toBe("a");
    });
  }
});
