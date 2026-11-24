import { panic } from "better-result";
import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import type { Document } from "../types/document";
import type { FolioDerivedFields } from "../types/fields";
import {
  createHiddenEditorManager,
  type HiddenEditorManagerDeps,
} from "./hiddenEditorManager";

const PREFIX = "EVIDENCE ";

/** Exhibits numbered by the order each source first appears. */
const exhibits: FolioDerivedFields = {
  owns: (instruction) => instruction.startsWith(PREFIX),
  displayTexts: (instructions) => {
    const numbers = new Map<string, number>();
    return instructions.map((instruction) => {
      const number = numbers.get(instruction) ?? numbers.size + 1;
      numbers.set(instruction, number);
      return `Exhibit ${number}`;
    });
  },
};

const source = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [{ type: "run", content: [{ type: "text", text: "See " }] }],
        },
      ],
    },
  },
});

const mount = () => {
  const releaseDom = acquireHarnessDom();
  const host = document.createElement("div");
  document.body.append(host);
  const refusals: string[] = [];
  // What a host plugin's view saw last: the document the editor shows.
  let observed: PMNode | null = null;
  const observer = new Plugin({
    view: (view) => {
      observed = view.state.doc;
      return {
        update: (next) => {
          observed = next.state.doc;
        },
      };
    },
  });
  const manager = createHiddenEditorManager({
    getHost: () => host,
    getDocument: source,
    getDocumentContext: source,
    getDocumentIdentity: () => "fields",
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [observer],
    getReadOnly: () => false,
    getDerivedFields: () => exhibits,
    getEditingMode: () => "editing",
    getSuggestionAuthor: () => "Reviewer",
    onTransaction: () => {},
    onCanonicalCompositionLayout: () => {},
    getEmbeddedFontFamilyMap: () => undefined,
    onSelectionChange: () => {},
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => {},
    onEditorViewReady: () => {},
    onEditorViewDestroy: () => {},
    onRemoteSelectionsChange: () => {},
    onSessionRefusal: (message) => refusals.push(message),
  } satisfies HiddenEditorManagerDeps);
  manager.ensureView();
  const view = manager.getView() ?? panic("Expected mounted editor");
  const toEnd = () => {
    manager.api.setSelection(view.state.doc.content.size - 1);
  };
  /** What each field shows, in document order. */
  const shown = () => {
    const texts: string[] = [];
    view.state.doc.descendants((node) => {
      if (node.type.name === "field")
        texts.push(String(node.attrs["displayText"]));
      return true;
    });
    return texts;
  };
  return {
    api: manager.api,
    refusals,
    shown,
    observedShown: () => observed === view.state.doc,
    toEnd,
    destroy: () => {
      manager.destroyView();
      host.remove();
      releaseDom();
    },
  };
};

const insert = (instruction: string) => ({
  instruction: `${PREFIX}${instruction}`,
  displayText: "Exhibit ?",
  locked: true,
});

test("inserts fields and keeps their derived texts with each edit, undone together", () => {
  const editor = mount();
  try {
    editor.toEnd();
    expect(editor.api.insertField(insert("invoice"))).toBe(true);
    expect(editor.shown()).toEqual(["Exhibit 1"]);

    // A second source numbers after the first; the first again reuses 1.
    editor.toEnd();
    expect(editor.api.insertField(insert("contract"))).toBe(true);
    editor.toEnd();
    expect(editor.api.insertField(insert("invoice"))).toBe(true);
    expect(editor.shown()).toEqual(["Exhibit 1", "Exhibit 2", "Exhibit 1"]);

    // A new first source renumbers every field in the same step, and one
    // undo takes the insertion and the renumbering back together.
    editor.api.setSelection(1);
    expect(editor.api.insertField(insert("letter"))).toBe(true);
    expect(editor.shown()).toEqual([
      "Exhibit 1",
      "Exhibit 2",
      "Exhibit 3",
      "Exhibit 2",
    ]);
    expect(editor.api.undo()).toBe(true);
    expect(editor.shown()).toEqual(["Exhibit 1", "Exhibit 2", "Exhibit 1"]);
    expect(editor.api.redo()).toBe(true);
    expect(editor.shown()).toEqual([
      "Exhibit 1",
      "Exhibit 2",
      "Exhibit 3",
      "Exhibit 2",
    ]);
    expect(editor.observedShown()).toBe(true);
    expect(editor.refusals).toEqual([]);
  } finally {
    editor.destroy();
  }
});
