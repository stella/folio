import { describe, expect, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { schema } from "../schema";
import {
  createSuggestionModePlugin,
  handleSuggestionPaste,
  suggestionModeKey,
} from "./suggestionMode";

type AtomFixture = {
  name: string;
  create: () => PMNode;
};

const ATOM_FIXTURES = [
  { name: "hard break", create: () => schema.node("hardBreak") },
  { name: "tab", create: () => schema.node("tab") },
] as const satisfies readonly AtomFixture[];

const marked = (doc: PMNode, nodeType: string, markType: string): boolean => {
  let found = false;
  doc.descendants((node) => {
    if (node.type.name === nodeType && node.marks.some(({ type }) => type.name === markType)) {
      found = true;
    }
    return !found;
  });
  return found;
};

const documentWithAtom = (atom: PMNode): PMNode =>
  schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text("a"), atom, schema.text("b")]),
  ]);

type FakeView = {
  state: EditorState;
  dispatch: (tr: Transaction) => void;
};

const fakeView = (state: EditorState): FakeView => {
  const view: FakeView = {
    state,
    dispatch(tr) {
      view.state = view.state.apply(tr);
    },
  };
  return view;
};

const editorView = (view: FakeView): EditorView => {
  // SAFETY: suggestion-mode handlers exercised here read only state and dispatch.
  return view as unknown as EditorView;
};

const backspaceEvent = (): KeyboardEvent => {
  // SAFETY: the handler branches only on KeyboardEvent.key for Backspace.
  return { key: "Backspace" } as KeyboardEvent;
};

describe("suggestion mode uses tracked-run atom ownership", () => {
  for (const { name, create } of ATOM_FIXTURES) {
    test(`catch-all marks an inserted ${name}`, () => {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      let state = EditorState.create({
        doc: schema.node("doc", null, [schema.node("paragraph", null, [schema.text("ab")])]),
        plugins: [plugin],
      });

      state = state.apply(state.tr.insert(2, create()));

      expect(marked(state.doc, create().type.name, "insertion")).toBe(true);
    });

    test(`paste marks an inserted ${name}`, () => {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      let state = EditorState.create({
        doc: schema.node("doc", null, [schema.node("paragraph", null, [schema.text("x")])]),
        plugins: [plugin],
      });
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 2)));
      const view = fakeView(state);
      const pluginState = suggestionModeKey.getState(state);
      expect(pluginState).toBeDefined();
      if (!pluginState) {
        return;
      }

      expect(
        handleSuggestionPaste(
          editorView(view),
          new Slice(Fragment.from(create()), 0, 0),
          pluginState,
        ),
      ).toBe(true);

      expect(marked(view.state.doc, create().type.name, "insertion")).toBe(true);
    });

    test(`selection deletion marks a ${name}`, () => {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      let state = EditorState.create({ doc: documentWithAtom(create()), plugins: [plugin] });
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 3)));
      const view = fakeView(state);

      const handled = plugin.props.handleKeyDown?.call(plugin, editorView(view), backspaceEvent());

      expect(handled).toBe(true);
      expect(marked(view.state.doc, create().type.name, "deletion")).toBe(true);
    });

    test(`single deletion marks a ${name}`, () => {
      const plugin = createSuggestionModePlugin(true, "Reviewer");
      let state = EditorState.create({ doc: documentWithAtom(create()), plugins: [plugin] });
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
      const view = fakeView(state);

      const handled = plugin.props.handleKeyDown?.call(plugin, editorView(view), backspaceEvent());

      expect(handled).toBe(true);
      expect(marked(view.state.doc, create().type.name, "deletion")).toBe(true);
    });
  }
});
