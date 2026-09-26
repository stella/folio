import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import { ContextMenuManager } from "./ContextMenuManager";

const stateWithCaretIn = (markdown: string, text: string): EditorState => {
  const document = fromMarkdown(markdown);
  const state = EditorState.create({
    doc: toProseDoc(document),
    plugins: [createDocumentNumberingPlugin(document.package.numbering)],
  });
  let caret: number | null = null;
  state.doc.descendants((node, pos) => {
    if (caret === null && node.type.name === "paragraph" && node.textContent === text) {
      caret = pos + 1;
    }
    return caret === null;
  });
  if (caret === null) {
    throw new Error(`no paragraph reads "${text}"`);
  }
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, caret)));
};

const listNumberingAt = (markdown: string, text: string) => {
  const manager = new ContextMenuManager();
  manager.openMenu(
    { x: 0, y: 0 },
    { state: stateWithCaretIn(markdown, text), cursorInTrackedChange: false },
  );
  return manager.getSnapshot().listNumbering;
};

describe("ContextMenuManager list numbering entries", () => {
  test("are absent outside a list", () => {
    expect(listNumberingAt("Plain.\n\n1. Alpha", "Plain.")).toEqual({ type: "none" });
  });

  test("offer continuing only when an earlier list of the kind exists", () => {
    expect(listNumberingAt("1. Alpha\n2. Beta", "Beta")).toEqual({
      type: "listItem",
      canContinue: false,
    });
    expect(listNumberingAt("1. Alpha\n\nProse.\n\n1. Gamma", "Gamma")).toEqual({
      type: "listItem",
      canContinue: true,
    });
  });
});
