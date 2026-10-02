import { afterAll, afterEach, beforeAll, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { panic } from "better-result";
import { TextSelection } from "prosemirror-state";
import type { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import { documentShape } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
  summarizeState,
} from "../../__tests__/editorHarness";

const views: EditorView[] = [];
const MAC = /Mac|iP(?:hone|[oa]d)/u.test(globalThis.navigator?.platform ?? "");

beforeAll(() => GlobalRegistrator.register());
afterEach(() => {
  for (const view of views.splice(0)) {
    const mount = view.dom.parentElement;
    view.destroy();
    mount?.remove();
  }
});
afterAll(() => GlobalRegistrator.unregister());

const mountEditor = (state: EditorState) => {
  const mount = document.createElement("div");
  document.body.append(mount);
  const view = new EditorView(mount, { state });
  views.push(view);
  view.focus();
  return view;
};

const KEYS = { enter: "Enter", undo: "z", redo: "Z" } as const;

const key = (view: EditorView, binding: keyof typeof KEYS) => {
  const event = new KeyboardEvent("keydown", {
    key: KEYS[binding],
    keyCode: binding === "enter" ? 13 : 90,
    ctrlKey: binding !== "enter" && !MAC,
    metaKey: binding !== "enter" && MAC,
    shiftKey: binding === "redo",
    bubbles: true,
    cancelable: true,
  });
  view.dom.dispatchEvent(event);
  expect(event.defaultPrevented).toBe(true);
};

for (const mode of ["editing", "suggesting"] as const) {
  test(`${mode}: native IME insertion and undo restore a document with an inline object`, async () => {
    const base = await parseShapeDocument(await documentShape("image").build());
    const view = mountEditor(createHarnessState(base, mode));
    const initial = view.state.doc;
    view.dom.dispatchEvent(new Event("compositionstart", { bubbles: true }));
    view.dispatch(view.state.tr.insertText("alpha").setMeta("composition", 101));
    view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
    await Promise.resolve();
    const composed = view.state.doc;
    key(view, "undo");
    expect(view.state.doc.eq(initial)).toBe(true);
    key(view, "redo");
    expect(view.state.doc.eq(composed)).toBe(true);
  });
}

// Exercise paragraph-start and interior edges in both selection directions.
// The previous Enter conformance cases did not cover whole-paragraph removal.
const selectionCases = ["mixed-lists", "bare-package"].flatMap((shape) =>
  ["forward", "backward"].flatMap((direction) =>
    [1, 2].flatMap((startOffset) =>
      [1, 2].map((endOffset) => ({ shape, direction, startOffset, endOffset })),
    ),
  ),
);

for (const { shape, direction, startOffset, endOffset } of selectionCases) {
  test(`selected ${shape} Enter survives undo/redo (${direction}, ${startOffset}, ${endOffset})`, async () => {
    const base = await parseShapeDocument(await documentShape(shape).build());
    const baseline = summarizeState(createHarnessState(base, "editing"));
    const outcomes: ReturnType<typeof summarizeState>[] = [];
    for (const mode of ["editing", "suggesting"] as const) {
      const view = mountEditor(createHarnessState(base, mode));
      const positions: number[] = [];
      view.state.doc.descendants((node, pos) => {
        if (
          node.type.name === "paragraph" &&
          (shape !== "mixed-lists" || node.attrs["numPr"] != null) &&
          positions.length < 2
        ) {
          positions.push(pos);
        }
      });
      const first = positions.at(0);
      const second = positions.at(1);
      if (first === undefined || second === undefined) panic("Missing selection paragraphs");
      const from = first + startOffset;
      const to = second + endOffset;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(
            view.state.doc,
            direction === "forward" ? from : to,
            direction === "forward" ? to : from,
          ),
        ),
      );
      const initial = view.state.doc;
      key(view, "enter");
      const entered = view.state.doc;
      key(view, "undo");
      expect(view.state.doc.eq(initial)).toBe(true);
      key(view, "redo");
      expect(view.state.doc.eq(entered)).toBe(true);
      const accepted = resolveAllChanges(view.state, "accept");
      outcomes.push(summarizeState(accepted));
      const saved = await saveHarnessState(accepted, base);
      expect(
        summarizeState(createHarnessState(await parseShapeDocument(saved.bytes), "editing")),
      ).toEqual(summarizeState(accepted));
      if (mode === "suggesting") {
        expect(summarizeState(resolveAllChanges(view.state, "reject"))).toEqual(baseline);
      }
    }
    expect(outcomes.at(1)).toEqual(outcomes.at(0));
  });
}
