import { describe, expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";
import type { EditorState } from "prosemirror-state";

import { shapeArrayBuffer } from "../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
} from "../../__tests__/editorHarness";
import { isInlineAnchor } from "../extensions/nodes/inlineAnchorDom";

const SHAPES = ["fields-links-bookmarks", "comments"] as const;

const anchorCount = (state: EditorState) => {
  let count = 0;
  state.doc.descendants((node) => {
    if (isInlineAnchor(node)) count += 1;
  });
  return count;
};

/** The anchors with text on the side the key reaches, in document order. */
const caretsBesideAnchors = (
  state: EditorState,
  direction: "backward" | "forward",
) => {
  const carets: { caret: number; name: string }[] = [];
  state.doc.descendants((node, pos) => {
    if (!isInlineAnchor(node)) return;
    const $before = state.doc.resolve(pos);
    const $after = state.doc.resolve(pos + node.nodeSize);
    if (direction === "forward" && $after.nodeAfter?.isText) {
      carets.push({ caret: pos, name: node.type.name });
    }
    if (direction === "backward" && $before.nodeBefore?.isText) {
      carets.push({ caret: pos + node.nodeSize, name: node.type.name });
    }
  });
  return carets;
};

for (const mode of ["editing", "suggesting"] as const) {
  describe(`a key beside an inline anchor while ${mode}`, () => {
    for (const shape of SHAPES) {
      for (const [key, direction] of [
        ["Delete", "forward"],
        ["Backspace", "backward"],
      ] as const) {
        test(`${key} keeps every anchor in ${shape} and reaches the text`, async () => {
          const document = await parseShapeDocument(
            new Uint8Array(await shapeArrayBuffer(shape)),
          );
          const state = createHarnessState(document, mode);
          const carets = caretsBesideAnchors(state, direction);
          expect(carets.length).toBeGreaterThan(0);
          for (const { caret } of carets) {
            const view = new HeadlessEditorView(
              state.apply(
                state.tr.setSelection(TextSelection.create(state.doc, caret)),
              ),
            );
            expect(view.pressKey(key)).toBe(true);
            expect(anchorCount(view.state)).toBe(anchorCount(state));
            if (mode === "editing") {
              expect(view.state.doc.textContent.length).toBe(
                state.doc.textContent.length - 1,
              );
            } else {
              expect(view.state.doc.textContent).toBe(state.doc.textContent);
              let deleted = 0;
              view.state.doc.descendants((node) => {
                if (
                  node.isText &&
                  node.marks.some(({ type }) => type.name === "deletion")
                ) {
                  deleted += node.nodeSize;
                }
              });
              expect(deleted).toBeGreaterThan(0);
            }
          }
        });
      }
    }
  });
}
