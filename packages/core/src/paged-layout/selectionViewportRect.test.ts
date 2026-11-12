import { expect, test } from "bun:test";
import type { EditorView } from "prosemirror-view";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import {
  FOLIO_CARET_RECT_ATTRIBUTE,
  FOLIO_INPUT_HIDDEN_ATTRIBUTE,
  getFolioCaretViewportRect,
} from "./selectionViewportRect";

/** A scroll root holding the hidden editor and an overlay with a painted caret. */
const withCaret = (
  check: (view: EditorView, overlay: HTMLElement) => void,
): void => {
  const releaseDom = acquireHarnessDom();
  try {
    const root = document.createElement("div");
    root.setAttribute("data-folio-scroll", "");
    const editor = document.createElement("div");
    const overlay = document.createElement("div");
    const caret = document.createElement("div");
    caret.setAttribute(FOLIO_CARET_RECT_ATTRIBUTE, "");
    caret.getBoundingClientRect = () => new DOMRect(40, 60, 2, 18);
    overlay.append(caret);
    root.append(editor, overlay);
    check({ dom: editor } as unknown as EditorView, overlay);
  } finally {
    releaseDom();
  }
};

test("the painted caret anchors floating UI", () => {
  withCaret((view) => {
    expect(getFolioCaretViewportRect(view)).toEqual(new DOMRect(40, 60, 2, 18));
  });
});

test("a caret the overlay hides while typed text lays out anchors nothing", () => {
  withCaret((view, overlay) => {
    overlay.setAttribute(FOLIO_INPUT_HIDDEN_ATTRIBUTE, "");
    expect(getFolioCaretViewportRect(view)).toBeNull();
    overlay.removeAttribute(FOLIO_INPUT_HIDDEN_ATTRIBUTE);
    expect(getFolioCaretViewportRect(view)).not.toBeNull();
  });
});
