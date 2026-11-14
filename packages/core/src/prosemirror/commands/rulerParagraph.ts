import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { executeEditorCommand } from "../executeEditorCommand";
import {
  removeTabStop,
  setIndentFirstLine,
  setIndentLeft,
  setIndentRight,
} from "./paragraph";

/** A paragraph edit made by dragging or clicking the horizontal ruler. */
export type RulerParagraphEdit =
  | { type: "indentLeft"; twips: number }
  | { type: "indentRight"; twips: number }
  /** Negative twips are a hanging indent of that size. */
  | { type: "indentFirstLine"; twips: number }
  | { type: "removeTabStop"; position: number };

/** The registry command for a ruler edit, which carries its canonical intent. */
export const rulerParagraphCommand = (edit: RulerParagraphEdit): Command => {
  switch (edit.type) {
    case "indentLeft":
      return setIndentLeft(edit.twips);
    case "indentRight":
      return setIndentRight(edit.twips);
    case "indentFirstLine":
      return edit.twips < 0
        ? setIndentFirstLine(-edit.twips, true)
        : setIndentFirstLine(edit.twips, false);
    case "removeTabStop":
      return removeTabStop(edit.position);
  }
};

/**
 * Apply a ruler edit to the selected paragraphs through the view's command
 * owner, so a canonical session journals it as paragraph formatting.
 */
export const applyRulerParagraphEdit = (
  view: EditorView,
  edit: RulerParagraphEdit,
): boolean => executeEditorCommand(view, rulerParagraphCommand(edit));
