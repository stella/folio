import { AllSelection, Selection, type EditorState } from "prosemirror-state";

/** Whole-document selection addresses paragraph content rather than document edges. */
export const canonicalSelectionRange = ({ doc, selection }: EditorState) => {
  if (selection instanceof AllSelection) {
    return { from: Selection.atStart(doc).from, to: Selection.atEnd(doc).to };
  }
  return { from: selection.from, to: selection.to };
};
