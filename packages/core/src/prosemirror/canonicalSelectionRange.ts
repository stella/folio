import { AllSelection, Selection, type EditorState } from "prosemirror-state";
import type { Node } from "prosemirror-model";
import { CellSelection } from "prosemirror-tables";

/** Document edges denote the first and last editable content gaps. */
export const canonicalContentPosition = (doc: Node, position: number): number => {
  if (position === 0) return Selection.atStart(doc).from;
  if (position === doc.content.size) return Selection.atEnd(doc).to;
  return position;
};

/** Whole-document selection addresses paragraph content rather than document edges. */
export const canonicalSelectionRange = ({ doc, selection }: EditorState) => {
  if (selection instanceof CellSelection) {
    const first = Math.min(selection.$anchorCell.pos, selection.$headCell.pos);
    const last = Math.max(selection.$anchorCell.pos, selection.$headCell.pos);
    const firstCell = doc.nodeAt(first);
    const lastCell = doc.nodeAt(last);
    if (firstCell !== null && lastCell !== null) {
      return {
        from: first + 1 + Selection.atStart(firstCell).from,
        to: last + 1 + Selection.atEnd(lastCell).to,
      };
    }
  }
  if (selection instanceof AllSelection) {
    return {
      from: canonicalContentPosition(doc, selection.from),
      to: canonicalContentPosition(doc, selection.to),
    };
  }
  return { from: selection.from, to: selection.to };
};
