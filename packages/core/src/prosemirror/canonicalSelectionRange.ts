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
    throw new TypeError("A cell selection has no canonical linear range.");
  }
  if (selection instanceof AllSelection) {
    return {
      from: canonicalContentPosition(doc, selection.from),
      to: canonicalContentPosition(doc, selection.to),
    };
  }
  return { from: selection.from, to: selection.to };
};

/** Cell selections address each selected cell independently, in document order. */
export const canonicalCellSelectionRanges = (selection: CellSelection) => {
  const ranges: { from: number; to: number }[] = [];
  selection.forEachCell((cell, cellPosition) => {
    cell.descendants((node, position) => {
      if (!node.isTextblock) return true;
      ranges.push({
        from: cellPosition + 1 + position + 1,
        to: cellPosition + 1 + position + 1 + node.content.size,
      });
      return false;
    });
  });
  return ranges.toSorted((left, right) => left.from - right.from);
};

/** A selection is either one linear range or a document-ordered set of cell ranges. */
export const canonicalSelectionRanges = (state: EditorState) =>
  state.selection instanceof CellSelection
    ? canonicalCellSelectionRanges(state.selection)
    : [canonicalSelectionRange(state)];
