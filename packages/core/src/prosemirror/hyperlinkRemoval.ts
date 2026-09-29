/**
 * Removing a hyperlink, tracked when suggesting.
 *
 * A hyperlink is a wrapper around its runs, not a run property, so no
 * property revision can record its removal. With suggestions on, the linked
 * text is struck through and the same text follows it as an insertion,
 * without the link: rejecting brings the link back, accepting leaves the
 * plain text an untracked removal leaves.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";

import {
  isSuggestionModeActive,
  makeRevisionInfo,
  suggestRangeDeletion,
} from "./plugins/suggestionMode";

type LinkedRange = { from: number; to: number; nodes: PMNode[] };

/** Contiguous stretches of linked text in [`from`, `to`) of `doc`. */
const linkedRanges = (doc: PMNode, from: number, to: number): LinkedRange[] => {
  const ranges: LinkedRange[] = [];
  doc.nodesBetween(from, to, (node, position) => {
    if (!node.isText || !node.marks.some(({ type }) => type.name === "hyperlink")) {
      return true;
    }
    if (node.marks.some(({ type }) => type.name === "deletion")) {
      return true;
    }
    const start = Math.max(position, from);
    const end = Math.min(position + node.nodeSize, to);
    if (start >= end) {
      return true;
    }
    const piece = node.cut(start - position, end - position);
    const last = ranges.at(-1);
    if (last && last.to === start) {
      last.to = end;
      last.nodes.push(piece);
    } else {
      ranges.push({ from: start, to: end, nodes: [piece] });
    }
    return true;
  });
  return ranges;
};

/**
 * Remove the hyperlinks from [`from`, `to`) of `tr.doc`: directly, or, with
 * suggestions on, as a tracked replacement of the linked text by plain text.
 */
export const removeHyperlinkInRange = (
  state: EditorState,
  tr: Transaction,
  from: number,
  to: number,
): Transaction => {
  const hyperlinkType = state.schema.marks["hyperlink"];
  if (!hyperlinkType) {
    return tr;
  }
  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  const revision = makeRevisionInfo(state);
  if (!isSuggestionModeActive(state) || !insertionType || !deletionType || !revision) {
    return tr.removeMark(from, to, hyperlinkType);
  }
  const insertion = insertionType.create({
    revisionId: revision.id,
    author: revision.author,
    date: revision.date,
  });
  // Last stretch first, so the earlier positions stay valid.
  for (const range of linkedRanges(tr.doc, from, to).toReversed()) {
    const plain = range.nodes.map((node) =>
      node.mark(
        insertion.addToSet(
          node.marks.filter(
            ({ type }) => type !== hyperlinkType && type !== insertionType && type !== deletionType,
          ),
        ),
      ),
    );
    tr.insert(range.to, plain);
    suggestRangeDeletion(state, tr, range.from, range.to);
  }
  return tr;
};
