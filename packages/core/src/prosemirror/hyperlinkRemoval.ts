/**
 * Removing a hyperlink takes the link and its Hyperlink character style off
 * the text, as the Remove Hyperlink command does; other run formatting stays.
 *
 * A link is a wrapper around its runs, not a run property, so no property
 * revision can record its removal. With suggestions on, the linked text is
 * struck through and the same text follows it as an insertion, without the
 * link. Both carry a run-property change whose previous formatting is the
 * Hyperlink style: the struck text no longer states the style, the inserted
 * text never did. Rejecting brings the link back as it was; accepting leaves
 * the plain text an untracked removal leaves.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";

import { BUILT_IN_STYLE_NAME } from "../docx/builtInStyles";
import type { RunPropertyChange, TextFormatting } from "../types/document";
import { getDocumentBuiltInStyles, getDocumentStyleResolver } from "./plugins/documentStyles";
import {
  isSuggestionModeActive,
  makeRevisionInfo,
  suggestRangeDeletion,
} from "./plugins/suggestionMode";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { paragraphRunStyleContextAt } from "./runStyleFormatting";

type LinkedPiece = { from: number; to: number; node: PMNode };
type LinkedRange = { from: number; to: number; pieces: LinkedPiece[] };

/** Contiguous stretches of linked, not yet deleted, text in [`from`, `to`) of `doc`. */
const linkedRanges = (doc: PMNode, from: number, to: number): LinkedRange[] => {
  const ranges: LinkedRange[] = [];
  doc.nodesBetween(from, to, (node, position) => {
    if (
      !node.isText ||
      !node.marks.some(({ type }) => type.name === "hyperlink") ||
      node.marks.some(({ type }) => type.name === "deletion")
    ) {
      return true;
    }
    const start = Math.max(position, from);
    const end = Math.min(position + node.nodeSize, to);
    if (start >= end) {
      return true;
    }
    const piece = { from: start, to: end, node: node.cut(start - position, end - position) };
    const last = ranges.at(-1);
    if (last && last.to === start) {
      last.to = end;
      last.pieces.push(piece);
    } else {
      ranges.push({ from: start, to: end, pieces: [piece] });
    }
    return true;
  });
  return ranges;
};

type Unlinked = {
  /** The piece's marks without its Hyperlink style (the link mark itself kept). */
  marks: readonly Mark[];
  /** Its authored formatting while it had the style, or null when it had none. */
  previousFormatting: TextFormatting | null;
};

/** A piece's marks once the Hyperlink character style is taken off it. */
const withoutHyperlinkStyle = (state: EditorState, doc: PMNode, piece: LinkedPiece): Unlinked => {
  const styleResolver = getDocumentStyleResolver(state);
  const builtInStyles = getDocumentBuiltInStyles(state);
  const context = paragraphRunStyleContextAt({ doc, pos: piece.from, styleResolver });
  const previousFormatting = readAuthoredRunFormatting({
    context,
    marks: piece.node.marks,
    styleResolver,
  });
  const styleId = previousFormatting.styleId;
  const isHyperlinkStyle =
    styleId !== undefined &&
    (builtInStyles.builtInNameOf(styleId) === BUILT_IN_STYLE_NAME.hyperlink ||
      styleId === "Hyperlink");
  if (!isHyperlinkStyle) {
    return { marks: piece.node.marks, previousFormatting: null };
  }
  const { styleId: _removed, ...formatting } = previousFormatting;
  return {
    marks: reconcileRunFormattingMarks({
      authoredFormatting: formatting,
      context,
      node: piece.node,
      styleResolver,
    }),
    previousFormatting,
  };
};

/** Replace the marks of the text at [`from`, `to`) of `tr.doc` with `marks`. */
const setTextMarks = (
  tr: Transaction,
  from: number,
  to: number,
  current: readonly Mark[],
  marks: readonly Mark[],
): void => {
  for (const mark of current) {
    if (!mark.isInSet(marks)) {
      tr.removeMark(from, to, mark);
    }
  }
  for (const mark of marks) {
    if (!mark.isInSet(current)) {
      tr.addMark(from, to, mark);
    }
  }
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
  const { schema } = state;
  const hyperlinkType = schema.marks["hyperlink"];
  if (!hyperlinkType) {
    return tr;
  }
  const insertionType = schema.marks["insertion"];
  const deletionType = schema.marks["deletion"];
  const propertyChangeType = schema.marks["runPropertyChange"];
  const tracked =
    isSuggestionModeActive(state) &&
    insertionType !== undefined &&
    deletionType !== undefined &&
    propertyChangeType !== undefined;
  const revisionAttrs = () => {
    const revision = makeRevisionInfo(state);
    return revision ? { id: revision.id, author: revision.author, date: revision.date } : null;
  };
  const insertionInfo = tracked ? revisionAttrs() : null;
  const insertion =
    insertionType && insertionInfo
      ? insertionType.create({
          revisionId: insertionInfo.id,
          author: insertionInfo.author,
          date: insertionInfo.date,
        })
      : null;
  const propertyChange = (previousFormatting: TextFormatting): Mark | null => {
    const info = revisionAttrs();
    if (!propertyChangeType || !info) {
      return null;
    }
    const change: RunPropertyChange = { type: "runPropertyChange", info, previousFormatting };
    return propertyChangeType.create({ changes: [change] });
  };
  const hasPendingPropertyChange = (marks: readonly Mark[]): boolean =>
    marks.some(({ type }) => type === propertyChangeType);

  // Last stretch first, so the earlier positions stay valid.
  for (const range of linkedRanges(tr.doc, from, to).toReversed()) {
    const unlinked = range.pieces.map((piece) => ({
      piece,
      ...withoutHyperlinkStyle(state, tr.doc, piece),
    }));
    if (!insertion || !deletionType) {
      for (const { piece, marks } of unlinked) {
        setTextMarks(
          tr,
          piece.from,
          piece.to,
          piece.node.marks,
          marks.filter(({ type }) => type !== hyperlinkType),
        );
      }
      continue;
    }
    const plain = unlinked.map(({ piece, marks, previousFormatting }) => {
      const change =
        previousFormatting && !hasPendingPropertyChange(marks)
          ? propertyChange(previousFormatting)
          : null;
      const kept = marks.filter(
        ({ type }) => type !== hyperlinkType && type !== insertionType && type !== deletionType,
      );
      return piece.node.mark(insertion.addToSet(change ? change.addToSet(kept) : kept));
    });
    tr.insert(range.to, plain);
    for (const { piece, marks, previousFormatting } of unlinked.toReversed()) {
      const change =
        previousFormatting && !hasPendingPropertyChange(piece.node.marks)
          ? propertyChange(previousFormatting)
          : null;
      setTextMarks(
        tr,
        piece.from,
        piece.to,
        piece.node.marks,
        change ? change.addToSet(marks) : marks,
      );
    }
    suggestRangeDeletion(state, tr, range.from, range.to);
  }
  return tr;
};
