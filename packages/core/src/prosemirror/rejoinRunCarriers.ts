/**
 * Rejoin a run that resolving a revision left split.
 *
 * A revision splits the run it cuts into (`w:r`, `w:del > w:r`, `w:r`), and a
 * run-property change rewrites the formatting marks of the piece it covers.
 * Once the revision is resolved the pieces write the same `w:rPr`, but the
 * editor can hold that formatting on different mark carriers (a restored
 * `w:rPrChange` is rebuilt from its `w:rPr`), and leaves with different marks
 * are written as different runs. At a resolved boundary, a leaf that writes
 * the same run as the leaf before it takes that leaf's marks, so the two are
 * one run again.
 *
 * Leaves the source wrote as separate runs keep them apart: a run identity
 * (`runIdentity`) differs between them, and it is compared, not rewritten.
 * The pieces of one run read back from a saved redline share an identity
 * (`runIdentityAcrossRevisions.ts`).
 */

import { Mark, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { formattingEquals } from "../docx/runConsolidator";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { readAuthoredRunFormatting } from "./runFormattingReconciliation";
import {
  paragraphRunStyleContextAt,
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "./runStyleFormatting";

const isFormattingMark = (mark: Mark): boolean => RUN_FORMATTING_MARK_NAMES.has(mark.type.name);

type ContinuedRunMarksOptions = {
  left: PMNode;
  right: PMNode;
  /** The paragraph's run style context, read only when the marks differ in formatting alone. */
  context: () => ParagraphRunStyleContext;
  styleResolver: RunStyleResolver | null;
};

/**
 * The marks `right` takes to continue `left`'s run, or `null` when the two
 * already share their marks or write different runs: another mark differs,
 * or the formatting they state does.
 */
export const continuedRunMarks = ({
  left,
  right,
  context,
  styleResolver,
}: ContinuedRunMarksOptions): readonly Mark[] | null => {
  if (!left.isText || !right.isText || Mark.sameSet(left.marks, right.marks)) {
    return null;
  }
  const otherMarks = (node: PMNode) => node.marks.filter((mark) => !isFormattingMark(mark));
  if (!Mark.sameSet(otherMarks(left), otherMarks(right))) {
    return null;
  }
  const paragraphContext = context();
  const stated = (node: PMNode) =>
    readAuthoredRunFormatting({
      context: paragraphContext,
      marks: node.marks,
      styleResolver,
    });
  return formattingEquals(stated(left), stated(right)) ? left.marks : null;
};

type RejoinRunsAtOptions = {
  tr: Transaction;
  /** Positions in `tr.doc` where resolved content began or ended. */
  boundaries: Iterable<number>;
  styleResolver: RunStyleResolver | null;
};

/** Rejoin the runs meeting at each of `boundaries`, left to right. */
export const rejoinRunsAt = ({ tr, boundaries, styleResolver }: RejoinRunsAtOptions): void => {
  for (const boundary of [...new Set(boundaries)].toSorted((left, right) => left - right)) {
    if (boundary <= 0 || boundary >= tr.doc.content.size) {
      continue;
    }
    const at = tr.doc.resolve(boundary);
    const { nodeBefore: left, nodeAfter: right } = at;
    if (!left || !right || !at.parent.inlineContent) {
      continue;
    }
    const marks = continuedRunMarks({
      left,
      right,
      context: () => paragraphRunStyleContextAt({ doc: tr.doc, pos: boundary, styleResolver }),
      styleResolver,
    });
    if (marks === null) {
      continue;
    }
    const to = boundary + right.nodeSize;
    for (const mark of right.marks) {
      if (!mark.isInSet(marks)) {
        tr.removeMark(boundary, to, mark);
      }
    }
    for (const mark of marks) {
      if (!mark.isInSet(right.marks)) {
        tr.addMark(boundary, to, mark);
      }
    }
  }
};
