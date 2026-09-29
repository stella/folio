import { Mark, type Fragment, type Node as PMNode } from "prosemirror-model";
import { recreateProseNodeWithParagraphPropertySource } from "../docx/paragraphPropertySource";
import type { Transaction } from "prosemirror-state";

import {
  applyMarksToRunFormattingRepresentation,
  selectRunFormattingCarrierRepresentations,
} from "./runFormattingInlineCarriers";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { paragraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";

type RebaseParagraphRunsOptions = {
  tr: Transaction;
  position: number;
  previous: PMNode;
  target?: PMNode;
  styleResolver: RunStyleResolver | null;
};

/** Rebuild rendered marks after the paragraph's style cascade changes. */
export const rebaseParagraphRuns = ({
  tr,
  position,
  previous,
  target,
  styleResolver,
}: RebaseParagraphRunsOptions): Transaction => {
  const paragraph = tr.doc.nodeAt(position);
  if (!paragraph) return tr;
  const sourceContext = paragraphRunStyleContext(previous, styleResolver);
  const targetContext = paragraphRunStyleContext(target ?? paragraph, styleResolver);
  const representations = selectRunFormattingCarrierRepresentations({
    doc: tr.doc,
    from: position + 1,
    to: position + paragraph.nodeSize - 1,
  });
  for (const representation of representations) {
    const authoredFormatting = readAuthoredRunFormatting({
      context: sourceContext,
      marks: representation.node.marks,
      styleResolver,
    });
    const marks = reconcileRunFormattingMarks({
      authoredFormatting,
      context: targetContext,
      node: representation.node,
      styleResolver,
    });
    applyMarksToRunFormattingRepresentation({ tr, representation, marks });
  }
  return tr;
};

type RebasedRunFormatting = {
  before: PMNode;
  after: PMNode;
  position: number;
};

type RebaseParagraphRunContentOptions = {
  paragraph: PMNode;
  target: PMNode;
  position: number;
  styleResolver: RunStyleResolver | null;
  onRebased: (resolution: RebasedRunFormatting) => void;
};

/** Rebase an immutable paragraph chunk once, before assembling a resolved join chain. */
export const rebaseParagraphRunContent = ({
  paragraph,
  target,
  position,
  styleResolver,
  onRebased,
}: RebaseParagraphRunContentOptions): Fragment => {
  const sourceContext = paragraphRunStyleContext(paragraph, styleResolver);
  const targetContext = paragraphRunStyleContext(target, styleResolver);
  const replacements = new Map<PMNode, PMNode>();
  const representations = selectRunFormattingCarrierRepresentations({
    doc: paragraph,
    from: 0,
    to: paragraph.content.size,
  });
  for (const representation of representations) {
    const marks = reconcileRunFormattingMarks({
      authoredFormatting: readAuthoredRunFormatting({
        context: sourceContext,
        marks: representation.node.marks,
        styleResolver,
      }),
      context: targetContext,
      node: representation.node,
      styleResolver,
    });
    if (Mark.sameSet(marks, representation.node.marks)) continue;
    const rebased = representation.node.mark(marks);
    replacements.set(representation.node, rebased);
    onRebased({
      before: representation.node,
      after: rebased,
      position: position + 1 + representation.position,
    });
  }
  if (replacements.size === 0) return paragraph.content;
  const rebuild = (node: PMNode): PMNode => {
    const replacement = replacements.get(node) ?? node;
    if (node.isLeaf) return replacement;
    const children: PMNode[] = [];
    let changed = false;
    replacement.forEach((child) => {
      const next = rebuild(child);
      children.push(next);
      changed ||= next !== child;
    });
    return changed
      ? recreateProseNodeWithParagraphPropertySource(replacement, { content: children })
      : replacement;
  };
  return rebuild(paragraph).content;
};
