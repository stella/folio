import { Mark, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import type { Transform } from "prosemirror-transform";
import { panic } from "better-result";

import {
  expandRunFormattingCarrier,
  type RunFormattingCarrierRepresentation,
} from "./runFormattingInlineCarriers";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { sameRunStyleContext } from "./rebaseParagraphRuns";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import {
  paragraphRunStyleContext,
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "./runStyleFormatting";

type RebaseParagraphRunFormattingOptions = {
  nextAttrs: Record<string, unknown>;
  paragraphPosition: number;
  shouldRebase?: (node: PMNode) => boolean;
  styleResolver: RunStyleResolver;
  tr: Transaction;
};

const paragraphAt = (tr: Transform, paragraphPosition: number): PMNode => {
  const paragraph = tr.doc.nodeAt(paragraphPosition);
  if (!paragraph || paragraph.type.name !== "paragraph") {
    return panic("Cannot rebase run formatting outside a paragraph", {
      nodeType: paragraph?.type.name,
      paragraphPosition,
    });
  }
  return paragraph;
};

type RebaseParagraphRunsOptions<T extends Transform> = {
  /** The style context the runs' marks were resolved in. */
  previousContext: ParagraphRunStyleContext;
  paragraphPosition: number;
  /** Content offsets inside the paragraph to rebase; all of it when omitted. */
  range?: { from: number; to: number };
  shouldRebase?: (node: PMNode) => boolean;
  styleResolver: RunStyleResolver;
  tr: T;
};

/**
 * Re-resolve the inherited marks of runs resolved in `previousContext` in the
 * paragraph's current style context: direct formatting stays direct, and what
 * the old context lent the runs goes with it.
 */
export const rebaseParagraphRuns = <T extends Transform>({
  previousContext,
  paragraphPosition,
  range,
  shouldRebase,
  styleResolver,
  tr,
}: RebaseParagraphRunsOptions<T>): T => {
  const paragraph = paragraphAt(tr, paragraphPosition);
  const nextContext = paragraphRunStyleContext(paragraph, styleResolver);
  // An unchanged cascade leaves the runs, their carriers and provenance as they are.
  if (sameRunStyleContext(previousContext, nextContext)) {
    return tr;
  }
  const changes: {
    attrs: Readonly<Record<string, unknown>>;
    currentFormattingMarks: readonly Mark[];
    from: number;
    isText: boolean;
    marks: readonly Mark[];
    to: number;
  }[] = [];

  const collectRebasedRepresentation = ({
    node,
    position,
  }: RunFormattingCarrierRepresentation): void => {
    if (shouldRebase && !shouldRebase(node)) {
      return;
    }
    const authoredFormatting = readAuthoredRunFormatting({
      context: previousContext,
      marks: node.marks,
      styleResolver,
    });
    const nextMarks = reconcileRunFormattingMarks({
      authoredFormatting,
      context: nextContext,
      node,
      styleResolver,
    });
    if (Mark.sameSet(node.marks, nextMarks)) {
      return;
    }
    // A text node can run past the range: two runs with the same marks are
    // one node once their paragraphs join. Only the range's part is rebased.
    const start = paragraphPosition + 1;
    const from = node.isText && range ? Math.max(position, start + range.from) : position;
    const to =
      node.isText && range
        ? Math.min(position + node.nodeSize, start + range.to)
        : position + node.nodeSize;
    if (to <= from) {
      return;
    }
    changes.push({
      attrs: node.attrs,
      currentFormattingMarks: node.marks.filter(({ type }) =>
        RUN_FORMATTING_MARK_NAMES.has(type.name),
      ),
      from,
      isText: node.isText,
      marks: nextMarks,
      to,
    });
  };

  const visit = (node: PMNode, relativePosition: number): boolean => {
    if (!node.isInline) {
      return true;
    }
    const carrier = expandRunFormattingCarrier(node, paragraphPosition + 1 + relativePosition);
    if (!carrier) {
      return !node.isAtom;
    }
    for (const representation of carrier.representations) {
      collectRebasedRepresentation(representation);
    }
    return false;
  };
  if (range === undefined) {
    paragraph.descendants(visit);
  } else {
    paragraph.nodesBetween(range.from, range.to, visit);
  }

  for (const { attrs, currentFormattingMarks, from, isText, marks, to } of changes) {
    if (!isText) {
      tr = tr.setNodeMarkup(from, undefined, attrs, marks);
      continue;
    }
    for (const mark of currentFormattingMarks) {
      tr = tr.removeMark(from, to, mark.type);
    }
    for (const mark of marks) {
      if (RUN_FORMATTING_MARK_NAMES.has(mark.type.name)) {
        tr = tr.addMark(from, to, mark);
      }
    }
  }
  return tr;
};

/**
 * Change paragraph attrs and re-resolve inherited run marks in the new style
 * context without turning the old rendered style into direct run formatting.
 */
export const setParagraphAttrsWithRebasedRunFormatting = ({
  nextAttrs,
  paragraphPosition,
  shouldRebase,
  styleResolver,
  tr,
}: RebaseParagraphRunFormattingOptions): Transaction => {
  const previousContext = paragraphRunStyleContext(
    paragraphAt(tr, paragraphPosition),
    styleResolver,
  );
  tr = tr.setNodeMarkup(paragraphPosition, undefined, nextAttrs);
  return rebaseParagraphRuns({
    previousContext,
    paragraphPosition,
    ...(shouldRebase ? { shouldRebase } : {}),
    styleResolver,
    tr,
  });
};
