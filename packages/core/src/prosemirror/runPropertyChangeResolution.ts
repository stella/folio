import type { Mark, Node as PMNode } from "prosemirror-model";

import type { RunPropertyChange } from "../types/document";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { runFormattingInlineAtomDisposition } from "./runFormattingInlineCarriers";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { type ParagraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";

type ReconstructRejectedRunFormattingMarksOptions = {
  node: PMNode;
  paragraphContext: ParagraphRunStyleContext;
  previousFormatting: RunPropertyChange["previousFormatting"];
  styleResolver?: RunStyleResolver | null;
};

/**
 * Rebuild the pre-change run marks against the paragraph's live style cascade.
 * Both granular editor commands and the headless linear rewriter use this
 * owner so their reject semantics cannot diverge at the bulk threshold.
 */
export const reconstructRejectedRunFormattingMarks = ({
  node,
  paragraphContext,
  previousFormatting,
  styleResolver,
}: ReconstructRejectedRunFormattingMarksOptions): readonly Mark[] => {
  const marks = reconcileRunFormattingMarks({
    authoredFormatting: previousFormatting ?? {},
    context: paragraphContext,
    node,
    ...(styleResolver !== undefined ? { styleResolver } : {}),
  });
  return marks.filter(({ type }) => RUN_FORMATTING_MARK_NAMES.has(type.name));
};

type RestoreHistoricalRunFormattingOptions = {
  node: PMNode;
  paragraphContext: ParagraphRunStyleContext;
  styleResolver?: RunStyleResolver | null;
};

/** Restored deletions inherit the live paragraph cascade without authoring their historical visuals. */
export const restoreHistoricalRunFormatting = ({
  node,
  paragraphContext,
  styleResolver,
}: RestoreHistoricalRunFormattingOptions): readonly Mark[] => {
  const disposition = runFormattingInlineAtomDisposition(node);
  if (disposition === null || disposition === "not-a-run") return node.marks;
  if (paragraphContext.paragraphMarkFormatting === undefined) return node.marks;
  const authoredFormatting = readAuthoredRunFormatting({
    context: {
      baseParagraphFormatting: paragraphContext.baseParagraphFormatting,
      paragraphFormatting: paragraphContext.baseParagraphFormatting,
      paragraphMarkFormatting: undefined,
      paragraphMarkPrecedesStyle: false,
    },
    marks: node.marks,
    ...(styleResolver !== undefined ? { styleResolver } : {}),
  });
  return reconcileRunFormattingMarks({
    authoredFormatting,
    context: paragraphContext,
    node,
    ...(styleResolver !== undefined ? { styleResolver } : {}),
  });
};
