import type { Mark, Node as PMNode } from "prosemirror-model";

import type { RunPropertyChange } from "../types/document";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { reconcileRunFormattingMarks } from "./runFormattingReconciliation";
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
