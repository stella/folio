import type { Mark, Node as PMNode } from "prosemirror-model";

import type { RunPropertyChange } from "../types/document";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { type ParagraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";

type ReconstructResolvedRunFormattingMarksOptions = {
  node: PMNode;
  paragraphContext: ParagraphRunStyleContext;
  previousFormatting: RunPropertyChange["previousFormatting"];
  mode: "accept" | "reject";
  styleResolver?: RunStyleResolver | null;
};

/**
 * Rebuild the accepted or rejected run marks against the live style cascade.
 * Resolution is also a canonicalization boundary: merely removing the
 * property-change carrier can leave redundant visual/provenance attrs that do
 * not survive Document -> PM projection. Both granular commands and the
 * headless linear rewriter use this owner so neither semantics nor source
 * identity can diverge at the bulk threshold.
 */
export const reconstructResolvedRunFormattingMarks = ({
  node,
  paragraphContext,
  previousFormatting,
  mode,
  styleResolver,
}: ReconstructResolvedRunFormattingMarksOptions): readonly Mark[] => {
  const authoredFormatting =
    mode === "reject"
      ? (previousFormatting ?? {})
      : readAuthoredRunFormatting({
          context: paragraphContext,
          marks: node.marks,
          ...(styleResolver !== undefined ? { styleResolver } : {}),
        });
  const marks = reconcileRunFormattingMarks({
    authoredFormatting,
    context: paragraphContext,
    node,
    ...(styleResolver !== undefined ? { styleResolver } : {}),
  });
  return marks.filter(({ type }) => RUN_FORMATTING_MARK_NAMES.has(type.name));
};
