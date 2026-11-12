import type { Mark, Node as PMNode } from "prosemirror-model";

import type { RunPropertyChange } from "../types/document";
import { expectRunIdentityMarkAttrs } from "./attrs";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { reconcileRunFormattingMarks } from "./runFormattingReconciliation";
import { RUN_IDENTITY_MARK_NAME, runIdentityAttrs } from "./runIdentity";
import {
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "./runStyleFormatting";

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
  const formatting = marks.filter(({ type }) =>
    RUN_FORMATTING_MARK_NAMES.has(type.name),
  );
  if (
    previousFormatting === undefined ||
    Object.keys(previousFormatting).length > 0
  )
    return formatting;
  const identityType = node.type.schema.marks[RUN_IDENTITY_MARK_NAME];
  if (!identityType) return formatting;
  const identity = node.marks.find(({ type }) => type === identityType);
  const attrs = identity ? expectRunIdentityMarkAttrs(identity) : undefined;
  // The id groups inline leaves during conversion; it is not a package identity.
  return [
    ...formatting,
    identityType.create(
      runIdentityAttrs(attrs?.id ?? 0, {
        preservedAttributes: attrs?.preservedAttributes ?? undefined,
        preserved: attrs?.preserved ?? undefined,
        emptyFormatting: true,
      }),
    ),
  ];
};
