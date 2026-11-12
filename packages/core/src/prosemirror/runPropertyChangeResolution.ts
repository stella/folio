import type { Mark, Node as PMNode } from "prosemirror-model";

import type { RunPropertyChange, TextFormatting } from "../types/document";
import { expectRunIdentityMarkAttrs } from "./attrs";
import { runFormattingInlineAtomDisposition } from "./runFormattingInlineCarriers";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { RUN_IDENTITY_MARK_NAME, runIdentityAttrs } from "./runIdentity";
import {
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "./runStyleFormatting";

type RejectedRunFormattingOptions = {
  node: PMNode;
  paragraphContext: ParagraphRunStyleContext;
  previousFormatting: RunPropertyChange["previousFormatting"];
  styleResolver?: RunStyleResolver | null;
};

/** The one owner of what a reject restores: marks and an atom's run formatting. */
export const rejectedRunFormatting = ({
  node,
  paragraphContext,
  previousFormatting,
  styleResolver,
}: RejectedRunFormattingOptions): {
  marks: readonly Mark[] | null;
  atomRunFormatting: TextFormatting | null;
} => {
  // A picture, a shape or a text box anchor keeps its run's formatting in
  // its attrs; a picture or a shape carries no formatting marks besides.
  // Null means nothing to restore: no marks to rebuild, or no record kept.
  const keepsRunFormatting = Object.hasOwn(node.attrs, "_docxRunFormatting");
  return {
    marks:
      !keepsRunFormatting ||
      node.marks.some(({ type }) => RUN_FORMATTING_MARK_NAMES.has(type.name))
        ? rebuiltRunFormattingMarks({
            node,
            paragraphContext,
            previousFormatting,
            styleResolver,
          })
        : null,
    atomRunFormatting: keepsRunFormatting ? { ...previousFormatting } : null,
  };
};

/** The pre-change run marks, against the paragraph's live style cascade. */
const rebuiltRunFormattingMarks = ({
  node,
  paragraphContext,
  previousFormatting,
  styleResolver,
}: RejectedRunFormattingOptions): readonly Mark[] => {
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
