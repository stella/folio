import { Mark, type Node as PMNode } from "prosemirror-model";

import type { TextFormatting } from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import { marksToTextFormatting } from "./conversion/fromProseDoc";
import { textFormattingToMarks } from "./conversion/toProseDoc";
import { expectCharacterStyleMarkAttrs } from "./attrs";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import {
  getParagraphMarkSuppressionOverrides,
  hasDirectRunFormatting,
  paragraphFormattingForRun,
  resolveEffectiveRunStyleFormatting,
  type ParagraphRunStyleContext,
  type RunStyleResolver,
} from "./runStyleFormatting";

type ReadAuthoredRunFormattingOptions = {
  context: ParagraphRunStyleContext;
  marks: readonly Mark[];
  styleResolver?: RunStyleResolver | null;
};

/** Read authored run properties without confusing inherited visuals for direct formatting. */
export const readAuthoredRunFormatting = ({
  context,
  marks,
  styleResolver,
}: ReadAuthoredRunFormattingOptions): TextFormatting =>
  marksToTextFormatting(marks, {
    baseParagraphFormatting: context.baseParagraphFormatting,
    inheritedFormatting: context.paragraphFormatting,
    paragraphMarkFormatting: context.paragraphMarkFormatting,
    paragraphMarkPrecedesStyle: context.paragraphMarkPrecedesStyle,
    ...(styleResolver !== undefined ? { styleResolver } : {}),
  });

type ReconcileRunFormattingMarksOptions = {
  authoredFormatting: TextFormatting;
  context: ParagraphRunStyleContext;
  node: PMNode;
  styleResolver?: RunStyleResolver | null;
};

/**
 * Rebuild a run's physical formatting marks from its authored properties and
 * current style context. This is the inverse of {@link readAuthoredRunFormatting}:
 * direct provenance stays direct while inherited formatting remains visual only.
 */
export const reconcileRunFormattingMarks = ({
  authoredFormatting,
  context,
  node,
  styleResolver,
}: ReconcileRunFormattingMarksOptions): readonly Mark[] => {
  const { marks } = node;
  const currentCharacterStyle = marks.find(({ type }) => type.name === "characterStyle");
  const currentCharacterStyleAttrs = currentCharacterStyle
    ? expectCharacterStyleMarkAttrs(currentCharacterStyle)
    : undefined;
  const characterStyleType = node.type.schema.marks["characterStyle"];
  const characterStyle =
    authoredFormatting.styleId !== undefined && characterStyleType
      ? characterStyleType.create(
          currentCharacterStyleAttrs?.styleId === authoredFormatting.styleId
            ? currentCharacterStyleAttrs
            : { styleId: authoredFormatting.styleId },
        )
      : undefined;
  const directCarrierType = node.type.schema.marks["runFormattingOverride"];
  const styleResolutionMarks = [
    ...(hasDirectRunFormatting(authoredFormatting) && directCarrierType
      ? [directCarrierType.create()]
      : []),
    ...(characterStyle ? [characterStyle] : []),
  ];
  const paragraphFormatting = paragraphFormattingForRun({
    marks: styleResolutionMarks,
    context,
    directFormatting: authoredFormatting,
  });
  const inheritedFormatting = resolveEffectiveRunStyleFormatting({
    marks: styleResolutionMarks,
    paragraphFormatting,
    ...(styleResolver !== undefined ? { styleResolver } : {}),
  });
  const effectiveFormatting = mergeTextFormatting(inheritedFormatting, authoredFormatting);
  const paragraphMarkOverrides = getParagraphMarkSuppressionOverrides({
    directFormatting: authoredFormatting,
    paragraphMarkFormatting: context.paragraphMarkFormatting,
    suppressedFormatting: paragraphFormatting,
  });
  const overrideFormatting = mergeTextFormatting(paragraphMarkOverrides, authoredFormatting);
  const formattingMarks = textFormattingToMarks(effectiveFormatting, {
    overrideFormatting,
    directFormatting: authoredFormatting,
  });

  return Mark.setFrom([
    ...marks.filter(({ type }) => !RUN_FORMATTING_MARK_NAMES.has(type.name)),
    ...formattingMarks,
    ...(characterStyle ? [characterStyle] : []),
  ]);
};
