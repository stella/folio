import { Mark } from "prosemirror-model";

import type { TextFormatting } from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import { marksToTextFormatting } from "./conversion/fromProseDoc";
import { textFormattingToMarks } from "./conversion/toProseDoc";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import {
  getParagraphMarkSuppressionOverrides,
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
  marks: readonly Mark[];
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
  marks,
  styleResolver,
}: ReconcileRunFormattingMarksOptions): readonly Mark[] => {
  const paragraphFormatting = paragraphFormattingForRun(marks, context, authoredFormatting);
  const inheritedFormatting = resolveEffectiveRunStyleFormatting({
    marks,
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
  const characterStyleMarks = marks.filter(({ type }) => type.name === "characterStyle");

  return Mark.setFrom([
    ...marks.filter(({ type }) => !RUN_FORMATTING_MARK_NAMES.has(type.name)),
    ...formattingMarks,
    ...characterStyleMarks,
  ]);
};
