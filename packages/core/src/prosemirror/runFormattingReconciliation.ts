import { Mark, type Node as PMNode } from "prosemirror-model";

import type { TextFormatting } from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import {
  expectCharacterStyleMarkAttrs,
  expectRunIdentityMarkAttrs,
} from "./attrs";
import { textFormattingToMarks } from "./extensions/marks/markUtils";
import { marksToTextFormatting } from "./runFormattingFromMarks";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { RUN_IDENTITY_MARK_NAME } from "./runIdentity";
import {
  getParagraphMarkSuppressionOverrides,
  hasDirectRunFormatting,
  isParagraphMarkSuppressionEligible,
  paragraphFormattingForRun,
  type ParagraphRunStyleContext,
  resolveEffectiveRunStyleFormatting,
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

/** Preserve the authored distinction between an absent and empty property set in revisions. */
export const readAuthoredRunPropertySnapshot = (
  options: ReadAuthoredRunFormattingOptions,
): TextFormatting | undefined => {
  const formatting = readAuthoredRunFormatting(options);
  if (Object.keys(formatting).length > 0) return formatting;
  const identity = options.marks.find(
    ({ type }) => type.name === RUN_IDENTITY_MARK_NAME,
  );
  return identity && expectRunIdentityMarkAttrs(identity).emptyFormatting
    ? formatting
    : undefined;
};

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
  const currentCharacterStyle = marks.find(
    ({ type }) => type.name === "characterStyle",
  );
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
  const effectiveFormatting = mergeTextFormatting(
    inheritedFormatting,
    authoredFormatting,
  );
  const paragraphMarkOverrides = isParagraphMarkSuppressionEligible(
    authoredFormatting,
  )
    ? getParagraphMarkSuppressionOverrides({
        directFormatting: authoredFormatting,
        paragraphMarkFormatting: context.paragraphMarkFormatting,
        suppressedFormatting: paragraphFormatting,
      })
    : undefined;
  const overrideFormatting = mergeTextFormatting(
    paragraphMarkOverrides,
    authoredFormatting,
  );
  const formattingMarks = textFormattingToMarks(
    effectiveFormatting,
    node.type.schema,
    {
      overrideFormatting,
      directFormatting: authoredFormatting,
      authoredCarrier: "preserve",
    },
  );

  return Mark.setFrom([
    ...marks.filter(({ type }) => !RUN_FORMATTING_MARK_NAMES.has(type.name)),
    ...formattingMarks,
    ...(characterStyle ? [characterStyle] : []),
  ]);
};
