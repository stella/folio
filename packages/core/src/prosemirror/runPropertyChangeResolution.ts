import type { Mark, Node as PMNode } from "prosemirror-model";

import type { RunPropertyChange } from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import { expectCharacterStyleMarkAttrs } from "./attrs";
import { textFormattingToMarks } from "./conversion/toProseDoc";
import {
  paragraphFormattingForRun,
  type ParagraphRunStyleContext,
  resolveEffectiveRunStyleFormatting,
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
  const characterStyleMark = node.marks.find(({ type }) => type.name === "characterStyle");
  const characterStyleAttrs = characterStyleMark
    ? expectCharacterStyleMarkAttrs(characterStyleMark)
    : undefined;
  const preservedCharacterStyleAttrs =
    previousFormatting?.styleId !== undefined &&
    characterStyleAttrs?.styleId === previousFormatting.styleId
      ? characterStyleAttrs
      : undefined;
  const styleFormatting = preservedCharacterStyleAttrs
    ? resolveEffectiveRunStyleFormatting({
        marks: node.marks,
        paragraphFormatting: paragraphFormattingForRun(
          node.marks,
          paragraphContext,
          previousFormatting,
        ),
        ...(styleResolver !== undefined ? { styleResolver } : {}),
      })
    : paragraphFormattingForRun(node.marks, paragraphContext, previousFormatting);
  const effectivePreviousFormatting = mergeTextFormatting(styleFormatting, previousFormatting);
  let marks: readonly Mark[] = [];
  for (const mark of textFormattingToMarks(effectivePreviousFormatting, {
    overrideFormatting: previousFormatting,
    directFormatting: previousFormatting,
  })) {
    marks = mark.addToSet(marks);
  }
  if (previousFormatting?.styleId) {
    const characterStyle = node.type.schema.marks["characterStyle"];
    if (characterStyle) {
      marks = characterStyle
        .create(preservedCharacterStyleAttrs ?? { styleId: previousFormatting.styleId })
        .addToSet(marks);
    }
  }
  return marks;
};
