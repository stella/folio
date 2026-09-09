import { Mark } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { panic } from "better-result";

import { mergeTextFormatting } from "../utils/textFormattingMerge";
import { marksToTextFormatting } from "./conversion/fromProseDoc";
import { textFormattingToMarks } from "./conversion/toProseDoc";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import {
  getParagraphMarkSuppressionOverrides,
  paragraphFormattingForRun,
  paragraphRunStyleContext,
  resolveEffectiveRunStyleFormatting,
  type RunStyleResolver,
} from "./runStyleFormatting";

type RebaseParagraphRunFormattingOptions = {
  nextAttrs: Record<string, unknown>;
  paragraphPosition: number;
  styleResolver: RunStyleResolver;
  tr: Transaction;
};

/**
 * Change paragraph attrs and re-resolve inherited run marks in the new style
 * context without turning the old rendered style into direct run formatting.
 */
export const setParagraphAttrsWithRebasedRunFormatting = ({
  nextAttrs,
  paragraphPosition,
  styleResolver,
  tr,
}: RebaseParagraphRunFormattingOptions): Transaction => {
  const paragraph = tr.doc.nodeAt(paragraphPosition);
  if (!paragraph || paragraph.type.name !== "paragraph") {
    return panic("Cannot rebase run formatting outside a paragraph", {
      nodeType: paragraph?.type.name,
      paragraphPosition,
    });
  }

  const previousContext = paragraphRunStyleContext(paragraph, styleResolver);
  const nextParagraph = paragraph.type.create(nextAttrs, paragraph.content, paragraph.marks);
  const nextContext = paragraphRunStyleContext(nextParagraph, styleResolver);
  const changes: {
    currentFormattingMarks: readonly Mark[];
    from: number;
    marks: readonly Mark[];
    to: number;
  }[] = [];

  paragraph.descendants((node, relativePosition) => {
    if (!node.isInline || (!node.isText && !node.isLeaf)) {
      return true;
    }
    const authoredFormatting = marksToTextFormatting(node.marks, {
      baseParagraphFormatting: previousContext.baseParagraphFormatting,
      inheritedFormatting: previousContext.paragraphFormatting,
      paragraphMarkFormatting: previousContext.paragraphMarkFormatting,
      paragraphMarkPrecedesStyle: previousContext.paragraphMarkPrecedesStyle,
      styleResolver,
    });
    const paragraphFormatting = paragraphFormattingForRun(
      node.marks,
      nextContext,
      authoredFormatting,
    );
    const inheritedFormatting = resolveEffectiveRunStyleFormatting({
      marks: node.marks,
      paragraphFormatting,
      styleResolver,
    });
    const effectiveFormatting = mergeTextFormatting(inheritedFormatting, authoredFormatting);
    const paragraphMarkOverrides = getParagraphMarkSuppressionOverrides({
      directFormatting: authoredFormatting,
      paragraphMarkFormatting: nextContext.paragraphMarkFormatting,
      suppressedFormatting: paragraphFormatting,
    });
    const overrideFormatting = mergeTextFormatting(paragraphMarkOverrides, authoredFormatting);
    const formattingMarks = textFormattingToMarks(effectiveFormatting, {
      overrideFormatting,
      directFormatting: authoredFormatting,
    });
    const characterStyle = node.marks.find(({ type }) => type.name === "characterStyle");
    if (characterStyle) {
      formattingMarks.push(characterStyle);
    }
    const nextMarks = Mark.setFrom([
      ...node.marks.filter(({ type }) => !RUN_FORMATTING_MARK_NAMES.has(type.name)),
      ...formattingMarks,
    ]);
    if (Mark.sameSet(node.marks, nextMarks)) {
      return false;
    }
    const from = paragraphPosition + 1 + relativePosition;
    changes.push({
      currentFormattingMarks: node.marks.filter(({ type }) =>
        RUN_FORMATTING_MARK_NAMES.has(type.name),
      ),
      from,
      marks: nextMarks,
      to: from + node.nodeSize,
    });
    return false;
  });

  tr = tr.setNodeMarkup(paragraphPosition, undefined, nextAttrs);
  for (const { currentFormattingMarks, from, marks, to } of changes) {
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
