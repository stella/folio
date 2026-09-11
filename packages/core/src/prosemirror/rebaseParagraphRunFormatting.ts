import { Mark, type Schema } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { panic } from "better-result";

import type { TextFormatting } from "../types/document";
import {
  expandRunFormattingCarrier,
  type RunFormattingCarrierRepresentation,
} from "./runFormattingInlineCarriers";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { paragraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";
import {
  applyNonParagraphMarkup,
  applyParagraphPropertyProjection,
  createParagraphNodeFromProjection,
  type ParagraphPropertyProjection,
} from "./paragraphPropertyMutation";

type RebaseParagraphRunFormattingOptions = {
  paragraphPosition: number;
  projection: ParagraphPropertyProjection;
  previousTableRunFormatting?: TextFormatting | null;
  styleResolver?: RunStyleResolver | null;
  tableRunFormatting?: TextFormatting | null;
  tr: Transaction;
};

type InheritedParagraphRunFormattingOptions = {
  paragraph: Parameters<typeof paragraphRunStyleContext>[0];
  styleResolver?: RunStyleResolver | null;
  tableRunFormatting?: TextFormatting | null;
};

const markInheritedFormattingAsKnownEmpty = (
  marks: readonly Mark[],
  schema: Schema,
): readonly Mark[] => {
  const override = schema.marks["runFormattingOverride"];
  if (!override || marks.length === 0) {
    return marks;
  }
  return override.create({ _authoredOn: [] }).addToSet(marks);
};

/** Build physical inherited marks with an explicit known-empty direct baseline. */
export const inheritedRunFormattingMarks = (
  formatting: TextFormatting,
  schema: Schema,
): readonly Mark[] => {
  const carrier = schema.text("projection");
  const marks = reconcileRunFormattingMarks({
    authoredFormatting: {},
    context: {
      baseParagraphFormatting: formatting,
      paragraphFormatting: formatting,
      paragraphMarkFormatting: undefined,
      paragraphMarkPrecedesStyle: false,
    },
    node: carrier,
  }).filter(({ type }) => RUN_FORMATTING_MARK_NAMES.has(type.name));
  return markInheritedFormattingAsKnownEmpty(marks, schema);
};

/** Resolve inherited-only live marks/defaults for an existing or empty paragraph. */
export const inheritedParagraphRunFormatting = ({
  paragraph,
  styleResolver,
  tableRunFormatting,
}: InheritedParagraphRunFormattingOptions) => {
  const context = paragraphRunStyleContext(paragraph, styleResolver, tableRunFormatting);
  const carrier = paragraph.type.schema.text("projection");
  const marks = reconcileRunFormattingMarks({
    authoredFormatting: {},
    context,
    node: carrier,
    styleResolver,
  }).filter(({ type }) => RUN_FORMATTING_MARK_NAMES.has(type.name));
  return {
    formatting: context.paragraphFormatting,
    marks: markInheritedFormattingAsKnownEmpty(marks, paragraph.type.schema),
  };
};

/**
 * Change paragraph attrs and re-resolve inherited run marks in the new style
 * context without turning the old rendered style into direct run formatting.
 */
export const setParagraphPropertiesWithRebasedRunFormatting = ({
  paragraphPosition,
  projection,
  previousTableRunFormatting,
  styleResolver,
  tableRunFormatting,
  tr,
}: RebaseParagraphRunFormattingOptions): Transaction => {
  const paragraph = tr.doc.nodeAt(paragraphPosition);
  if (!paragraph || paragraph.type.name !== "paragraph") {
    return panic("Cannot rebase run formatting outside a paragraph", {
      nodeType: paragraph?.type.name,
      paragraphPosition,
    });
  }

  const previousContext = paragraphRunStyleContext(
    paragraph,
    styleResolver,
    previousTableRunFormatting === undefined
      ? tableRunFormatting
      : previousTableRunFormatting,
  );
  const nextParagraph = createParagraphNodeFromProjection({
    type: paragraph.type,
    projection,
    content: paragraph.content,
    marks: paragraph.marks,
  });
  const nextContext = paragraphRunStyleContext(
    nextParagraph,
    styleResolver,
    tableRunFormatting,
  );
  const changes: {
    attrs: Readonly<Record<string, unknown>>;
    currentFormattingMarks: readonly Mark[];
    from: number;
    isText: boolean;
    marks: readonly Mark[];
    to: number;
  }[] = [];

  const collectRebasedRepresentation = ({
    node,
    position,
  }: RunFormattingCarrierRepresentation): void => {
    const authoredFormatting = readAuthoredRunFormatting({
      context: previousContext,
      marks: node.marks,
      styleResolver,
    });
    const nextMarks = reconcileRunFormattingMarks({
      authoredFormatting,
      context: nextContext,
      node,
      styleResolver,
    });
    if (Mark.sameSet(node.marks, nextMarks)) {
      return;
    }
    changes.push({
      attrs: node.attrs,
      currentFormattingMarks: node.marks.filter(({ type }) =>
        RUN_FORMATTING_MARK_NAMES.has(type.name),
      ),
      from: position,
      isText: node.isText,
      marks: nextMarks,
      to: position + node.nodeSize,
    });
  };

  paragraph.descendants((node, relativePosition) => {
    if (!node.isInline) {
      return true;
    }
    const carrier = expandRunFormattingCarrier(node, paragraphPosition + 1 + relativePosition);
    if (!carrier) {
      return !node.isAtom;
    }
    for (const representation of carrier.representations) {
      collectRebasedRepresentation(representation);
    }
    return false;
  });

  applyParagraphPropertyProjection({
    transaction: tr,
    pos: paragraphPosition,
    projection,
    source: { type: "preserve" },
  });
  for (const { attrs, currentFormattingMarks, from, isText, marks, to } of changes) {
    if (!isText) {
      applyNonParagraphMarkup({ transaction: tr, pos: from, attrs, marks });
      continue;
    }
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
