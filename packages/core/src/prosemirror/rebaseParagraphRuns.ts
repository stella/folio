import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import {
  applyMarksToRunFormattingRepresentation,
  selectRunFormattingCarrierRepresentations,
} from "./runFormattingInlineCarriers";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { paragraphRunStyleContext, type RunStyleResolver } from "./runStyleFormatting";

type RebaseParagraphRunsOptions = {
  tr: Transaction;
  position: number;
  previous: PMNode;
  styleResolver: RunStyleResolver | null;
};

/** Change the paragraph cascade while retaining each run's character style and authored properties. */
export const rebaseParagraphRuns = ({
  tr,
  position,
  previous,
  styleResolver,
}: RebaseParagraphRunsOptions): void => {
  const paragraph = tr.doc.nodeAt(position);
  if (!paragraph) return;
  const storedMarks = tr.storedMarks ?? tr.selection.$from.marks();
  const sourceContext = paragraphRunStyleContext(previous, styleResolver);
  const targetContext = paragraphRunStyleContext(paragraph, styleResolver);
  for (const representation of selectRunFormattingCarrierRepresentations({
    doc: tr.doc,
    from: position + 1,
    to: position + paragraph.nodeSize - 1,
  })) {
    const marks = reconcileRunFormattingMarks({
      authoredFormatting: readAuthoredRunFormatting({
        context: sourceContext,
        marks: representation.node.marks,
        styleResolver,
      }),
      context: targetContext,
      node: representation.node,
      styleResolver,
    });
    applyMarksToRunFormattingRepresentation({ tr, representation, marks });
  }
  if (
    tr.selection.empty &&
    tr.selection.from > position &&
    tr.selection.from < position + paragraph.nodeSize
  ) {
    const node = paragraph.type.schema.text(" ", storedMarks);
    tr.setStoredMarks(
      reconcileRunFormattingMarks({
        authoredFormatting: readAuthoredRunFormatting({
          context: sourceContext,
          marks: node.marks,
          styleResolver,
        }),
        context: targetContext,
        node,
        styleResolver,
      }),
    );
  }
};
