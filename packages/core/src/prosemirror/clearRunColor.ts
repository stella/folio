import type { Command } from "prosemirror-state";

import { getDocumentStyleResolver } from "./plugins/documentStyles";
import {
  applyMarksToRunFormattingRepresentation,
  selectRunFormattingCarrierRepresentations,
} from "./runFormattingInlineCarriers";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "./runFormattingReconciliation";
import { paragraphRunStyleContextAt } from "./runStyleFormatting";

/** Clearing direct color reveals the run's character/paragraph style color. */
export const clearRunColor = (): Command => (state, dispatch) => {
  if (!dispatch) return true;
  const styleResolver = getDocumentStyleResolver(state);
  const tr = state.tr;
  const { from, to, empty } = state.selection;
  const representations = empty
    ? [
        {
          node: state.schema.text(" ", state.storedMarks ?? state.selection.$from.marks()),
          position: from,
          from,
          to,
        },
      ]
    : selectRunFormattingCarrierRepresentations({ doc: state.doc, from, to });
  for (const representation of representations) {
    const context = paragraphRunStyleContextAt({
      doc: state.doc,
      pos: representation.position,
      styleResolver,
    });
    const { color: _color, ...authoredFormatting } = readAuthoredRunFormatting({
      context,
      marks: representation.node.marks,
      styleResolver,
    });
    const marks = reconcileRunFormattingMarks({
      authoredFormatting,
      context,
      node: representation.node,
      styleResolver,
    });
    if (empty) tr.setStoredMarks(marks);
    else applyMarksToRunFormattingRepresentation({ tr, representation, marks });
  }
  dispatch(tr.scrollIntoView());
  return true;
};
