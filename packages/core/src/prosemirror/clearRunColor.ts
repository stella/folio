import type { Node as PMNode } from "prosemirror-model";
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
  const clearedMarks = (node: PMNode, position: number) => {
    const context = paragraphRunStyleContextAt({
      doc: state.doc,
      pos: position,
      styleResolver,
    });
    const { color: _color, ...authoredFormatting } = readAuthoredRunFormatting({
      context,
      marks: node.marks,
      styleResolver,
    });
    return reconcileRunFormattingMarks({
      authoredFormatting,
      context,
      node,
      styleResolver,
    });
  };
  if (empty) {
    const node = state.schema.text(" ", state.storedMarks ?? state.selection.$from.marks());
    tr.setStoredMarks(clearedMarks(node, from));
  } else {
    for (const representation of selectRunFormattingCarrierRepresentations({
      doc: state.doc,
      from,
      to,
    })) {
      const marks = clearedMarks(representation.node, representation.position);
      applyMarksToRunFormattingRepresentation({ tr, representation, marks });
    }
  }
  dispatch(tr.scrollIntoView());
  return true;
};
