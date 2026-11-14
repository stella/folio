/**
 * Replace a document range with plain text, as typing over it would: menu
 * actions such as cut, delete and delete image run this through the view's
 * command owner, so a canonical session takes it into its journal and a
 * suggesting session records it as a suggestion.
 */

import type { Command } from "prosemirror-state";

import { withCanonicalCommand } from "../canonicalCommands";
import { deleteSelectionAsSuggestion } from "../plugins/suggestionMode";

export type RangeTextReplacement = {
  from: number;
  to: number;
  text: string;
};

export const replaceRangeText = ({
  from,
  to,
  text,
}: RangeTextReplacement): Command =>
  withCanonicalCommand(
    (state, dispatch) => {
      if (from > to || to > state.doc.content.size) return false;
      if (from === to && text === "") return false;
      if (!dispatch) return true;
      const { selection } = state;
      // A deletion of the selection is what the keyboard's cut deletes,
      // tracked when suggesting.
      if (
        text === "" &&
        selection.from === from &&
        selection.to === to &&
        deleteSelectionAsSuggestion(state, dispatch)
      )
        return true;
      dispatch(
        (text === ""
          ? state.tr.delete(from, to)
          : state.tr.insertText(text, from, to)
        ).scrollIntoView(),
      );
      return true;
    },
    () =>
      from === to && text === ""
        ? []
        : [{ type: "replaceText", from, to, text }],
  );
