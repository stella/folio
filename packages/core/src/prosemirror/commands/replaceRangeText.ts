/**
 * Replace a document range with plain text, as typing over it would: menu
 * actions such as cut, delete and delete image run this through the view's
 * command owner, so a canonical session takes it into its journal and a
 * suggesting session records it as a suggestion.
 */

import type { Command } from "prosemirror-state";

import { withCanonicalCommand } from "../canonicalCommands";
import { deleteSelectionAsSuggestion } from "../suggestionCommands";

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
        text === ""
          ? state.tr.delete(from, to)
          : state.tr.insertText(text, from, to),
      );
      return true;
    },
    () =>
      from === to && text === ""
        ? []
        : [{ type: "replaceText", from, to, text }],
  );

/**
 * Replace several ranges with the same text in one step, as find and replace
 * does. The ranges must not overlap and run from the end of the document
 * backwards, so each replacement leaves the ones before it in place.
 */
export const replaceRangesText = (
  ranges: readonly { from: number; to: number }[],
  text: string,
): Command => {
  const [only] = ranges;
  if (only !== undefined && ranges.length === 1)
    return replaceRangeText({ ...only, text });
  return withCanonicalCommand(
    (state, dispatch) => {
      if (
        ranges.length === 0 ||
        ranges.some(({ from, to }) => from > to || to > state.doc.content.size)
      )
        return false;
      if (!dispatch) return true;
      const transaction = state.tr;
      for (const { from, to } of ranges)
        if (text === "") transaction.delete(from, to);
        else transaction.insertText(text, from, to);
      dispatch(transaction);
      return true;
    },
    () =>
      ranges.map(({ from, to }) => ({
        type: "replaceText" as const,
        from,
        to,
        text,
      })),
  );
};
