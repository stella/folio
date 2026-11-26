import type { Slice } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { createPluginKeyOwner } from "./pluginKeyOwner";
import { mintRevisionId } from "./plugins/revisionIds";

export const SUGGESTION_META = "suggestionModeApplied";
export const SUGGESTION_BYPASS_META = "suggestionModeBypass";

/** Text typed before list autoformat rules get to see it. */
export const SUGGESTED_TEXT_INPUT_META = "suggestionModeTextInput";

export type SuggestionMode = {
  active: boolean;
  author: string;
};

export type SuggestedRangeDeletion = {
  state: EditorState;
  tr: Transaction;
  from: number;
  to: number;
};

type SuggestionModeState = SuggestionMode & {
  commands: {
    paste: (view: EditorView, slice: Slice, mode: SuggestionMode) => boolean;
    deleteRange: (range: SuggestedRangeDeletion) => boolean;
    deleteSelection: (
      state: EditorState,
      dispatch: (tr: Transaction) => void,
    ) => boolean;
  };
};

// Only the reference editor installs this state and its command implementations.
export const suggestionModeKey =
  createPluginKeyOwner<SuggestionModeState>("suggestionMode");

export const isSuggestionModeActive = (state: EditorState): boolean =>
  suggestionModeKey.getState(state)?.active ?? false;

export type SuggestionMarkAttrs = {
  revisionId: number;
  author: string;
  date: string;
};

export const makeMarkAttrs = ({
  author,
}: SuggestionMode): SuggestionMarkAttrs => ({
  revisionId: mintRevisionId(),
  author,
  date: new Date().toISOString(),
});

export const makeRevisionInfo = (state: EditorState) => {
  const mode = suggestionModeKey.getState(state);
  if (!mode?.active) return null;
  const { revisionId, author, date } = makeMarkAttrs(mode);
  return { id: revisionId, author, date };
};

export const pasteAsSuggestion = (view: EditorView, slice: Slice): boolean => {
  const mode = suggestionModeKey.getState(view.state);
  return mode?.active ? mode.commands.paste(view, slice, mode) : false;
};

export const suggestRangeDeletion = (range: SuggestedRangeDeletion): boolean =>
  suggestionModeKey.getState(range.state)?.commands.deleteRange(range) ?? false;

export const deleteSelectionAsSuggestion = (
  state: EditorState,
  dispatch: (tr: Transaction) => void,
): boolean =>
  suggestionModeKey
    .getState(state)
    ?.commands.deleteSelection(state, dispatch) ?? false;
