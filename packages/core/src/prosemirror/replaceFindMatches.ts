/**
 * Find and replace writes through the editor that shows the document.
 *
 * Every adapter replaces matches with one editor command run on the body
 * view's command owner: a canonical session turns it into text replacement
 * intents in its own journal (tracked when the session suggests, one history
 * step), and the default session applies the command's transaction, which its
 * suggestion plugin tracks the same way as typing.
 */

import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import {
  type FindMatch,
  type ReplaceAllSummary,
  replaceEligibility,
} from "../utils/findReplace";
import { withCanonicalCommand } from "./canonicalCommands";
import { executeEditorCommand } from "./executeEditorCommand";
import { resolveFindMatchRanges } from "./findReplaceSelection";

type TextRange = { from: number; to: number };

/**
 * Replace each body range with `text` in one step. Ranges run from the end of
 * the document backwards, so each replacement leaves the ones before it at
 * their positions.
 */
const replaceTextRanges = (
  ranges: readonly TextRange[],
  text: string,
): Command =>
  withCanonicalCommand(
    (state, dispatch) => {
      if (ranges.length === 0) return false;
      if (dispatch) {
        const transaction = state.tr;
        for (const { from, to } of ranges)
          transaction.insertText(text, from, to);
        dispatch(transaction.scrollIntoView());
      }
      return true;
    },
    () =>
      ranges.map(({ from, to }) => ({ type: "replaceText", from, to, text })),
  );

/**
 * Replace `matches` with `text` through the body view's command owner. Only
 * matches replacement policy allows are written; a match that already reads
 * `text` counts as replaced without an edit. Nothing is replaced when the
 * owner refuses the edit.
 */
export const replaceFindMatches = (
  view: EditorView,
  matches: readonly FindMatch[],
  text: string,
): ReplaceAllSummary => {
  const allowed = matches.filter(
    (match) => replaceEligibility(match).status === "allowed",
  );
  const unchangedCount = allowed.filter((match) => match.text === text).length;
  const ranges = resolveFindMatchRanges(
    view.state.doc,
    allowed.filter((match) => match.text !== text),
  )
    .filter((range) => range !== null)
    .map(({ from, to }) => ({ from, to }))
    .toSorted((left, right) => right.from - left.from);
  const unwritten = {
    replacedCount: unchangedCount,
    skippedCount: matches.length - unchangedCount,
  };
  if (ranges.length === 0) return unwritten;
  if (!executeEditorCommand(view, replaceTextRanges(ranges, text)))
    return unwritten;
  const replacedCount = ranges.length + unchangedCount;
  return { replacedCount, skippedCount: matches.length - replacedCount };
};

/** Replace one match; true when the match reads `text` afterwards. */
export const replaceFindMatch = (
  view: EditorView,
  match: FindMatch,
  text: string,
): boolean => replaceFindMatches(view, [match], text).replacedCount === 1;
