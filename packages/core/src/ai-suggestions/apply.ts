/**
 * Apply AI suggestions to the document.
 *
 * Two modes:
 *  - "direct": replace text in place.
 *  - "tracked-changes": insertion + deletion marks, attributed to the
 *    configured author. Reuses the editor's existing tracked-change
 *    schema and resolution commands.
 *
 * Apply is always preceded by anchor resolution; stale suggestions are
 * skipped and reported back to the caller.
 */

import type { EditorView } from "prosemirror-view";

import { addTrackedDeletionMark } from "../prosemirror/addTrackedDeletionMark";
import {
  mintRevisionId,
  seedRevisionIdsFromDoc,
} from "../prosemirror/plugins/revisionIds";
import { resolveSuggestionAnchor } from "./conflict";
import type { AISuggestion, AISuggestionApplyMode } from "./types";

export type ApplyResult = {
  applied: string[];
  stale: string[];
};

type ApplyOptions = {
  view: Pick<EditorView, "state" | "dispatch">;
  suggestions: AISuggestion[];
  mode: AISuggestionApplyMode;
  author: string;
};

export function applySuggestions(options: ApplyOptions): ApplyResult {
  const { view, suggestions, mode, author } = options;
  const applied: string[] = [];
  const stale: string[] = [];

  if (suggestions.length === 0) {
    return { applied, stale };
  }

  const resolved: { suggestion: AISuggestion; from: number; to: number }[] = [];
  for (const suggestion of suggestions) {
    const anchor = resolveSuggestionAnchor(view.state.doc, suggestion);
    if (!anchor) {
      stale.push(suggestion.id);
      continue;
    }
    resolved.push({ suggestion, from: anchor.from, to: anchor.to });
  }

  if (resolved.length === 0) {
    return { applied, stale };
  }

  // Applied last to first, so each edit leaves the positions before it
  // intact. A range reaching into one already taken would be applied at
  // coordinates the earlier edit moved: it is reported stale instead.
  // At one start, the wider range first: an insertion there then lands
  // before the replaced text instead of overlapping it.
  resolved.sort((a, b) => b.from - a.from || b.to - a.to);
  const accepted: typeof resolved = [];
  for (const entry of resolved) {
    const following = accepted.at(-1);
    if (following !== undefined && entry.to > following.from) {
      stale.push(entry.suggestion.id);
      continue;
    }
    accepted.push(entry);
  }

  const tr = view.state.tr;
  if (mode === "direct") {
    for (const { suggestion, from, to } of accepted) {
      tr.insertText(suggestion.suggestedText, from, to);
      applied.push(suggestion.id);
    }
  } else {
    const insertionType = view.state.schema.marks["insertion"];
    const deletionType = view.state.schema.marks["deletion"];
    if (!insertionType || !deletionType) {
      return {
        applied,
        stale: [...stale, ...accepted.map((r) => r.suggestion.id)],
      };
    }
    const date = new Date().toISOString();
    seedRevisionIdsFromDoc(view.state.doc);

    for (const { suggestion, from, to } of accepted) {
      const revisionId = mintRevisionId();
      const attrs = { revisionId, author, date };

      if (suggestion.suggestedText.length > 0) {
        tr.insertText(suggestion.suggestedText, to, to);
        tr.addMark(
          to,
          to + suggestion.suggestedText.length,
          insertionType.create(attrs),
        );
      }
      if (to > from) {
        addTrackedDeletionMark({
          insertionPolicy: "preserve-pending",
          tr,
          from,
          to,
          mark: deletionType.create(attrs),
        });
      }
      applied.push(suggestion.id);
    }
  }

  if (tr.docChanged) {
    view.dispatch(tr);
  }
  return { applied, stale };
}
