/**
 * Pure ref-API query helpers — read-only inspectors over the PM document.
 * Back the adapters' `highlightRange` / comment / tracked-change ref
 * methods.
 *
 * Every function takes the PM view (or doc) as a parameter instead of
 * closing over a framework ref, so the React and Vue adapters (and the
 * future vanilla wrapper) share one implementation.
 */

import type { EditorView } from "prosemirror-view";
import type { Node as ProseMirrorNode } from "prosemirror-model";
import { extractTrackedChanges } from "./utils/extractTrackedChanges";

/** A resolved PM position range — half-open `[from, to)` in PM coordinates. */
export type PmRange = {
  from: number;
  to: number;
}

/**
 * Clamp a caller-supplied `[from, to]` range to a valid in-document span, or
 * return `null` when it cannot be made valid: non-integer, negative, reversed
 * (`to < from`), or a `from` past the document end. `to` is clamped to the
 * document size so an out-of-range end never makes `doc.resolve()` throw a
 * `RangeError`. Both adapters' `highlightRange` route raw caller positions
 * through this so the no-op contract holds identically.
 */
export function clampRangeToDoc(doc: ProseMirrorNode, from: number, to: number): PmRange | null {
  if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || to < from) return null;
  const max = doc.content.size;
  if (from > max) return null;
  return { from, to: Math.min(to, max) };
}

/**
 * Resolve a `commentId` to the PM position range its `comment` mark
 * spans. Walks every inline node carrying a `comment` mark with the
 * matching id and returns the union range (earliest start → latest end),
 * so a comment whose range is interrupted by un-marked inline atoms still
 * resolves to a single span. Returns `null` when the id is no longer
 * present (the comment was deleted, or the marked text was removed) — the
 * caller distinguishes "scrolled" from "stale" on that signal.
 *
 * Pure read over `view.state`; no dispatch.
 */
export function findCommentRange(view: EditorView | null, commentId: number): PmRange | null {
  if (!view) return null;
  const commentType = view.state.schema.marks["comment"];
  if (!commentType) return null;
  let from = Infinity;
  let to = -Infinity;
  view.state.doc.descendants((node, pos) => {
    if (!node.isInline) return;
    for (const mark of node.marks) {
      if (mark.type === commentType && mark.attrs["commentId"] === commentId) {
        from = Math.min(from, pos);
        to = Math.max(to, pos + node.nodeSize);
      }
    }
  });
  if (to < 0) return null;
  return { from, to };
}

/**
 * Resolve a tracked-change `revisionId` to the PM position range of its
 * first site. Delegates to {@link extractTrackedChanges} so a coalesced
 * revision (sites scattered across paragraphs, replace pairs, Enter
 * chains) resolves to the same entry the sidebar shows. Matches on the
 * entry's primary `revisionId`, its `insertionRevisionId`, or any
 * `coalescedRevisionIds` member. Returns `null` when no entry carries the
 * id (the change was accepted/rejected/deleted) — the caller uses this to
 * show a "location no longer exists" affordance instead of a silent
 * no-op.
 *
 * Pure read over `view.state`; no dispatch.
 */
export function findChangeRange(view: EditorView | null, revisionId: number): PmRange | null {
  if (!view) return null;
  const { entries } = extractTrackedChanges(view.state);
  const entry = entries.find(
    (e) =>
      e.revisionId === revisionId ||
      e.insertionRevisionId === revisionId ||
      e.coalescedRevisionIds?.includes(revisionId),
  );
  if (!entry) return null;
  return { from: entry.from, to: entry.to };
}
