/**
 * Paste without formatting
 *
 * Inserts clipboard text as unformatted paragraphs, dropping every source mark
 * (bold, colour, font, links) so the pasted text takes on the destination
 * paragraph's own style. This is the "paste without formatting" escape hatch,
 * distinct from the default rich paste that preserves source formatting.
 */

import { Fragment, Slice } from "prosemirror-model";
import type { Schema } from "prosemirror-model";
import type { Command } from "prosemirror-state";

const PARAGRAPH_BREAK = /(?:\r\n?|\n)+/;

/**
 * Build a mark-free slice from plain text, one paragraph per run of newlines.
 *
 * Pure and DOM-free so it is unit-testable and reusable from both the keymap
 * command and the context menu. The slice is opened at both ends
 * ({@link Slice.maxOpen}) so its first and last paragraphs merge into the
 * surrounding block on {@link Transaction.replaceSelection}, matching how the
 * editor's default text paste flows into the current line.
 */
export function buildPlainTextSlice(text: string, schema: Schema): Slice {
  const paragraphType = schema.nodes["paragraph"];
  if (!paragraphType) {
    // No paragraph node (e.g. an inline-only schema): fall back to flat text.
    return text ? new Slice(Fragment.from(schema.text(text)), 0, 0) : Slice.empty;
  }

  const paragraphs = text
    .split(PARAGRAPH_BREAK)
    .map((block) => paragraphType.create(null, block ? schema.text(block) : null));

  return Slice.maxOpen(Fragment.fromArray(paragraphs));
}

/**
 * Read the clipboard and insert it as unformatted text at the current selection.
 *
 * Dry-runnable per the ProseMirror command convention: applicability does not
 * depend on the (asynchronously read) clipboard contents, so the command is
 * "available" whenever the runtime exposes a clipboard reader. A probe with no
 * `dispatch` (menus/toolbars) returns `true` without touching the clipboard or
 * the document. Only when `dispatch` is provided does it read the clipboard and
 * insert; it then returns `true` (claiming the key) so the browser's own
 * "paste without formatting" shortcut does not also fire. The read is
 * best-effort: a denied or unavailable clipboard is swallowed rather than thrown.
 */
export const pasteWithoutFormatting: Command = (state, dispatch, view) => {
  if (typeof navigator === "undefined" || !navigator.clipboard?.readText) {
    return false;
  }
  // Dry run: report availability without side effects.
  if (!dispatch) {
    return true;
  }
  // Executing needs the view so the async clipboard result dispatches against
  // fresh state once it resolves; a captured stale state would misplace it.
  if (!view) {
    return false;
  }

  const { schema } = state;
  void navigator.clipboard
    .readText()
    .then((text) => {
      // The clipboard read is async: the editor may have been torn down while
      // it was in flight, so re-check before dispatching against a dead view.
      if (!text || view.isDestroyed) {
        return;
      }
      const slice = buildPlainTextSlice(text, schema);
      view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView());
    })
    .catch(() => {
      // Best-effort: clipboard read can be denied or unsupported.
    });

  return true;
};
