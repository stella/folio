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
 * Returns `true` (claiming the key) whenever a clipboard read can be attempted,
 * so the browser's own "paste without formatting" shortcut does not also fire.
 * The read is best-effort: a denied or unavailable clipboard is swallowed
 * rather than thrown.
 */
export const pasteWithoutFormatting: Command = (state, dispatch, view) => {
  if (!dispatch || !view) {
    return false;
  }
  if (typeof navigator === "undefined" || !navigator.clipboard?.readText) {
    return false;
  }

  const { schema } = state;
  void navigator.clipboard
    .readText()
    .then((text) => {
      if (!text) {
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
