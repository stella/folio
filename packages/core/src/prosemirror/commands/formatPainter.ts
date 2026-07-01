/**
 * Format Painter commands — copy the character formatting of one selection and
 * apply ("paint") it onto another, mirroring a word processor's copy/paste
 * formatting (Ctrl+Shift+C / Ctrl+Shift+V).
 *
 * These are pure, headless commands: capture reads the marks at the current
 * selection, apply lays a captured mark set onto the current selection's range.
 * The armed/sticky toolbar interaction lives in the React layer; all editor-state
 * logic is here so it can be unit-tested without a DOM.
 */

import type { Mark } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

/**
 * Character-formatting marks the painter copies. Structural marks — comments,
 * hyperlinks, tracked changes (insertion/deletion), footnote references — are
 * deliberately excluded so painting never moves an anchor, link, or revision.
 * `rtl` (run direction) and `hidden` are excluded too: painting them could flip
 * text direction or hide content, which is surprising for a formatting brush.
 */
export const PAINTABLE_MARK_NAMES: ReadonlySet<string> = new Set([
  "bold",
  "italic",
  "underline",
  "strike",
  "textColor",
  "highlight",
  "fontSize",
  "fontFamily",
  "superscript",
  "subscript",
  "allCaps",
  "smallCaps",
  "characterSpacing",
  "runShading",
  "runFormattingOverride",
  "emboss",
  "imprint",
  "textShadow",
  "emphasisMark",
  "textOutline",
  "textEffect",
  "characterStyle",
]);

/**
 * Marks of the first text node inside [from, to). A word processor's format
 * brush copies from the start of the source selection, so a mixed selection
 * yields the formatting of its leading run.
 */
function firstTextMarks(state: EditorState, from: number, to: number): readonly Mark[] {
  let marks: readonly Mark[] | null = null;

  state.doc.nodesBetween(from, to, (node) => {
    if (marks) {
      return false;
    }
    if (node.isText) {
      marks = node.marks;
      return false;
    }
    return true;
  });

  return marks ?? [];
}

/**
 * Capture the paintable character marks active over the current selection.
 * Returns the marks (with their attrs) to hand to `applyFormatMarks`. An empty
 * selection reads the stored/insertion marks; a range reads its leading run.
 * Returns an empty array when the source carries no direct character formatting.
 */
export function captureFormatMarks(state: EditorState): Mark[] {
  const { empty, $from, from, to } = state.selection;
  const source = empty ? (state.storedMarks ?? $from.marks()) : firstTextMarks(state, from, to);
  return source.filter((mark) => PAINTABLE_MARK_NAMES.has(mark.type.name));
}

/**
 * Apply captured marks onto the current selection's range: clear every paintable
 * mark type first (replace, not merge — so painting Georgia-12-bold onto
 * Arial-10 leaves only Georgia-12-bold), then lay down the captured marks.
 *
 * Best-effort: a collapsed selection or an empty capture is a no-op (returns
 * false so a keymap binding stays unclaimed and the keystroke passes through).
 * Never throws.
 */
export function applyFormatMarks(marks: readonly Mark[]): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty || marks.length === 0) {
      return false;
    }

    if (!dispatch) {
      return true;
    }

    let tr = state.tr;
    for (const name of PAINTABLE_MARK_NAMES) {
      const markType = state.schema.marks[name];
      if (markType) {
        tr = tr.removeMark(from, to, markType);
      }
    }
    for (const mark of marks) {
      tr = tr.addMark(from, to, mark);
    }

    dispatch(tr.scrollIntoView());
    return true;
  };
}
