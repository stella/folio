/** Framework-neutral arbitration for the editor view that receives commands. */

import type { EditorView } from "prosemirror-view";

export type ActiveEditorStoryTarget =
  | { type: "note"; view: EditorView }
  | { type: "headerFooter"; view: EditorView }
  | { type: "body"; view: EditorView }
  | { type: "none"; view: null };

export type ActiveEditorStoryCandidates = {
  bodyView: EditorView | null;
  headerFooterView: EditorView | null;
  noteView: EditorView | null;
};

const NO_ACTIVE_STORY: ActiveEditorStoryTarget = { type: "none", view: null };

/**
 * Resolve the single command target across editable document stories.
 *
 * A visible note panel takes precedence over header/footer mode, which takes
 * precedence over the body. Keeping this order in core prevents framework
 * adapters from sending formatting, focus, or history commands to a story
 * hidden behind another active editor surface during an async UI transition.
 */
export const resolveActiveEditorStory = ({
  bodyView,
  headerFooterView,
  noteView,
}: ActiveEditorStoryCandidates): ActiveEditorStoryTarget => {
  if (noteView) {
    return { type: "note", view: noteView };
  }
  if (headerFooterView) {
    return { type: "headerFooter", view: headerFooterView };
  }
  if (bodyView) {
    return { type: "body", view: bodyView };
  }
  return NO_ACTIVE_STORY;
};
