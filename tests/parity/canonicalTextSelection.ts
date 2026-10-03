import type { EditorState } from "prosemirror-state";

/** Text slices use document positions, which include structural node boundaries. */
export const canonicalTextSelection = ({
  doc,
  selection,
}: Pick<EditorState, "doc" | "selection">) => ({
  before: doc.textBetween(0, selection.from, "", ""),
  selected: doc.textBetween(selection.from, selection.to, "", ""),
  after: doc.textBetween(selection.to, doc.content.size, "", ""),
});
