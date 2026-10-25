import { NodeSelection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

export const selectLockedInlineControl = (
  view: Pick<EditorView, "dispatch" | "state">,
  event: KeyboardEvent,
): boolean => {
  if (event.key !== "Backspace" && event.key !== "Delete") return false;
  const { state } = view;
  if (!(state.selection instanceof TextSelection) || !state.selection.empty) {
    return false;
  }
  const backward = event.key === "Backspace";
  const { $from } = state.selection;
  const node = backward ? $from.nodeBefore : $from.nodeAfter;
  if (
    !node ||
    node.type.name !== "sdt" ||
    node.attrs["lock"] !== "contentLocked"
  ) {
    return false;
  }
  const pos = backward ? $from.pos - node.nodeSize : $from.pos;
  view.dispatch(
    state.tr
      .setSelection(NodeSelection.create(state.doc, pos))
      .scrollIntoView(),
  );
  return true;
};
