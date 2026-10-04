import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

const owners = new WeakMap<EditorView, (command: Command) => boolean | undefined>();

/** The view's controller owns canonical command interpretation. */
export const registerEditorCommandOwner = (
  view: EditorView,
  execute: (command: Command) => boolean | undefined,
) => {
  owners.set(view, execute);
  return () => {
    owners.delete(view);
  };
};

/** Keep commands on their selected view; a handled refusal must not fall through. */
export const executeEditorCommand = (view: EditorView, command: Command): boolean => {
  const handled = owners.get(view)?.(command);
  if (handled !== undefined) return handled;
  return command(view.state, (transaction) => view.dispatch(transaction), view);
};

/** Legacy command chains stop at the first owner decision, including refusal. */
export const executeFirstEditorCommand = (
  view: EditorView,
  commands: readonly Command[],
): boolean => {
  for (const command of commands) {
    const handled = owners.get(view)?.(command);
    if (handled !== undefined) return handled;
    if (command(view.state, (transaction) => view.dispatch(transaction), view)) return true;
  }
  return false;
};
