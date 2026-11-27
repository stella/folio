/**
 * The blocks of the hidden body editor the browser lays out.
 *
 * The pages paint the document, so the hidden editor needs layout only where
 * the browser edits natively: around the selection. Each top-level block the
 * selection reaches, and the block before and after them, carries
 * {@link HIDDEN_EDITOR_LAID_OUT_CLASS}; the adapters' stylesheets let the
 * browser skip the layout of every other block (`content-visibility: auto`).
 * The class lands in the same view update as the selection, so the browser
 * never deletes, composes or moves the caret in a block it has skipped.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";

import { createPluginKeyOwner } from "../pluginKeyOwner";

/** The class of a hidden editor block the browser lays out. */
export const HIDDEN_EDITOR_LAID_OUT_CLASS = "folio-hidden-laid-out";

type LayoutWindow = {
  /** Index of the first top-level block laid out. */
  first: number;
  /** Index of the last top-level block laid out. */
  last: number;
  decorations: DecorationSet;
};

const layoutWindowOwner = createPluginKeyOwner<LayoutWindow>(
  "hiddenEditorLayoutWindow",
);

const LAID_OUT = { class: HIDDEN_EDITOR_LAID_OUT_CLASS };

/** The index of the top-level block holding, or following, `position`. */
const blockIndexAt = (doc: PMNode, position: number): number =>
  doc.resolve(position).index(0);

const windowOf = (
  state: EditorState,
  previous: LayoutWindow | undefined,
  docChanged: boolean,
): LayoutWindow => {
  const { doc } = state;
  if (doc.childCount === 0)
    return { first: 0, last: -1, decorations: DecorationSet.empty };
  let from = state.selection.from;
  let to = state.selection.to;
  for (const range of state.selection.ranges) {
    from = Math.min(from, range.$from.pos);
    to = Math.max(to, range.$to.pos);
  }
  const $from = doc.resolve(from);
  const first = Math.max(0, $from.index(0) - 1);
  const last = Math.min(doc.childCount - 1, blockIndexAt(doc, to) + 1);
  if (
    !docChanged &&
    previous !== undefined &&
    previous.first === first &&
    previous.last === last
  )
    return previous;
  // Where the block at `first` starts, stepping back from the one at `from`.
  let index = $from.index(0);
  let position = $from.depth > 0 ? $from.before(1) : from;
  while (index > first) {
    index -= 1;
    position -= doc.child(index).nodeSize;
  }
  const decorations: Decoration[] = [];
  for (; index <= last; index += 1) {
    const block = doc.child(index);
    decorations.push(
      Decoration.node(position, position + block.nodeSize, LAID_OUT),
    );
    position += block.nodeSize;
  }
  return { first, last, decorations: DecorationSet.create(doc, decorations) };
};

/** Mark the hidden body editor's blocks around the selection as laid out. */
export const createHiddenEditorLayoutWindowPlugin =
  (): Plugin<LayoutWindow> => {
    const key = layoutWindowOwner.createKey();
    return new Plugin<LayoutWindow>({
      key,
      state: {
        init: (_config, state) => windowOf(state, undefined, true),
        apply: (transaction, previous, _before, state) =>
          transaction.docChanged || transaction.selectionSet
            ? windowOf(state, previous, transaction.docChanged)
            : previous,
      },
      props: {
        decorations: (state) => key.getState(state)?.decorations,
      },
    });
  };
