import { EditorState, TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";

import { isInlineAnchor } from "../extensions/nodes/inlineAnchorDom";

/** The length of the first or last user-perceived character of `text`. */
const characterLength = (text: string, backward: boolean): number => {
  const segments = [...new Intl.Segmenter().segment(text)];
  const segment = backward ? segments.at(-1) : segments.at(0);
  return segment?.segment.length ?? 1;
};

/**
 * Delete what the caret reaches past the inline anchors beside it.
 *
 * Left to the browser, the key removes the anchor's element, and the next DOM
 * read loses the node: a bookmark or comment range, with the other half of its
 * pair, while the text stays. Suggesting mode tracks the key before it gets here.
 */
export const deleteBesideInlineAnchor =
  (join: Command, direction: "backward" | "forward"): Command =>
  (state, dispatch, view) => {
    const { selection, doc } = state;
    if (!(selection instanceof TextSelection) || !selection.empty) {
      return false;
    }
    const backward = direction === "backward";
    const adjacent = (position: number) => {
      const $position = doc.resolve(position);
      return backward ? $position.nodeBefore : $position.nodeAfter;
    };
    let cursor = selection.from;
    for (
      let node = adjacent(cursor);
      node && isInlineAnchor(node);
      node = adjacent(cursor)
    ) {
      cursor += backward ? -node.nodeSize : node.nodeSize;
    }
    if (cursor === selection.from) return false;

    const next = adjacent(cursor);
    if (!next) {
      // The paragraph's edge: join from there, as the key does beside no anchor.
      const atEdge = EditorState.create({
        doc,
        selection: TextSelection.create(doc, cursor),
      });
      return join(atEdge, dispatch, view);
    }
    const size = next.isText
      ? characterLength(next.text ?? "", backward)
      : next.nodeSize;
    dispatch?.(
      state.tr
        .delete(
          backward ? cursor - size : cursor,
          backward ? cursor : cursor + size,
        )
        .scrollIntoView(),
    );
    return true;
  };
