/**
 * A bookmark between a table's rows, or between a row's cells, has nowhere in
 * the schema to stand. The node carries it in `_bookmarks`, addressed by child
 * index, and that index has to follow the children as they come and go. Every
 * change to them passes here, whichever path made it.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorState, Transaction } from "prosemirror-state";
import { Mapping } from "prosemirror-transform";

type PositionedBookmark = { index: number; marker: unknown };

const positionedBookmarksOf = (
  node: PMNode,
): readonly PositionedBookmark[] | null => {
  const role = node.type.spec["tableRole"];
  if (role !== "table" && role !== "row") return null;
  const bookmarks: unknown = node.attrs["_bookmarks"];
  return Array.isArray(bookmarks) && bookmarks.length > 0
    ? (bookmarks as PositionedBookmark[])
    : null;
};

/** Where a boundary between the source's children falls among the new ones. */
const childIndexAt = (node: PMNode, nodePosition: number, at: number) => {
  let index = 0;
  let end = nodePosition + 1;
  while (index < node.childCount) {
    end += node.child(index).nodeSize;
    if (end > at) break;
    index += 1;
  }
  return index;
};

const followingIndex = ({
  source,
  sourcePosition,
  node,
  position,
  index,
  mapping,
}: {
  source: PMNode | null | undefined;
  sourcePosition: number;
  node: PMNode;
  position: number;
  index: number;
  mapping: Mapping;
}): number => {
  if (!source || source.type !== node.type || index > source.childCount) {
    return Math.min(index, node.childCount);
  }
  // After the last child it stays after the last child, whatever that is now.
  if (index === source.childCount) return node.childCount;
  let boundary = sourcePosition + 1;
  for (let child = 0; child < index; child += 1) {
    boundary += source.child(child).nodeSize;
  }
  return childIndexAt(node, position, mapping.map(boundary, -1));
};

const repairChanged = ({
  before,
  after,
  offset,
  mapping,
  back,
  sourceDoc,
  tr,
}: {
  before: PMNode;
  after: PMNode;
  offset: number;
  mapping: Mapping;
  back: Mapping;
  sourceDoc: PMNode;
  tr: Transaction;
}): void => {
  // Only what the transaction changed: a document's other tables cost nothing.
  after.forEach((child, childOffset, index) => {
    if (child === before.maybeChild(index) || child.isTextblock) return;
    const position = offset + childOffset;
    const bookmarks = positionedBookmarksOf(child);
    if (bookmarks) {
      const sourcePosition = back.map(position, 1);
      const source = sourceDoc.nodeAt(sourcePosition);
      const repaired = bookmarks.map((bookmark) => ({
        ...bookmark,
        index: followingIndex({
          source,
          sourcePosition,
          node: child,
          position,
          index: bookmark.index,
          mapping,
        }),
      }));
      if (repaired.some(({ index: at }, i) => at !== bookmarks[i]?.index)) {
        tr.setNodeAttribute(position, "_bookmarks", repaired);
      }
    }
    repairChanged({
      before: before.maybeChild(index) ?? child.type.create(),
      after: child,
      offset: position + 1,
      mapping,
      back,
      sourceDoc,
      tr,
    });
  });
};

export const tablePositionedBookmarksPlugin = (): Plugin =>
  new Plugin({
    key: new PluginKey("tablePositionedBookmarks"),
    appendTransaction(
      transactions: readonly Transaction[],
      oldState: EditorState,
      newState: EditorState,
    ) {
      if (!transactions.some(({ docChanged }) => docChanged)) return null;
      const mapping = new Mapping();
      for (const { mapping: each } of transactions) {
        mapping.appendMapping(each);
      }
      const tr = newState.tr;
      repairChanged({
        before: oldState.doc,
        after: newState.doc,
        offset: 0,
        mapping,
        back: mapping.invert(),
        sourceDoc: oldState.doc,
        tr,
      });
      if (!tr.docChanged) return null;
      return tr.setMeta("addToHistory", false);
    },
  });
