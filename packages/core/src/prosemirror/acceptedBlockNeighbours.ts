/**
 * What a place between blocks has next to it once every pending change is
 * accepted.
 *
 * Accepting removes some blocks whole: a paragraph with nothing left of it
 * whose mark is deleted (`paragraphGoes`), and a table every row of which is
 * deleted (`tableGoes`). An edit reads a document with pending changes as
 * accepting leaves it, so it reads past those blocks: a merge joins what
 * stands beyond them, and the two sides of a table accepting removes are one
 * gap. Applied directly after those blocks are gone, the same edit meets the
 * same neighbours.
 */

import type { Node as PMNode } from "prosemirror-model";

import { paragraphJoinAfter } from "./commands/paragraphBookmarkJoin";
import { isBlockBoundaryNode } from "./keptRangeBoundaries";
import { paragraphGoes, tableGoes } from "./tableRunIn";

const isTable = (node: PMNode): boolean =>
  node.type.spec["tableRole"] === "table";

/**
 * The join a merge of the paragraph ending at `position` makes with the
 * paragraph after it, or null where it has no paragraph to join: the next
 * block is no paragraph (`paragraphJoinAfter`), or accepting removes every
 * block from that paragraph on up to a table that stays. Accepted, such a join
 * runs the merged paragraph's words on into the table's first cell, which a
 * merge right before a table never does, so it is refused in every mode alike.
 * Paragraphs that accepting removes up to the end of the container leave the
 * merged paragraph last, as before.
 */
export const paragraphMergeJoinAfter = (
  doc: PMNode,
  position: number,
): ReturnType<typeof paragraphJoinAfter> => {
  const join = paragraphJoinAfter(doc, position);
  if (!join) return null;
  const $next = doc.resolve(join.nextPosition);
  const { parent } = $next;
  for (let index = $next.index(); index < parent.childCount; index++) {
    const block = parent.child(index);
    if (isBlockBoundaryNode(block)) continue;
    if (isTable(block)) {
      if (tableGoes(block, "accept")) continue;
      return null;
    }
    if (!paragraphGoes(block, "accept")) return join;
  }
  return join;
};

/**
 * Where a block put at `position`, a boundary between blocks, lands: before
 * every table right before it that accepting removes whole. Accepted, the two
 * sides of such a table are one gap, so a block put on either side of it lands
 * at the gap's start, and the blocks a batch puts into that gap keep the order
 * the batch names them in, as they do once the table is removed directly.
 */
export const insertionGapStart = (doc: PMNode, position: number): number => {
  let at = position;
  for (;;) {
    const before = doc.resolve(at).nodeBefore;
    if (!before || !isTable(before) || !tableGoes(before, "accept")) return at;
    at -= before.nodeSize;
  }
};
