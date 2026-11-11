/**
 * What a place between blocks has next to it once every pending change is
 * accepted.
 *
 * Accepting removes some blocks whole: a paragraph with nothing left of it
 * whose mark is deleted (`paragraphGoes`), and a table every row of which is
 * deleted (`tableGoes`). An edit reads a document with pending changes as
 * accepting leaves it: a join into a paragraph accepting removes, right
 * before a table, ends there, and the two sides of a table accepting removes
 * are one gap. Applied directly, the same edit reads the same.
 */

import type { Node as PMNode } from "prosemirror-model";

import { paragraphLeftAfter } from "./paragraphPropertyCarry";
import { paragraphGoes, tableGoes } from "./tableRunIn";

const isTable = (node: PMNode): boolean =>
  node.type.spec["tableRole"] === "table";

/**
 * Whether accepting every change ends a join at the paragraph at `position`,
 * leaving the words joined into it apart from the table right after it:
 * nothing of that paragraph stays and its mark goes, and a table that stays
 * follows it. A deleted paragraph stood before the table, and the join ended
 * there (`tableRunIn`), so a merge into it applied directly keeps its break
 * before the table, as the same merge tracked and accepted leaves it.
 */
export const joinEndsBeforeTable = (doc: PMNode, position: number): boolean => {
  const paragraph = doc.nodeAt(position);
  if (!paragraph || !paragraphGoes(paragraph, "accept")) return false;
  const after = doc.resolve(position + paragraph.nodeSize).nodeAfter;
  return after !== null && isTable(after) && !tableGoes(after, "accept");
};

/**
 * The paragraph accepting every change leaves of a merge from the paragraph
 * at `paragraphPos`, whose break goes: the one `paragraphLeftAfter` finds,
 * unless the join ends at it before a table (`joinEndsBeforeTable`), where
 * accepting leaves the paragraph the merge started from. The merge's
 * properties go there.
 */
export const paragraphMergeLeaves = (
  options: Parameters<typeof paragraphLeftAfter>[0],
): number | null => {
  const left = paragraphLeftAfter(options);
  return left !== null && joinEndsBeforeTable(options.doc, left)
    ? options.paragraphPos
    : left;
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
