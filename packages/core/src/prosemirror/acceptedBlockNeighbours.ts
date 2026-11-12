/**
 * What a place between blocks has next to it once every pending change is
 * accepted.
 *
 * Accepting removes some blocks whole: a paragraph with nothing left of it
 * whose mark is deleted (`paragraphGoes`), and a table every row of which is
 * deleted (`tableGoes`). An edit reads a document with pending changes as
 * accepting leaves it: a join into a paragraph accepting removes, right
 * before a table, ends there, and the two sides of blocks accepting removes
 * whole are one gap (`insertionGap`). Applied directly, the same edit reads
 * the same.
 */

import type { Node as PMNode } from "prosemirror-model";

import { expectParagraphAttrs } from "./attrs";
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
 * What accepting every change ends a pending join at: the paragraph whose
 * break the joined words take, or a table that stays and stops them, as
 * `joinEndsBeforeTable` reads it.
 */
export type PendingJoinEnd = {
  kind: "paragraph" | "table";
  position: number;
};

/**
 * Where accepting every change ends the join from the paragraph at
 * `position`, whose break is a plain pending deletion while something of it
 * stays. The join runs on past the paragraphs accepting removes whole and
 * takes the break of the next paragraph, when that break stays; a table that
 * stays after at least one removed paragraph stops it, and the words keep
 * their own paragraph. Null for any other end: a paragraph whose break goes
 * as well (the longer join is read from its last paragraph), a table right
 * after the words, or the end of the container.
 */
export const pendingJoinEnd = (
  doc: PMNode,
  position: number,
): PendingJoinEnd | null => {
  const paragraph = doc.nodeAt(position);
  if (
    paragraph?.type.name !== "paragraph" ||
    expectParagraphAttrs(paragraph).pPrMark?.kind !== "del" ||
    paragraphGoes(paragraph, "accept")
  ) {
    return null;
  }
  const $after = doc.resolve(position + paragraph.nodeSize);
  let at = $after.pos;
  for (let index = $after.index(); index < $after.parent.childCount; index++) {
    const next = $after.parent.child(index);
    if (next.type.name === "paragraph" && paragraphGoes(next, "accept")) {
      at += next.nodeSize;
      continue;
    }
    if (next.type.name === "paragraph") {
      return expectParagraphAttrs(next).pPrMark == null
        ? { kind: "paragraph", position: at }
        : null;
    }
    return isTable(next) && at > $after.pos && !tableGoes(next, "accept")
      ? { kind: "table", position: at }
      : null;
  }
  return null;
};
