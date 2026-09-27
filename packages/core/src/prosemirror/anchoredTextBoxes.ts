/**
 * A text box and the paragraph it is drawn in.
 *
 * In DOCX a text box is run content: a `w:drawing` in a run of its paragraph.
 * The editor lifts it out into a `textBox` block after that paragraph and
 * leaves a zero-width `textBoxAnchor` where the run was; the save puts the box
 * back at its anchor, wherever the block has ended up. So the anchor, not the
 * block's position, says which paragraph owns the box, and two rules follow:
 *
 * - The box follows its anchor. A paragraph deleted, or its anchor deleted,
 *   takes the box with it: a drawing is content of the paragraph it is in.
 *   Left behind, the block was saved into whatever paragraph came before it.
 * - The box sits right after its paragraph. The reader lists the box's
 *   paragraphs where the block is, and a reopened package puts the block
 *   right after the paragraph holding its anchor, so a block written between
 *   the two is read in a different order once saved.
 *
 * Only a box placed `inlineWithPrevious` has an anchor: a `standalone` one
 * stands in for a paragraph that held nothing else.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { expectTextBoxAttrs } from "./attrs";
import { TEXT_BOX_ANCHOR_NODE_NAME } from "./extensions/nodes/TextBoxAnchorExtension";
import { readTextBoxAnchorAttrs } from "./textBoxAnchorAttrs";

/**
 * Add the anchor id of `node`, when it is an anchor that names one. An anchor
 * naming none ties no box to its paragraph.
 */
const collectAnchorId = (node: PMNode, ids: Set<string>): void => {
  const anchorId = textBoxAnchorIdOf(node);
  if (anchorId !== undefined) {
    ids.add(anchorId);
  }
};

/** The id a text box anchor names, or `undefined` for anything else. */
export const textBoxAnchorIdOf = (node: PMNode): string | undefined => {
  if (node.type.name !== TEXT_BOX_ANCHOR_NODE_NAME) {
    return undefined;
  }
  const attrs = readTextBoxAnchorAttrs(node);
  return attrs.ok ? attrs.value.anchorId : undefined;
};

/** The anchor ids of the text boxes drawn in `node`. */
export const textBoxAnchorIdsIn = (node: PMNode): Set<string> => {
  const ids = new Set<string>();
  node.descendants((child) => {
    collectAnchorId(child, ids);
    return true;
  });
  return ids;
};

/** The anchor id of a text box block drawn in a paragraph, or `undefined`. */
export const anchoredTextBoxId = (node: PMNode): string | undefined => {
  if (node.type.name !== "textBox") {
    return undefined;
  }
  const attrs = expectTextBoxAttrs(node);
  return attrs._docxPlacement === "inlineWithPrevious" ? attrs._docxAnchorId : undefined;
};

/**
 * Where the paragraph at `position` ends together with the text boxes drawn
 * in it that follow it: the place for a block inserted after the paragraph.
 */
export const endOfParagraphWithItsTextBoxes = (doc: PMNode, position: number): number => {
  const paragraph = doc.nodeAt(position);
  if (!paragraph) {
    return position;
  }
  let end = position + paragraph.nodeSize;
  const anchors = textBoxAnchorIdsIn(paragraph);
  if (anchors.size === 0) {
    return end;
  }
  const after = doc.resolve(end);
  for (let index = after.index(); index < after.parent.childCount; index += 1) {
    const sibling = after.parent.child(index);
    const anchorId = anchoredTextBoxId(sibling);
    if (anchorId === undefined || !anchors.has(anchorId)) {
      break;
    }
    end += sibling.nodeSize;
  }
  return end;
};

/** Anchor ids `before` holds and `after` does not. */
export const droppedTextBoxAnchorIds = (before: PMNode, after: PMNode): Set<string> => {
  const dropped = textBoxAnchorIdsIn(before);
  if (dropped.size === 0) {
    return dropped;
  }
  for (const anchorId of textBoxAnchorIdsIn(after)) {
    dropped.delete(anchorId);
  }
  return dropped;
};

/**
 * Delete the text box blocks whose anchors are in `anchorIds`. Returns whether
 * any went.
 */
export const deleteTextBoxesAnchoredAt = (
  tr: Transaction,
  anchorIds: ReadonlySet<string>,
): boolean => {
  if (anchorIds.size === 0) {
    return false;
  }
  const ranges: { from: number; to: number }[] = [];
  tr.doc.descendants((node, position) => {
    const anchorId = anchoredTextBoxId(node);
    if (anchorId !== undefined && anchorIds.has(anchorId)) {
      ranges.push({ from: position, to: position + node.nodeSize });
      return false;
    }
    return !node.isTextblock;
  });
  for (const { from, to } of ranges.toReversed()) {
    tr.delete(from, to);
  }
  return ranges.length > 0;
};

/**
 * Put every text box drawn in the paragraph at `position` right after it,
 * keeping the boxes' order. For a paragraph just split: the boxes follow the
 * second half, and those anchored in the first half move up to it.
 */
export const gatherTextBoxesAfterParagraph = (tr: Transaction, position: number): void => {
  const paragraph = tr.doc.nodeAt(position);
  if (!paragraph) {
    return;
  }
  const anchors = textBoxAnchorIdsIn(paragraph);
  if (anchors.size === 0) {
    return;
  }
  const boxesEnd = endOfParagraphWithItsTextBoxes(tr.doc, position);
  const at = tr.doc.resolve(boxesEnd);
  const moved: { node: PMNode; from: number }[] = [];
  let offset = boxesEnd;
  for (let index = at.index(); index < at.parent.childCount; index += 1) {
    const sibling = at.parent.child(index);
    const anchorId = anchoredTextBoxId(sibling);
    if (anchorId !== undefined && anchors.has(anchorId)) {
      moved.push({ node: sibling, from: offset });
    }
    offset += sibling.nodeSize;
  }
  if (moved.length === 0) {
    return;
  }
  // Every box moved sits after `boxesEnd`, so deleting them leaves it where
  // it is.
  for (const { node, from } of moved.toReversed()) {
    tr.delete(from, from + node.nodeSize);
  }
  tr.insert(boxesEnd, Fragment.fromArray(moved.map(({ node }) => node)));
};

/** The anchor ids of the text boxes drawn between `from` and `to`. */
export const textBoxAnchorIdsBetween = (doc: PMNode, from: number, to: number): Set<string> => {
  const ids = new Set<string>();
  doc.nodesBetween(from, to, (node) => {
    collectAnchorId(node, ids);
    return true;
  });
  return ids;
};

/**
 * Before the paragraph at `position` joins the one after it: move the text
 * boxes drawn in it, which stand between the two, past that paragraph, where
 * they follow the joined one. Returns whether any moved.
 */
export const moveTextBoxesPastNextParagraph = (tr: Transaction, position: number): boolean => {
  const paragraph = tr.doc.nodeAt(position);
  if (!paragraph) {
    return false;
  }
  const paragraphEnd = position + paragraph.nodeSize;
  const boxesEnd = endOfParagraphWithItsTextBoxes(tr.doc, position);
  const next = boxesEnd > paragraphEnd ? tr.doc.nodeAt(boxesEnd) : null;
  if (!next || next.type !== paragraph.type) {
    return false;
  }
  const boxes = tr.doc.slice(paragraphEnd, boxesEnd).content;
  tr.delete(paragraphEnd, boxesEnd);
  tr.insert(paragraphEnd + next.nodeSize, boxes);
  return true;
};
