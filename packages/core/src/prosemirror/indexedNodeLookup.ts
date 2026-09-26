/**
 * `nodesBetween`, `nodeAt` and the paragraph a position sits in, without
 * scanning a long fragment from its first child.
 *
 * ProseMirror finds the child holding a position by walking the parent's
 * children from index 0, so every lookup in a flat document is O(blocks).
 * Plugins that look up once per step, or once per paragraph they mapped, pay
 * O(blocks) that many times. These helpers binary-search a cached table of
 * child offsets instead and otherwise do exactly what the ProseMirror methods
 * do: the same nodes, visited in the same order, with the same arguments.
 */

import type { Fragment, Node as PMNode } from "prosemirror-model";

/** Below this many children a linear scan is as fast as building the table. */
const LINEAR_SCAN_LIMIT = 16;

const childOffsetCache = new WeakMap<Fragment, Int32Array>();

/** `offsets[i]` is where child `i` starts; `offsets[childCount]` is the size. */
const childOffsets = (fragment: Fragment): Int32Array => {
  const cached = childOffsetCache.get(fragment);
  if (cached) {
    return cached;
  }
  const offsets = new Int32Array(fragment.childCount + 1);
  let pos = 0;
  fragment.forEach((child, _offset, index) => {
    offsets[index] = pos;
    pos += child.nodeSize;
  });
  offsets[fragment.childCount] = pos;
  childOffsetCache.set(fragment, offsets);
  return offsets;
};

/** The first child that ends after `pos` (its end is past `pos`), and its start. */
const firstChildEndingAfter = (fragment: Fragment, pos: number): [index: number, start: number] => {
  if (fragment.childCount <= LINEAR_SCAN_LIMIT) {
    let start = 0;
    for (let index = 0; index < fragment.childCount; index++) {
      const end = start + fragment.child(index).nodeSize;
      if (end > pos) {
        return [index, start];
      }
      start = end;
    }
    return [fragment.childCount, start];
  }
  const offsets = childOffsets(fragment);
  let low = 0;
  let high = fragment.childCount;
  while (low < high) {
    const middle = (low + high) >>> 1;
    // SAFETY: middle + 1 <= childCount, inside the offsets table.
    if (offsets[middle + 1]! > pos) {
      high = middle;
    } else {
      low = middle + 1;
    }
  }
  // SAFETY: low <= childCount, inside the offsets table.
  return [low, offsets[low]!];
};

type NodesBetweenCallback = (
  node: PMNode,
  pos: number,
  parent: PMNode | null,
  index: number,
) => boolean | void;

const visitBetween = (
  parent: PMNode,
  from: number,
  to: number,
  visit: NodesBetweenCallback,
  nodeStart: number,
): void => {
  const fragment = parent.content;
  // Children ending at or before `from` are skipped by `Fragment.nodesBetween`
  // too; starting at the first one that does not is the whole difference.
  let [index, pos] = firstChildEndingAfter(fragment, from);
  for (; pos < to; index++) {
    const child = fragment.child(index);
    const end = pos + child.nodeSize;
    if (visit(child, nodeStart + pos, parent, index) !== false && child.content.size) {
      const start = pos + 1;
      visitBetween(
        child,
        Math.max(0, from - start),
        Math.min(child.content.size, to - start),
        visit,
        nodeStart + start,
      );
    }
    pos = end;
  }
};

/** `doc.nodesBetween(from, to, visit)`, locating `from` by binary search. */
export const nodesBetweenIndexed = (
  doc: PMNode,
  from: number,
  to: number,
  visit: NodesBetweenCallback,
): void => {
  visitBetween(doc, from, to, visit, 0);
};

/** `Fragment.findIndex(pos)`: the child at `pos`, or the one after a boundary. */
const findChildIndex = (fragment: Fragment, pos: number): [index: number, offset: number] => {
  if (pos === 0) {
    return [0, 0];
  }
  if (pos === fragment.size) {
    return [fragment.childCount, pos];
  }
  if (pos > fragment.size || pos < 0) {
    throw new RangeError(`Position ${String(pos)} outside of fragment (${fragment.toString()})`);
  }
  // The first child whose end reaches `pos`: a child ending exactly there
  // hands `pos` to its successor, as `findIndex` does.
  const [index, start] = firstChildEndingAfter(fragment, pos - 1);
  const end = start + fragment.child(index).nodeSize;
  return end === pos ? [index + 1, end] : [index, start];
};

/** `doc.nodeAt(pos)`. */
export const nodeAtIndexed = (doc: PMNode, pos: number): PMNode | null => {
  let rest = pos;
  for (let node = doc; ;) {
    const [index, offset] = findChildIndex(node.content, rest);
    const child = node.maybeChild(index);
    if (!child) {
      return null;
    }
    if (offset === rest || child.isText) {
      return child;
    }
    rest -= offset + 1;
    node = child;
  }
};

/**
 * The innermost paragraph `doc.resolve(pos)` passes through, and the position
 * before it (`$pos.before(depth)`), or null outside every paragraph.
 */
export const enclosingParagraphIndexed = (
  doc: PMNode,
  pos: number,
): { node: PMNode; pos: number } | null => {
  if (!(pos >= 0 && pos <= doc.content.size)) {
    throw new RangeError(`Position ${String(pos)} out of range`);
  }
  let paragraph: { node: PMNode; pos: number } | null = null;
  let start = 0;
  let parentOffset = pos;
  for (let node = doc; ;) {
    const [index, offset] = findChildIndex(node.content, parentOffset);
    const rest = parentOffset - offset;
    if (!rest) {
      return paragraph;
    }
    const child = node.child(index);
    if (child.isText) {
      return paragraph;
    }
    if (child.type.name === "paragraph") {
      paragraph = { node: child, pos: start + offset };
    }
    parentOffset = rest - 1;
    start += offset + 1;
    node = child;
  }
};
