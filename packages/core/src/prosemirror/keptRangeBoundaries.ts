/**
 * Range boundaries rejecting a change keeps from the content it removes.
 *
 * The engine keeps every bookmark and move range boundary of rejected content
 * where that content was and makes their ranges well-formed
 * (`rangeBoundaries.ts` in `@stll/docx-core/ops`); a comment whose whole range
 * was rejected goes with its comment, and accepting removes what a direct
 * deletion would. The editor resolves the same way, through one mapping from
 * every node that carries range boundaries to the engine's boundary items:
 * what it keeps and what it repairs are both read from that mapping, so a
 * kept boundary is always one the repair sees.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import {
  type BoundaryOccurrence,
  planWellFormedRanges,
} from "@stll/docx-core/ops";

import type { ParagraphContent } from "../types/document";
import { expectBookmarkBoundaryAttrs } from "./bookmarkBoundaryAttrs";
import { BLOCK_BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BlockBookmarkBoundaryExtension";
import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import {
  expectMoveRangeBoundaryAttrs,
  type MoveRangeMarker,
} from "./moveRangeBoundaryAttrs";
import { expectRangeAnchorAttrs } from "./rangeAnchorAttrs";

/** A range boundary, as the engine's model writes it. */
type BoundaryItem = Extract<
  ParagraphContent,
  {
    type:
      | "bookmarkStart"
      | "bookmarkEnd"
      | "commentRangeStart"
      | "commentRangeEnd"
      | "moveFromRangeStart"
      | "moveFromRangeEnd"
      | "moveToRangeStart"
      | "moveToRangeEnd";
  }
>;

/** Every editor node that carries range boundaries. Comment ranges carried as marks are the comment's. */
type BoundaryNodeName =
  | typeof BOOKMARK_BOUNDARY_NODE_NAME
  | typeof BLOCK_BOOKMARK_BOUNDARY_NODE_NAME
  | typeof MOVE_RANGE_BOUNDARY_NODE_NAME
  | typeof RANGE_ANCHOR_NODE_NAME;

const bookmarkItems = (node: PMNode): BoundaryItem[] => {
  const attrs = expectBookmarkBoundaryAttrs(node);
  return attrs.type === "start"
    ? [{ type: "bookmarkStart", id: attrs.id, name: attrs.name }]
    : [{ type: "bookmarkEnd", id: attrs.id }];
};

/** The boundaries each node carries, in document order. A missing kind fails to type-check. */
const BOUNDARY_ITEMS: Record<
  BoundaryNodeName,
  (node: PMNode) => BoundaryItem[]
> = {
  [BOOKMARK_BOUNDARY_NODE_NAME]: bookmarkItems,
  [BLOCK_BOOKMARK_BOUNDARY_NODE_NAME]: bookmarkItems,
  [MOVE_RANGE_BOUNDARY_NODE_NAME]: (node) => [
    expectMoveRangeBoundaryAttrs(node),
  ],
  // A point range holds its start and its end in one node.
  [RANGE_ANCHOR_NODE_NAME]: (node) => {
    const { start, end } = expectRangeAnchorAttrs(node);
    return [start, end];
  },
};

const isBoundaryNodeName = (name: string): name is BoundaryNodeName =>
  Object.hasOwn(BOUNDARY_ITEMS, name);

/** The range boundaries `node` carries; none for any other node. */
export const rangeBoundaryItemsOf = (node: PMNode): readonly BoundaryItem[] =>
  isBoundaryNodeName(node.type.name)
    ? BOUNDARY_ITEMS[node.type.name](node)
    : [];

const isComment = (item: BoundaryItem): boolean =>
  item.type === "commentRangeStart" || item.type === "commentRangeEnd";

/**
 * Whether resolving away the content around `node` keeps it in place: it
 * carries range boundaries and none is a comment's, which goes with the
 * comment.
 */
export const keepsRangeBoundary = (node: PMNode): boolean => {
  const items = rangeBoundaryItemsOf(node);
  return items.length > 0 && !items.some(isComment);
};

/** Each boundary's range family and side, as the engine pairs them. */
const BOUNDARY_SIDES: Record<
  BoundaryItem["type"],
  { family: string; side: "start" | "end" }
> = {
  bookmarkStart: { family: "bookmark", side: "start" },
  bookmarkEnd: { family: "bookmark", side: "end" },
  commentRangeStart: { family: "comment", side: "start" },
  commentRangeEnd: { family: "comment", side: "end" },
  moveFromRangeStart: { family: "moveFrom", side: "start" },
  moveFromRangeEnd: { family: "moveFrom", side: "end" },
  moveToRangeStart: { family: "moveTo", side: "start" },
  moveToRangeEnd: { family: "moveTo", side: "end" },
};

const occurrenceOf = (
  item: BoundaryItem,
  kept: boolean,
): BoundaryOccurrence => {
  const { family, side } = BOUNDARY_SIDES[item.type];
  return {
    boundary: `${item.type}:${item.id}`,
    range: `${family}:${item.id}`,
    side,
    kept,
  };
};

type Found = {
  position: number;
  node: PMNode;
  item: BoundaryItem;
  occurrence: BoundaryOccurrence;
};

const boundariesIn = (doc: PMNode, kept: ReadonlySet<number>): Found[] => {
  const found: Found[] = [];
  doc.descendants((node, position) => {
    for (const item of rangeBoundaryItemsOf(node))
      found.push({
        position,
        node,
        item,
        occurrence: occurrenceOf(item, kept.has(position)),
      });
    return true;
  });
  return found;
};

/** One boundary as a node of its own beside `beside`, at that node's level. */
const nodeFor = (item: BoundaryItem, beside: PMNode): PMNode | undefined => {
  const schema = beside.type.schema;
  if (item.type === "bookmarkStart" || item.type === "bookmarkEnd") {
    const attrs =
      item.type === "bookmarkStart"
        ? { type: "start", id: item.id, name: item.name }
        : { type: "end", id: item.id };
    return schema.nodes[
      beside.isInline
        ? BOOKMARK_BOUNDARY_NODE_NAME
        : BLOCK_BOOKMARK_BOUNDARY_NODE_NAME
    ]?.create(attrs);
  }
  if (isComment(item) || !beside.isInline) return undefined;
  return schema.nodes[MOVE_RANGE_BOUNDARY_NODE_NAME]?.create({
    marker: item satisfies MoveRangeMarker,
  });
};

/**
 * Make every range a rejection kept boundaries of well-formed, by the
 * engine's own rule (`planWellFormedRanges`): `kept` holds the positions, in
 * `tr.doc`, of the nodes the rejection kept from removed content.
 */
export const normalizeKeptRangeBoundaries = (
  tr: Transaction,
  kept: ReadonlySet<number>,
): void => {
  if (kept.size === 0) return;
  // Copies first: a node loses the boundaries the plan drops, and a point
  // range that keeps only one half becomes a node of that half.
  const found = boundariesIn(tr.doc, kept);
  const dropped = new Set(
    planWellFormedRanges(found.map(({ occurrence }) => occurrence)).dropped,
  );
  const start = tr.steps.length;
  const keptAfter = new Set<number>();
  const byNode = new Map<number, Found[]>();
  for (const entry of found)
    byNode.set(entry.position, [...(byNode.get(entry.position) ?? []), entry]);
  for (const [position, entries] of [...byNode].toReversed()) {
    const remaining = entries.filter(
      (entry) => !dropped.has(found.indexOf(entry)),
    );
    if (remaining.length === entries.length) continue;
    const node = entries[0]?.node;
    if (node === undefined) continue;
    const at = tr.mapping.slice(start).map(position);
    const [only] = remaining;
    const replacement =
      remaining.length === 1 && only !== undefined
        ? nodeFor(only.item, node)
        : undefined;
    if (replacement === undefined) tr.delete(at, at + node.nodeSize);
    else tr.replaceWith(at, at + node.nodeSize, replacement);
  }
  for (const position of kept)
    keptAfter.add(tr.mapping.slice(start).map(position));
  // Then ends before their starts, on what remains.
  const after = boundariesIn(tr.doc, keptAfter);
  const moves = planWellFormedRanges(
    after.map(({ occurrence }) => occurrence),
  ).moved;
  const movedFrom = tr.steps.length;
  for (const { end: endIndex, start: startIndex } of moves) {
    const end = after[endIndex];
    const startAt = after[startIndex];
    if (end === undefined || startAt === undefined) continue;
    // A point range holds both halves: it is never a reversed pair on its own.
    if (rangeBoundaryItemsOf(end.node).length > 1) continue;
    const moved = nodeFor(end.item, startAt.node);
    if (moved === undefined) continue;
    tr.insert(
      tr.mapping
        .slice(movedFrom)
        .map(startAt.position + startAt.node.nodeSize, -1),
      moved,
    );
    const at = tr.mapping.slice(movedFrom).map(end.position);
    tr.delete(at, at + end.node.nodeSize);
  }
};

/**
 * Where, in `doc`, a rejection keeps boundaries from the content it removes:
 * each node {@link keepsRangeBoundary} that carries a revision mark the
 * rejection removes, directly or as a nested layer.
 */
export const boundariesRejectionKeeps = (
  doc: PMNode,
  removes: (node: PMNode) => boolean,
): number[] => {
  const positions: number[] = [];
  doc.descendants((node, position) => {
    if (keepsRangeBoundary(node) && removes(node)) positions.push(position);
    return true;
  });
  return positions;
};
