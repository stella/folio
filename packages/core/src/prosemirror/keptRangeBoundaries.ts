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

const isComment = (
  item: BoundaryItem,
): item is Extract<
  BoundaryItem,
  { type: "commentRangeStart" | "commentRangeEnd" }
> => item.type === "commentRangeStart" || item.type === "commentRangeEnd";

/**
 * Whether resolving away the content around `node` keeps it in place: it
 * carries range boundaries and none is a comment's, which goes with the
 * comment.
 */
const keepsRangeBoundary = (node: PMNode): boolean => {
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

const boundariesIn = (doc: PMNode, kept: ReadonlySet<PMNode>): Found[] => {
  const found: Found[] = [];
  doc.descendants((node, position) => {
    for (const item of rangeBoundaryItemsOf(node))
      found.push({
        position,
        node,
        item,
        occurrence: occurrenceOf(item, kept.has(node)),
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

/** One change to `tr.doc` the plan asks for, at a position in that document. */
type BoundaryEdit =
  | { kind: "replace"; position: number; size: number; node?: PMNode }
  | { kind: "insert"; position: number; node: PMNode };

/**
 * Make every range a rejection kept boundaries of well-formed, by the
 * engine's own plan (`planWellFormedRanges`): `kept` holds the nodes
 * {@link removeResolvedSubtree} kept, found in `tr.doc` as themselves. The plan
 * runs once over every boundary in document order and names each by its
 * place in that list, as the engine names its own, so applying it never
 * reads a position the repair itself moved.
 */
export const normalizeKeptRangeBoundaries = (
  tr: Transaction,
  kept: ReadonlySet<PMNode>,
): void => {
  if (kept.size === 0) return;
  const found = boundariesIn(tr.doc, kept);
  const plan = planWellFormedRanges(found.map(({ occurrence }) => occurrence));
  const leaving = new Set(plan.dropped);
  for (const { end } of plan.moved) leaving.add(end);
  const edits: BoundaryEdit[] = [];
  const byNode = new Map<number, number[]>();
  for (const [index, { position }] of found.entries())
    byNode.set(position, [...(byNode.get(position) ?? []), index]);
  for (const [position, indices] of byNode) {
    const staying = indices.filter((index) => !leaving.has(index));
    const node = found[indices[0] ?? -1]?.node;
    if (node === undefined || staying.length === indices.length) continue;
    // A point range that keeps only one half becomes a node of that half.
    const only = staying.length === 1 ? found[staying[0] ?? -1] : undefined;
    const replacement =
      only === undefined ? undefined : nodeFor(only.item, node);
    edits.push({
      kind: "replace",
      position,
      size: node.nodeSize,
      ...(replacement === undefined ? {} : { node: replacement }),
    });
  }
  for (const { end: endIndex, start: startIndex } of plan.moved) {
    const end = found[endIndex];
    const start = found[startIndex];
    const moved =
      end === undefined || start === undefined
        ? undefined
        : nodeFor(end.item, start.node);
    if (start !== undefined && moved !== undefined)
      edits.push({
        kind: "insert",
        position: start.position + start.node.nodeSize,
        node: moved,
      });
  }
  // Last position first, so every edit lands where the plan read it; at one
  // position a node's own change goes before what is inserted ahead of it.
  const order = (edit: BoundaryEdit) => (edit.kind === "replace" ? 0 : 1);
  for (const edit of edits.toSorted(
    (left, right) =>
      right.position - left.position || order(left) - order(right),
  )) {
    if (edit.kind === "insert") tr.insert(edit.position, edit.node);
    else if (edit.node === undefined)
      tr.delete(edit.position, edit.position + edit.size);
    else tr.replaceWith(edit.position, edit.position + edit.size, edit.node);
  }
};

declare const removedSubtreeBrand: unique symbol;

/**
 * What stays of a subtree a resolution removes. Only
 * {@link removeResolvedSubtree} makes one, so no resolution can drop content
 * without being handed the boundaries it holds.
 */
export type RemovedSubtree = {
  /** Fresh nodes, in document order, to stand where the subtree was. */
  readonly kept: readonly PMNode[];
  readonly [removedSubtreeBrand]: true;
};

const REVISION_MARK_NAMES: ReadonlySet<string> = new Set([
  "insertion",
  "deletion",
]);

/**
 * Remove `node` and everything in it. Rejecting keeps every range boundary
 * within it, at any depth, as the engine's `removedContentBoundaries` does,
 * carrying no revision of its own; accepting keeps none, as a direct
 * deletion does. Each kept node is a new node, so the range repair can tell
 * it from a boundary that was already standing.
 */
export const removeResolvedSubtree = (
  node: PMNode,
  mode: "accept" | "reject",
): RemovedSubtree => {
  const kept: PMNode[] = [];
  const visit = (candidate: PMNode): void => {
    if (keepsRangeBoundary(candidate)) {
      kept.push(
        candidate.type.create(
          candidate.attrs,
          null,
          candidate.marks.filter(
            (mark) => !REVISION_MARK_NAMES.has(mark.type.name),
          ),
        ),
      );
      return;
    }
    candidate.forEach(visit);
  };
  if (mode === "reject") visit(node);
  // SAFETY: this is the brand's only constructor.
  return { kept } as unknown as RemovedSubtree;
};
