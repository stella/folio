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

import type { Mark, Node as PMNode, Schema } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import {
  type BoundaryOccurrence,
  paragraphRemains,
  planWellFormedRanges,
} from "@stll/docx-core/ops";

import type { CommentReference, ParagraphContent } from "../types/document";
import { expectBookmarkBoundaryAttrs } from "./bookmarkBoundaryAttrs";
import { expectCommentRangeBoundaryAttrs } from "./commentRangeBoundaryAttrs";
import { expectCommentReferenceAttrs } from "./commentReferenceAttrs";
import {
  copyOriginBoundaryAfterRemovingAddedAncestors,
  copyOriginBoundaryAfterResolution,
} from "./copyOriginResolution";
import { planEmptyRanges } from "./emptyRangeAnchor";
import { BLOCK_BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BlockBookmarkBoundaryExtension";
import { BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/BlockMoveRangeBoundaryExtension";
import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { COMMENT_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/CommentRangeBoundaryExtension";
import { COMMENT_REFERENCE_NODE_NAME } from "./extensions/nodes/CommentReferenceExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import {
  expectMoveRangeBoundaryAttrs,
  type MoveRangeMarker,
} from "./moveRangeBoundaryAttrs";
import { expectRangeAnchorAttrs } from "./rangeAnchorAttrs";
import type { TrackedRevisionAncestor } from "./schema/marks";
import {
  retainedTrackedRevisionMark,
  trackedRevisionPathOf,
} from "./trackedRevisionPath";

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

/** The marks that say a node is in a revision: its wrappers, in the engine. */
const REVISION_MARK_NAMES: ReadonlySet<string> = new Set([
  "insertion",
  "deletion",
]);

/** Every editor node that carries range boundaries; comment marks carry span coverage. */
type BoundaryNodeName =
  | typeof BOOKMARK_BOUNDARY_NODE_NAME
  | typeof BLOCK_BOOKMARK_BOUNDARY_NODE_NAME
  | typeof MOVE_RANGE_BOUNDARY_NODE_NAME
  | typeof BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME
  | typeof RANGE_ANCHOR_NODE_NAME
  | typeof COMMENT_RANGE_BOUNDARY_NODE_NAME;

const bookmarkItems = (node: PMNode): BoundaryItem[] => {
  const attrs = expectBookmarkBoundaryAttrs(node);
  return attrs.type === "start"
    ? [
        {
          type: "bookmarkStart",
          id: attrs.id,
          name: attrs.name,
          ...(attrs.copyOriginAncestors === undefined
            ? {}
            : { copyOriginAncestors: attrs.copyOriginAncestors }),
        },
      ]
    : [
        {
          type: "bookmarkEnd",
          id: attrs.id,
          ...(attrs.copyOriginAncestors === undefined
            ? {}
            : { copyOriginAncestors: attrs.copyOriginAncestors }),
        },
      ];
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
  [BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME]: (node) => [
    expectMoveRangeBoundaryAttrs(node),
  ],
  // A point range holds its start and its end in one node.
  [RANGE_ANCHOR_NODE_NAME]: (node) => {
    const { start, end } = expectRangeAnchorAttrs(node);
    return [start, end];
  },
  [COMMENT_RANGE_BOUNDARY_NODE_NAME]: (node) => [
    expectCommentRangeBoundaryAttrs(node).marker,
  ],
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

const isBookmark = (
  item: BoundaryItem,
): item is Extract<BoundaryItem, { type: "bookmarkStart" | "bookmarkEnd" }> =>
  item.type === "bookmarkStart" || item.type === "bookmarkEnd";

/**
 * The one half a move's point range keeps, as a node of its own on `marks`:
 * the only boundary the repair makes, since one node holds both halves of a
 * point range. A comment's point range goes with its comment, and a bookmark
 * is never one.
 */
const pointRangeHalf = (
  node: PMNode,
  half: BoundaryItem,
  marks: readonly Mark[],
): PMNode | undefined => {
  if (isComment(half) || isBookmark(half)) return undefined;
  return node.type.schema.nodes[MOVE_RANGE_BOUNDARY_NODE_NAME]?.create(
    { marker: half satisfies MoveRangeMarker },
    null,
    marks,
  );
};

/** One boundary as a node of its own, inline or between blocks. */
const boundaryNode = (
  item: BoundaryItem,
  schema: Schema,
  inline: boolean,
): PMNode | undefined => {
  if (item.type === "bookmarkStart" || item.type === "bookmarkEnd") {
    const attrs =
      item.type === "bookmarkStart"
        ? {
            type: "start",
            id: item.id,
            name: item.name,
            ...(item.copyOriginAncestors === undefined
              ? {}
              : { copyOriginAncestors: item.copyOriginAncestors }),
          }
        : {
            type: "end",
            id: item.id,
            ...(item.copyOriginAncestors === undefined
              ? {}
              : { copyOriginAncestors: item.copyOriginAncestors }),
          };
    return schema.nodes[
      inline ? BOOKMARK_BOUNDARY_NODE_NAME : BLOCK_BOOKMARK_BOUNDARY_NODE_NAME
    ]?.create(attrs);
  }
  if (isComment(item)) return undefined;
  return schema.nodes[
    inline ? MOVE_RANGE_BOUNDARY_NODE_NAME : BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME
  ]?.create({ marker: item satisfies MoveRangeMarker });
};

/** Range boundary nodes, as the nodes that stand for them between blocks. */
export const blockBoundariesOf = (nodes: readonly PMNode[]): PMNode[] =>
  nodes.flatMap((node) =>
    rangeBoundaryItemsOf(node).flatMap((item) => {
      const block = boundaryNode(item, node.type.schema, false);
      return block === undefined ? [] : [block];
    }),
  );

/** Whether `node` holds a comment's anchor, which only a paragraph holds. */
export const carriesCommentAnchor = (node: PMNode): boolean =>
  node.type.name === COMMENT_REFERENCE_NODE_NAME ||
  rangeBoundaryItemsOf(node).some(isComment);

/** Whether `node` is a range boundary standing between blocks. */
export const isBlockBoundaryNode = (node: PMNode | null | undefined): boolean =>
  node != null && !node.isInline && rangeBoundaryItemsOf(node).length > 0;

/**
 * The inline form of range boundaries between blocks whose paragraph break
 * goes, so they stand between the joined paragraphs' content; `null` when
 * one of `blocks` is not such a boundary.
 */
export const inlineBoundariesOf = (
  blocks: readonly PMNode[],
): PMNode[] | null => {
  const inline: PMNode[] = [];
  for (const block of blocks) {
    if (!isBlockBoundaryNode(block)) return null;
    for (const item of rangeBoundaryItemsOf(block)) {
      const node = boundaryNode(item, block.type.schema, true);
      if (node === undefined) return null;
      inline.push(node);
    }
  }
  return inline;
};

/**
 * What `paragraph` leaves where it was when it goes whole, by the engine's
 * own rule (`paragraphRemains`): before a table, its bookmark and move range
 * boundaries as block nodes in its place (`blocks`), unless a comment needs
 * the paragraph before it, which takes every carrier in source order
 * (`comments`); ending its
 * container, every node carrying a boundary (`inline`), for the paragraph
 * before it.
 */
export const paragraphRemainsOf = (
  paragraph: PMNode,
): { blocks: PMNode[]; comments: PMNode[]; inline: PMNode[] } => {
  const items: (BoundaryItem | CommentReference)[] = [];
  const inline: PMNode[] = [];
  paragraph.descendants((node) => {
    const carried = rangeBoundaryItemsOf(node);
    const remains =
      carried.length > 0
        ? copyOriginBoundaryAfterRemovingAddedAncestors(node)
        : node;
    if (
      (carried.length > 0 || node.type.name === COMMENT_REFERENCE_NODE_NAME) &&
      remains !== null
    )
      inline.push(remains);
    if (remains !== null) items.push(...rangeBoundaryItemsOf(remains));
    if (node.type.name === COMMENT_REFERENCE_NODE_NAME)
      items.push({
        type: "commentReference",
        id: expectCommentReferenceAttrs(node).commentId,
      });
    return true;
  });
  const remains = paragraphRemains(items);
  if (remains.placement === "previous")
    return { blocks: [], comments: inline, inline };
  const blocks = remains.blocks.flatMap((item) => {
    const block = boundaryNode(item, paragraph.type.schema, false);
    return block === undefined ? [] : [block];
  });
  return { blocks, comments: [], inline };
};

/** The revision marks of `anchor`, with the other marks of `node`. */
const marksBeside = (node: PMNode, anchor: PMNode): Mark[] => [
  ...node.marks.filter((mark) => !REVISION_MARK_NAMES.has(mark.type.name)),
  ...anchor.marks.filter((mark) => REVISION_MARK_NAMES.has(mark.type.name)),
];

/**
 * The kept boundary `moving` names, moved beside the other end of its range,
 * `anchor`: the same node, its attributes and its own marks, inside the
 * revisions `anchor` stands in, as the engine moves a boundary in among its
 * anchor's wrappers.
 */
const movedBeside = (moving: Found, anchor: PMNode): PMNode | undefined =>
  rangeBoundaryItemsOf(moving.node).length === 1
    ? moving.node.mark(marksBeside(moving.node, anchor))
    : pointRangeHalf(
        moving.node,
        moving.item,
        marksBeside(moving.node, anchor),
      );

/**
 * One change the plan asks of `tr.doc`, at a position in that document. The
 * repair deletes a copy or moves the original node, attributes and marks and
 * all; it builds a node only for the half a point range keeps
 * ({@link pointRangeHalf}), on the point range's own marks.
 */
type BoundaryEdit =
  | { kind: "delete"; position: number; node: PMNode }
  | { kind: "narrow"; position: number; node: PMNode; half: BoundaryItem }
  | { kind: "move"; position: number; node: PMNode };

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
  if (kept.size > 0) wellFormKeptRangeBoundaries(tr, kept);
  collapseEmptyMoveRanges(tr);
};

/**
 * A move range whose start and end a resolution left meeting with nothing
 * between, as one point anchor: the node a reopened package reads that empty
 * range as (`planEmptyRanges`), so the resolved document has the same shape
 * live and after a save.
 */
const collapseEmptyMoveRanges = (tr: Transaction): void => {
  const paragraphs: number[] = [];
  tr.doc.descendants((node, position) => {
    if (!node.isTextblock) return true;
    let boundaries = 0;
    node.forEach((child) => {
      if (child.type.name === MOVE_RANGE_BOUNDARY_NODE_NAME) boundaries += 1;
    });
    if (boundaries > 1) paragraphs.push(position);
    return false;
  });
  const anchorType = tr.doc.type.schema.nodes[RANGE_ANCHOR_NODE_NAME];
  if (anchorType === undefined) return;
  for (const position of paragraphs.toReversed()) {
    const paragraph = tr.doc.nodeAt(position);
    if (paragraph === null) continue;
    const children: { node: PMNode; offset: number }[] = [];
    paragraph.forEach((node, offset) => children.push({ node, offset }));
    // Runs of move range boundaries alone, last first so positions hold.
    const runs: { from: number; to: number }[] = [];
    for (let index = 0; index < children.length; index += 1) {
      if (children[index]?.node.type.name !== MOVE_RANGE_BOUNDARY_NODE_NAME)
        continue;
      let end = index;
      while (
        children[end + 1]?.node.type.name === MOVE_RANGE_BOUNDARY_NODE_NAME
      )
        end += 1;
      if (end > index) runs.push({ from: index, to: end + 1 });
      index = end;
    }
    for (const { from, to } of runs.toReversed()) {
      const run = children.slice(from, to);
      const plan = planEmptyRanges(
        run.map(({ node }) => expectMoveRangeBoundaryAttrs(node)),
      );
      if (plan.anchorAt.size === 0) continue;
      const replacement: PMNode[] = [];
      for (const [index, { node }] of run.entries()) {
        const anchored = plan.anchorAt.get(index);
        if (anchored !== undefined)
          replacement.push(anchorType.create(anchored, null, node.marks));
        else if (!plan.closedAt.has(index)) replacement.push(node);
      }
      const first = run[0];
      const last = run.at(-1);
      if (first === undefined || last === undefined) continue;
      const start = position + 1 + first.offset;
      tr.replaceWith(
        start,
        position + 1 + last.offset + last.node.nodeSize,
        replacement,
      );
    }
  }
};

const wellFormKeptRangeBoundaries = (
  tr: Transaction,
  kept: ReadonlySet<PMNode>,
): void => {
  const found = boundariesIn(tr.doc, kept);
  const plan = planWellFormedRanges(found.map(({ occurrence }) => occurrence));
  const leaving = new Set(plan.dropped);
  for (const { moving } of plan.moved) leaving.add(moving);
  const edits: BoundaryEdit[] = [];
  const byNode = new Map<number, number[]>();
  for (const [index, { position }] of found.entries())
    byNode.set(position, [...(byNode.get(position) ?? []), index]);
  for (const [position, indices] of byNode) {
    const staying = indices.filter((index) => !leaving.has(index));
    const node = found[indices[0] ?? -1]?.node;
    if (node === undefined || staying.length === indices.length) continue;
    const half = staying.length === 1 ? found[staying[0] ?? -1] : undefined;
    edits.push(
      half === undefined
        ? { kind: "delete", position, node }
        : { kind: "narrow", position, node, half: half.item },
    );
  }
  for (const { moving: movingIndex, anchor: anchorIndex, side } of plan.moved) {
    const moving = found[movingIndex];
    const anchor = found[anchorIndex];
    const moved =
      moving === undefined || anchor === undefined
        ? undefined
        : movedBeside(moving, anchor.node);
    if (anchor !== undefined && moved !== undefined)
      edits.push({
        kind: "move",
        position:
          side === "after"
            ? anchor.position + anchor.node.nodeSize
            : anchor.position,
        node: moved,
      });
  }
  // Last position first, so every edit lands where the plan read it; at one
  // position a node's own change goes before what moves in ahead of it.
  const order = (edit: BoundaryEdit) => (edit.kind === "move" ? 1 : 0);
  for (const edit of edits.toSorted(
    (left, right) =>
      right.position - left.position || order(left) - order(right),
  )) {
    if (edit.kind === "move") tr.insert(edit.position, edit.node);
    else if (edit.kind === "delete")
      tr.delete(edit.position, edit.position + edit.node.nodeSize);
    else {
      const half = pointRangeHalf(edit.node, edit.half, edit.node.marks);
      if (half === undefined)
        tr.delete(edit.position, edit.position + edit.node.nodeSize);
      else
        tr.replaceWith(edit.position, edit.position + edit.node.nodeSize, half);
    }
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

/** What a resolution resolves: the mode, and which revisions it selects. */
export type RevisionSelection = {
  mode: "accept" | "reject";
  resolves: (revisionId: number) => boolean;
};

/** Whether resolving `layer` by `mode` takes its content out. */
export const revisionLayerRemovesContent = (
  layer: TrackedRevisionAncestor,
  mode: "accept" | "reject",
): boolean =>
  mode === "accept"
    ? layer.type === "deletion" || layer.type === "moveFrom"
    : layer.type === "insertion" || layer.type === "moveTo";

const layerKey = (layer: TrackedRevisionAncestor): string =>
  `${layer.type}:${layer.revisionId}`;

/**
 * The revision marks a boundary carries where `removed` stood: of each of
 * its revision marks, the layers outside the outermost one this resolution
 * takes content out by, less those it resolves. A revision nested inside
 * removed content goes with that content, as the engine's wrappers do.
 */
const landingMarks = (removed: PMNode, selection: RevisionSelection): Mark[] =>
  removed.marks.flatMap((mark) => {
    if (!REVISION_MARK_NAMES.has(mark.type.name)) return [];
    const path = trackedRevisionPathOf(mark);
    const cut = path.findIndex(
      (layer) =>
        selection.resolves(layer.revisionId) &&
        revisionLayerRemovesContent(layer, selection.mode),
    );
    const standing = new Set(
      (cut === -1 ? path : path.slice(0, cut))
        .filter((layer) => !selection.resolves(layer.revisionId))
        .map(layerKey),
    );
    if (standing.size === path.length) return [mark];
    const landing = retainedTrackedRevisionMark(mark, (layer) =>
      standing.has(layerKey(layer)),
    );
    return landing === null ? [] : [landing];
  });

/**
 * Marks that say which content a range covers. A kept boundary is no content,
 * so it carries none: a comment whose content went goes with it, as in the
 * engine, whose comment anchors are boundaries of their own.
 */
const COVERAGE_MARK_NAMES: ReadonlySet<string> = new Set(["comment"]);

/**
 * Remove `node` and everything in it. Rejecting keeps every range boundary
 * within it, at any depth, as the engine's `removedContentBoundaries` does;
 * accepting keeps none, as a direct deletion does. A kept boundary keeps its
 * own other marks and carries the revisions of where `node` stood
 * ({@link landingMarks}), never one from inside the removed content. Each
 * kept node is a new node, so the range repair can tell it from a boundary
 * that was already standing.
 */
export const removeResolvedSubtree = (
  node: PMNode,
  selection: RevisionSelection,
): RemovedSubtree => {
  const kept: PMNode[] = [];
  const landing = landingMarks(node, selection);
  const visit = (candidate: PMNode): void => {
    if (keepsRangeBoundary(candidate)) {
      const withResolvedCopyOrigin = copyOriginBoundaryAfterResolution(
        candidate,
        selection,
      );
      if (withResolvedCopyOrigin === null) return;
      kept.push(
        withResolvedCopyOrigin.type.create(withResolvedCopyOrigin.attrs, null, [
          ...candidate.marks.filter(
            (mark) =>
              !REVISION_MARK_NAMES.has(mark.type.name) &&
              !COVERAGE_MARK_NAMES.has(mark.type.name),
          ),
          ...landing,
        ]),
      );
      return;
    }
    candidate.forEach(visit);
  };
  if (selection.mode === "reject") visit(node);
  // SAFETY: this is the brand's only constructor.
  return { kept } as unknown as RemovedSubtree;
};
