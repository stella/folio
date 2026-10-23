/**
 * Range boundaries rejecting a change keeps from the content it removes.
 *
 * The engine keeps every bookmark and move range boundary of rejected content
 * where that content was (`rangeBoundaries.ts` in `@stll/docx-core/ops`); a
 * comment whose whole range was rejected goes with its comment, and accepting
 * removes what a direct deletion would. The editor resolves the same way:
 * when rejecting, such a boundary loses its revision marks and stays.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import {
  type BoundaryOccurrence,
  planWellFormedRanges,
} from "@stll/docx-core/ops";

import { expectBookmarkBoundaryAttrs } from "./bookmarkBoundaryAttrs";
import { BLOCK_BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BlockBookmarkBoundaryExtension";
import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { expectMoveRangeBoundaryAttrs } from "./moveRangeBoundaryAttrs";
import { expectRangeAnchorAttrs } from "./rangeAnchorAttrs";

/** Whether resolving away the content around `node` keeps it in place. */
export const keepsRangeBoundary = (node: PMNode): boolean => {
  switch (node.type.name) {
    case BOOKMARK_BOUNDARY_NODE_NAME:
    case MOVE_RANGE_BOUNDARY_NODE_NAME:
      return true;
    // A point range: a move keeps it, a comment's whole range goes with the comment.
    case RANGE_ANCHOR_NODE_NAME:
      return expectRangeAnchorAttrs(node).start.type !== "commentRangeStart";
    default:
      return false;
  }
};

type EditorBoundary = {
  position: number;
  node: PMNode;
  occurrence: BoundaryOccurrence;
};

/** The boundary a node is, in the planner's terms; `undefined` for any other node. */
const boundaryOf = (
  node: PMNode,
): Omit<BoundaryOccurrence, "kept"> | undefined => {
  switch (node.type.name) {
    case BOOKMARK_BOUNDARY_NODE_NAME:
    case BLOCK_BOOKMARK_BOUNDARY_NODE_NAME: {
      const { type, id } = expectBookmarkBoundaryAttrs(node);
      return {
        boundary: `${type === "start" ? "bookmarkStart" : "bookmarkEnd"}:${id}`,
        range: `bookmark:${id}`,
        side: type,
      };
    }
    case MOVE_RANGE_BOUNDARY_NODE_NAME: {
      const marker = expectMoveRangeBoundaryAttrs(node);
      return {
        boundary: `${marker.type}:${marker.id}`,
        range: `${marker.type.startsWith("moveFrom") ? "moveFrom" : "moveTo"}:${marker.id}`,
        side: marker.type.endsWith("Start") ? "start" : "end",
      };
    }
    default:
      return undefined;
  }
};

/** Each boundary node's twin at the other level. */
const BOUNDARY_TWINS: ReadonlyMap<string, string> = new Map([
  [BOOKMARK_BOUNDARY_NODE_NAME, BLOCK_BOOKMARK_BOUNDARY_NODE_NAME],
  [BLOCK_BOOKMARK_BOUNDARY_NODE_NAME, BOOKMARK_BOUNDARY_NODE_NAME],
]);

/** The same boundary at the level of `beside`, for an end that moves next to its start. */
const atLevelOf = (end: PMNode, beside: PMNode): PMNode | undefined => {
  if (end.isInline === beside.isInline) return end;
  const twin = BOUNDARY_TWINS.get(end.type.name);
  return twin === undefined
    ? undefined
    : end.type.schema.nodes[twin]?.create(end.attrs);
};

/**
 * Make every range a rejection kept boundaries of well-formed, by the
 * engine's own rule (`planWellFormedRanges`): `kept` holds the positions, in
 * `tr.doc`, of the boundaries the rejection kept from removed content.
 */
export const normalizeKeptRangeBoundaries = (
  tr: Transaction,
  kept: ReadonlySet<number>,
): void => {
  if (kept.size === 0) return;
  const found: EditorBoundary[] = [];
  tr.doc.descendants((node, position) => {
    const boundary = boundaryOf(node);
    if (boundary !== undefined)
      found.push({
        position,
        node,
        occurrence: { ...boundary, kept: kept.has(position) },
      });
    return true;
  });
  const plan = planWellFormedRanges(found.map(({ occurrence }) => occurrence));
  const from = tr.steps.length;
  const mapped = (position: number, assoc = 1) =>
    tr.mapping.slice(from).map(position, assoc);
  for (const index of plan.dropped) {
    const dropped = found[index];
    if (dropped === undefined) continue;
    const at = mapped(dropped.position);
    tr.delete(at, at + dropped.node.nodeSize);
  }
  for (const { end: endIndex, start: startIndex } of plan.moved) {
    const end = found[endIndex];
    const start = found[startIndex];
    const moved =
      end === undefined || start === undefined
        ? undefined
        : atLevelOf(end.node, start.node);
    if (end === undefined || start === undefined || moved === undefined)
      continue;
    tr.insert(mapped(start.position + start.node.nodeSize, -1), moved);
    const at = mapped(end.position);
    tr.delete(at, at + end.node.nodeSize);
  }
};

/**
 * Where, in `doc`, a rejection keeps boundaries from the content it removes:
 * each boundary {@link keepsRangeBoundary} that carries a revision mark the
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
