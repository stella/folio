/**
 * A paragraph's content as a sequence of leaves, and the three moves every
 * edit is made of: partition the leaves between gaps, find the records that
 * run across a gap, and merge records that meet at one.
 *
 * A leaf is a unit of the offset space (a character of run text, a tab, a
 * field) or a zero-width child (a range marker, markup that shows nothing, a
 * run or container with nothing in it). A gap is a place between two leaves:
 * an offset plus how many zero-width leaves at that offset come before it.
 *
 * Partitioning keeps each record whose leaves all fall on one side as the
 * same object and cuts the others in two, each half keeping every field of
 * the original. Merging is the reverse: two records alike in everything but
 * their content become one. Because a deletion is "keep the leaves outside
 * the gaps" and an insertion is "cut at a gap and merge the new leaves in",
 * each is the other's exact inverse, whatever the records around them hold.
 */

import { panic } from "better-result";

import type { ParagraphContent, RunContent } from "../model/document";
import { structurallyEqual } from "./equality";
import {
  childrenOf,
  inlineWidth,
  isIdentifiedContainer,
  isInlineContainer,
  isOpeningMarker,
  withChildren,
} from "./offsets";

/** A place between two leaves. */
export type Gap = { offset: number; zeroWidthBefore: number };

/** A node of a paragraph's content tree: a paragraph child or a run child. */
export type InlineNode = ParagraphContent | RunContent;

const RUN_CONTENT_TYPES = {
  text: true,
  tab: true,
  break: true,
  symbol: true,
  footnoteRef: true,
  endnoteRef: true,
  fieldChar: true,
  instrText: true,
  softHyphen: true,
  noBreakHyphen: true,
  renderedPageBreak: true,
  preservedXml: true,
  drawing: true,
  shape: true,
} as const satisfies Record<RunContent["type"], true>;

const isRunContent = (node: InlineNode): node is RunContent =>
  Object.hasOwn(RUN_CONTENT_TYPES, node.type);

export const isParagraphContent = (node: InlineNode): node is ParagraphContent =>
  !isRunContent(node);

/** The children of a run or container; `undefined` for anything that has none. */
const childNodes = (node: InlineNode): readonly InlineNode[] | undefined => {
  if (node.type === "run") {
    return node.content;
  }
  if (isRunContent(node)) {
    return undefined;
  }
  return isInlineContainer(node) ? childrenOf(node) : undefined;
};

/** A record whose two halves would both carry one revision or control id. */
export const isIdentifiedNode = (node: InlineNode): boolean =>
  isParagraphContent(node) && isInlineContainer(node) && isIdentifiedContainer(node);

type NodeKind = "characters" | "branch" | "unit" | "zeroWidth";

const kindOf = (node: InlineNode): NodeKind => {
  if (node.type === "text") {
    return node.text === "" ? "zeroWidth" : "characters";
  }
  const children = childNodes(node);
  if (children !== undefined) {
    return children.length === 0 ? "zeroWidth" : "branch";
  }
  if (isRunContent(node)) {
    return "unit";
  }
  return inlineWidth(node) === 1 ? "unit" : "zeroWidth";
};

type Cursor = { position: number; zeroWidthSeen: number };

const startCursor = (): Cursor => ({ position: 0, zeroWidthSeen: 0 });

const unitRegion = (gaps: readonly Gap[], unit: number): number => {
  let region = 0;
  for (const gap of gaps) {
    if (gap.offset <= unit) region += 1;
  }
  return region;
};

const zeroWidthRegion = (gaps: readonly Gap[], cursor: Cursor): number => {
  let region = 0;
  for (const gap of gaps) {
    if (
      gap.offset < cursor.position ||
      (gap.offset === cursor.position && gap.zeroWidthBefore <= cursor.zeroWidthSeen)
    ) {
      region += 1;
    }
  }
  return region;
};

/** Narrow rebuilt children back to what their parent admits; anything else is a bug here. */
const narrowNodes = <Node extends InlineNode>(
  nodes: readonly InlineNode[],
  admits: (node: InlineNode) => node is Node,
): Node[] => {
  const out: Node[] = [];
  for (const node of nodes) {
    if (!admits(node)) {
      return panic(`A ${node.type} is not admitted here.`);
    }
    out.push(node);
  }
  return out;
};

/** Paragraph children, narrowed from nodes. */
export const asParagraphContent = (nodes: readonly InlineNode[]): ParagraphContent[] =>
  narrowNodes(nodes, isParagraphContent);

/** The same run or container holding other children. */
const rebuild = (node: InlineNode, children: readonly InlineNode[]): InlineNode => {
  if (node.type === "run") {
    return { ...node, content: narrowNodes(children, isRunContent) };
  }
  if (isRunContent(node) || !isInlineContainer(node)) {
    return panic(`A ${node.type} has no children to rebuild.`);
  }
  return withChildren(node, narrowNodes(children, isParagraphContent));
};

type NodePieces = { pieces: [number, InlineNode][]; min: number; max: number };

const partitionNode = (node: InlineNode, gaps: readonly Gap[], cursor: Cursor): NodePieces => {
  if (node.type === "text" && node.text !== "") {
    const start = cursor.position;
    const pieces: [number, InlineNode][] = [];
    let pieceStart = 0;
    let region = unitRegion(gaps, start);
    for (let index = 1; index < node.text.length; index += 1) {
      const next = unitRegion(gaps, start + index);
      if (next !== region) {
        pieces.push([region, { ...node, text: node.text.slice(pieceStart, index) }]);
        pieceStart = index;
        region = next;
      }
    }
    cursor.position += node.text.length;
    cursor.zeroWidthSeen = 0;
    if (pieces.length === 0) {
      return { pieces: [[region, node]], min: region, max: region };
    }
    pieces.push([region, { ...node, text: node.text.slice(pieceStart) }]);
    return { pieces, min: pieces[0]?.[0] ?? region, max: region };
  }
  const kind = kindOf(node);
  switch (kind) {
    case "unit": {
      const region = unitRegion(gaps, cursor.position);
      cursor.position += 1;
      cursor.zeroWidthSeen = 0;
      return { pieces: [[region, node]], min: region, max: region };
    }
    case "zeroWidth": {
      const region = zeroWidthRegion(gaps, cursor);
      cursor.zeroWidthSeen += 1;
      return { pieces: [[region, node]], min: region, max: region };
    }
    case "characters":
      return panic("Run text is partitioned by character above.");
    case "branch": {
      const inner = partitionNodes(childNodes(node) ?? [], gaps, cursor);
      if (inner.min === inner.max) {
        return { pieces: [[inner.min, node]], min: inner.min, max: inner.max };
      }
      const pieces: [number, InlineNode][] = [];
      for (const [region, part] of inner.parts.entries()) {
        if (part.length > 0) pieces.push([region, rebuild(node, part)]);
      }
      return { pieces, min: inner.min, max: inner.max };
    }
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};

type Parts = { parts: InlineNode[][]; min: number; max: number };

const partitionNodes = (
  items: readonly InlineNode[],
  gaps: readonly Gap[],
  cursor: Cursor,
): Parts => {
  const parts = Array.from({ length: gaps.length + 1 }, (): InlineNode[] => []);
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (const item of items) {
    const piece = partitionNode(item, gaps, cursor);
    for (const [region, node] of piece.pieces) {
      parts[region]?.push(node);
    }
    min = Math.min(min, piece.min);
    max = Math.max(max, piece.max);
  }
  return { parts, min, max };
};

/**
 * The content between consecutive gaps, one list per region: the records of
 * each region, cut where a gap passes through them. Gaps must be in order.
 */
export const partitionContent = (
  items: readonly InlineNode[],
  gaps: readonly Gap[],
): InlineNode[][] => partitionNodes(items, gaps, startCursor()).parts;

/**
 * The records holding leaves of both region `first` and region `last`,
 * outermost first. They form one chain: each is inside the one before.
 */
export const spanningRecords = (
  items: readonly InlineNode[],
  gaps: readonly Gap[],
  first: number,
  last: number,
): InlineNode[] => {
  const spine: InlineNode[] = [];
  const cursor = startCursor();
  let list: readonly InlineNode[] | undefined = items;
  while (list !== undefined) {
    let found: InlineNode | undefined;
    for (const item of list) {
      const start = { ...cursor };
      const { min, max } = partitionNode(item, gaps, cursor);
      if (min <= first && max >= last) {
        found = item;
        cursor.position = start.position;
        cursor.zeroWidthSeen = start.zeroWidthSeen;
        break;
      }
    }
    if (found === undefined) {
      return spine;
    }
    spine.push(found);
    list = found.type === "text" ? undefined : childNodes(found);
  }
  return spine;
};

const childrenKey = (node: InlineNode): string =>
  node.type === "hyperlink" ? "children" : "content";

const ownFields = (node: InlineNode): [string, unknown][] => {
  const key = childrenKey(node);
  return Object.entries(node).filter(([name]) => name !== key);
};

/** Whether two records are alike in everything but their content. */
const sameOwnFields = (left: InlineNode, right: InlineNode): boolean =>
  left.type === right.type &&
  structurallyEqual(Object.fromEntries(ownFields(left)), Object.fromEntries(ownFields(right)));

/**
 * How a merge treats records it cannot merge. `exact` fails the whole merge;
 * `asFarAsAlike` stops merging at the first pair that differs and leaves the
 * rest side by side. `identifiedFrom` is the first level at which a tracked
 * change or content control may not be merged: two of them are two revisions
 * or controls, and one record would carry only one id.
 */
export type MergeRule = { mode: "exact" | "asFarAsAlike"; identifiedFrom?: number };

export const EXACT_MERGE: MergeRule = Object.freeze({ mode: "exact" });

const mergeNode = (
  left: InlineNode,
  right: InlineNode,
  depth: number,
  rule: MergeRule,
  level: number,
): InlineNode | undefined => {
  // A record with nothing in it is a leaf of its own: merging it would lose it.
  const kind = kindOf(left);
  if (kind !== kindOf(right) || (kind !== "characters" && kind !== "branch")) {
    return undefined;
  }
  if (left.type === "text" && right.type === "text") {
    return depth === 1 ? { ...left, text: left.text + right.text } : undefined;
  }
  const leftChildren = childNodes(left);
  const rightChildren = childNodes(right);
  if (leftChildren === undefined || rightChildren === undefined || !sameOwnFields(left, right)) {
    return undefined;
  }
  if (rule.identifiedFrom !== undefined && level >= rule.identifiedFrom && isIdentifiedNode(left)) {
    return undefined;
  }
  const children = mergeLists(leftChildren, rightChildren, depth - 1, rule, level + 1);
  return children === undefined ? undefined : rebuild(left, children);
};

/**
 * Two lists end to end, the last record of the first merged with the first of
 * the second, and so on `depth` levels down. `undefined` when an exact merge
 * meets records that cannot be merged.
 */
export const mergeLists = (
  left: readonly InlineNode[],
  right: readonly InlineNode[],
  depth: number,
  rule: MergeRule = EXACT_MERGE,
  level = 1,
): InlineNode[] | undefined => {
  if (depth === 0) {
    return [...left, ...right];
  }
  const last = left.at(-1);
  const first = right.at(0);
  const merged =
    last === undefined || first === undefined
      ? undefined
      : mergeNode(last, first, depth, rule, level);
  if (merged === undefined) {
    return rule.mode === "asFarAsAlike" ? [...left, ...right] : undefined;
  }
  return [...left.slice(0, -1), merged, ...right.slice(1)];
};

/** The zero-width leaves at an offset, in document order. */
export const zeroWidthLeavesAt = (items: readonly InlineNode[], offset: number): InlineNode[] => {
  const out: InlineNode[] = [];
  const cursor = startCursor();
  const walk = (list: readonly InlineNode[]): void => {
    for (const node of list) {
      if (cursor.position > offset) {
        return;
      }
      const kind = kindOf(node);
      switch (kind) {
        case "characters":
          if (node.type === "text") cursor.position += node.text.length;
          break;
        case "unit":
          cursor.position += 1;
          break;
        case "zeroWidth":
          if (cursor.position === offset) out.push(node);
          break;
        case "branch":
          walk(childNodes(node) ?? []);
          break;
        default: {
          const unreachable: never = kind;
          return unreachable;
        }
      }
    }
  };
  walk(items);
  return out;
};

/**
 * The gap an insertion defaults to: before the first zero-width leaf at the
 * offset that opens a range, so a marker opening there stays with what
 * follows it and every other one with what precedes it.
 */
export const defaultInsertionGap = (items: readonly InlineNode[], offset: number): Gap => {
  const leaves = zeroWidthLeavesAt(items, offset);
  const opening = leaves.findIndex((leaf) => isParagraphContent(leaf) && isOpeningMarker(leaf));
  return { offset, zeroWidthBefore: opening === -1 ? leaves.length : opening };
};

/** Every run text in a list, nested ones included. */
export const textsIn = (items: readonly InlineNode[]): string[] => {
  const out: string[] = [];
  const walk = (list: readonly InlineNode[]): void => {
    for (const node of list) {
      if (node.type === "text") {
        out.push(node.text);
        continue;
      }
      const children = childNodes(node);
      if (children !== undefined) walk(children);
    }
  };
  walk(items);
  return out;
};

/** A record's gaps: the one before its first leaf and the one after its last. */
export type NodeGaps = { node: InlineNode; before: Gap; after: Gap };

const advance = (node: InlineNode, cursor: Cursor): void => {
  const kind = kindOf(node);
  switch (kind) {
    case "characters":
      if (node.type === "text") cursor.position += node.text.length;
      cursor.zeroWidthSeen = 0;
      return;
    case "unit":
      cursor.position += 1;
      cursor.zeroWidthSeen = 0;
      return;
    case "zeroWidth":
      cursor.zeroWidthSeen += 1;
      return;
    case "branch":
      for (const child of childNodes(node) ?? []) advance(child, cursor);
      return;
    default: {
      const unreachable: never = kind;
      return unreachable;
    }
  }
};

/**
 * Every run in document order with its gaps, runs inside hyperlinks, tracked
 * changes, content controls and wrappers included; runs inside a field are
 * part of the field's one unit and are not listed.
 */
export const runGaps = (items: readonly InlineNode[]): NodeGaps[] => {
  const out: NodeGaps[] = [];
  const cursor = startCursor();
  const walk = (list: readonly InlineNode[]): void => {
    for (const node of list) {
      if (node.type === "run") {
        const before = { offset: cursor.position, zeroWidthBefore: cursor.zeroWidthSeen };
        advance(node, cursor);
        out.push({
          node,
          before,
          after: { offset: cursor.position, zeroWidthBefore: cursor.zeroWidthSeen },
        });
        continue;
      }
      if (kindOf(node) === "branch") {
        walk(childNodes(node) ?? []);
        continue;
      }
      advance(node, cursor);
    }
  };
  walk(items);
  return out;
};
