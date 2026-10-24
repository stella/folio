import { panic } from "better-result";

import type {
  BlockContent,
  BookmarkEnd,
  BookmarkStart,
  MoveFromRangeEnd,
  MoveFromRangeStart,
  MoveToRangeEnd,
  MoveToRangeStart,
  ParagraphContent,
  TrackedRunContent,
} from "../model/document";
import {
  childNodes,
  type InlineNode,
  isParagraphContent,
  rebuildNode,
} from "./leaves";

export const MAX_COPY_ORIGIN_ANCESTORS = 256;

export type CopyOriginAncestorList = NonNullable<
  BookmarkEnd["copyOriginAncestors"]
>;

type CopyOriginBoundary =
  | BookmarkStart
  | BookmarkEnd
  | MoveFromRangeStart
  | MoveFromRangeEnd
  | MoveToRangeStart
  | MoveToRangeEnd;

const isCopyOriginBoundary = (node: InlineNode): node is CopyOriginBoundary =>
  node.type === "bookmarkStart" ||
  node.type === "bookmarkEnd" ||
  node.type === "moveFromRangeStart" ||
  node.type === "moveFromRangeEnd" ||
  node.type === "moveToRangeStart" ||
  node.type === "moveToRangeEnd";

export const copyOriginAncestorsOf = (
  node: InlineNode,
): CopyOriginAncestorList | undefined =>
  isCopyOriginBoundary(node) ? node.copyOriginAncestors : undefined;

const TRACKED_RUN_CONTENT_TYPES = {
  run: true,
  hyperlink: true,
  bookmarkStart: true,
  bookmarkEnd: true,
  commentRangeStart: true,
  commentRangeEnd: true,
  commentReference: true,
  moveFromRangeStart: true,
  moveFromRangeEnd: true,
  moveToRangeStart: true,
  moveToRangeEnd: true,
  simpleField: true,
  complexField: true,
  inlineSdt: true,
  inlineWrapper: true,
  mathEquation: true,
  preservedInline: true,
  insertion: true,
  deletion: true,
  moveFrom: true,
  moveTo: true,
} as const satisfies Record<TrackedRunContent["type"], true>;

export const isTrackedRunContent = (
  node: InlineNode,
): node is TrackedRunContent =>
  Object.hasOwn(TRACKED_RUN_CONTENT_TYPES, node.type);

export const copyOriginBoundaryKey = (node: InlineNode): string | undefined =>
  isCopyOriginBoundary(node) ? `${node.type}:${node.id}` : undefined;

export const cloneIncomingCopyBoundaries = (
  nodes: readonly ParagraphContent[],
  existingBoundaryKeys: ReadonlySet<string>,
): { content: ParagraphContent[]; copiedBoundaries: Set<InlineNode> } => {
  const copiedBoundaries = new Set<InlineNode>();
  const visit = (current: readonly InlineNode[]): InlineNode[] => {
    const out: InlineNode[] = [];
    for (const node of current) {
      const key = copyOriginBoundaryKey(node);
      if (key !== undefined && existingBoundaryKeys.has(key)) {
        const copy = { ...node };
        copiedBoundaries.add(copy);
        out.push(copy);
        continue;
      }
      const children = childNodes(node);
      if (children === undefined) {
        out.push(node);
        continue;
      }
      const nextChildren = visit(children);
      out.push(
        nextChildren.length === children.length &&
          nextChildren.every((child, index) => child === children[index])
          ? node
          : rebuildNode(node, nextChildren),
      );
    }
    return out;
  };
  const content: ParagraphContent[] = [];
  for (const node of visit(nodes)) {
    if (!isParagraphContent(node))
      return panic("A copied range boundary left paragraph content.");
    content.push(node);
  }
  return { content, copiedBoundaries };
};

const validOrdinals = (ordinals: readonly number[]): boolean => {
  if (ordinals.length === 0 || ordinals.length > MAX_COPY_ORIGIN_ANCESTORS)
    return false;
  for (let index = 0; index < ordinals.length; index += 1) {
    const ordinal = ordinals[index];
    const previous = ordinals[index - 1];
    if (
      typeof ordinal !== "number" ||
      !Number.isSafeInteger(ordinal) ||
      ordinal < 0 ||
      (index > 0 && (previous === undefined || previous >= ordinal))
    )
      return false;
  }
  return true;
};

export const isCopyOriginAncestors = (
  value: unknown,
): value is CopyOriginAncestorList =>
  Array.isArray(value) && validOrdinals(value);

const nonEmptyOrdinals = (
  ordinals: readonly number[],
): CopyOriginAncestorList | undefined => {
  const first = ordinals.at(0);
  return first === undefined ? undefined : [first, ...ordinals.slice(1)];
};

export const shiftCopyOriginAncestors = (
  origins: CopyOriginAncestorList | undefined,
  insertedOrdinal: number,
): CopyOriginAncestorList | undefined => {
  if (
    !Number.isSafeInteger(insertedOrdinal) ||
    insertedOrdinal < 0 ||
    (origins !== undefined && !validOrdinals(origins))
  )
    return panic("A copied-boundary nesting has invalid ancestry.");
  if (origins === undefined) return undefined;
  const shifted = origins.map((ordinal) =>
    ordinal < insertedOrdinal ? ordinal : ordinal + 1,
  );
  if (!validOrdinals(shifted)) return undefined;
  return nonEmptyOrdinals(shifted);
};

export const addCopyOriginAncestor = (
  origins: CopyOriginAncestorList | undefined,
  addedOrdinal: number,
): CopyOriginAncestorList | undefined => {
  const shifted = shiftCopyOriginAncestors(origins, addedOrdinal);
  if (origins !== undefined && shifted === undefined) return undefined;
  const next = [...(shifted ?? []), addedOrdinal].sort(
    (left, right) => left - right,
  );
  if (
    next.length > MAX_COPY_ORIGIN_ANCESTORS ||
    new Set(next).size !== next.length
  )
    return undefined;
  return nonEmptyOrdinals(next);
};

export const acceptCopyOriginAncestors = (
  origins: CopyOriginAncestorList | undefined,
  acceptedOrdinal: number,
): CopyOriginAncestorList | undefined => {
  if (
    !Number.isSafeInteger(acceptedOrdinal) ||
    acceptedOrdinal < 0 ||
    (origins !== undefined && !validOrdinals(origins))
  )
    return panic("A copied-boundary acceptance has invalid ancestry.");
  if (origins === undefined) return undefined;
  return nonEmptyOrdinals(
    origins.flatMap((ordinal) =>
      ordinal === acceptedOrdinal
        ? []
        : [ordinal > acceptedOrdinal ? ordinal - 1 : ordinal],
    ),
  );
};

export type RejectedCopyOriginAncestors =
  | { type: "drop" }
  | {
      type: "retain";
      copyOriginAncestors?: CopyOriginAncestorList;
    };

export const rejectCopyOriginAncestors = (
  origins: CopyOriginAncestorList | undefined,
  removedAddedAncestors: number,
): RejectedCopyOriginAncestors => {
  if (
    !Number.isSafeInteger(removedAddedAncestors) ||
    removedAddedAncestors < 0 ||
    (origins !== undefined && !validOrdinals(origins))
  )
    return panic("A copied-boundary rejection has invalid ancestry.");
  if (origins === undefined) return { type: "retain" };
  if (origins.some((ordinal) => ordinal < removedAddedAncestors))
    return { type: "drop" };
  const copyOriginAncestors = nonEmptyOrdinals(
    origins.map((ordinal) => ordinal - removedAddedAncestors),
  );
  return copyOriginAncestors === undefined
    ? { type: "drop" }
    : { type: "retain", copyOriginAncestors };
};

const isAddedWrapper = (node: InlineNode): boolean =>
  node.type === "insertion" || node.type === "moveTo";

export const hasValidCopyOriginAncestors = (
  nodes: readonly InlineNode[],
  outerAddedAncestors = 0,
): boolean => {
  if (!Number.isSafeInteger(outerAddedAncestors) || outerAddedAncestors < 0)
    return false;
  for (const node of nodes) {
    if (
      isCopyOriginBoundary(node) &&
      node.copyOriginAncestors !== undefined &&
      (!validOrdinals(node.copyOriginAncestors) ||
        node.copyOriginAncestors.some(
          (ordinal) => ordinal >= outerAddedAncestors,
        ))
    )
      return false;
    const children = childNodes(node);
    if (
      children !== undefined &&
      !hasValidCopyOriginAncestors(
        children,
        outerAddedAncestors + Number(isAddedWrapper(node)),
      )
    )
      return false;
  }
  return true;
};

export const hasValidCopyOriginAncestorsInBlocks = (
  blocks: readonly BlockContent[],
): boolean =>
  blocks.every((block) => {
    switch (block.type) {
      case "paragraph":
        return hasValidCopyOriginAncestors(block.content);
      case "table":
        return block.rows.every((row) =>
          row.cells.every((cell) =>
            hasValidCopyOriginAncestorsInBlocks(cell.content),
          ),
        );
      case "blockSdt":
      case "blockCustomXml":
        return hasValidCopyOriginAncestorsInBlocks(block.content);
      case "bookmarkStart":
      case "bookmarkEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return block.copyOriginAncestors === undefined;
      case "preservedBlock":
        return true;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

export const copyOriginBoundaryKeysInBlocks = (
  blocks: readonly BlockContent[],
): Set<string> => {
  const keys = new Set<string>();
  const visitInline = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      const key = copyOriginBoundaryKey(node);
      if (key !== undefined) keys.add(key);
      visitInline(childNodes(node) ?? []);
    }
  };
  const visitBlocks = (current: readonly BlockContent[]): void => {
    for (const block of current) {
      switch (block.type) {
        case "paragraph":
          visitInline(block.content);
          break;
        case "table":
          for (const row of block.rows)
            for (const cell of row.cells) visitBlocks(cell.content);
          break;
        case "blockSdt":
        case "blockCustomXml":
          visitBlocks(block.content);
          break;
        case "bookmarkStart":
        case "bookmarkEnd":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd": {
          const key = copyOriginBoundaryKey(block);
          if (key !== undefined) keys.add(key);
          break;
        }
        case "preservedBlock":
          break;
        default: {
          const unreachable: never = block;
          return unreachable;
        }
      }
    }
  };
  visitBlocks(blocks);
  return keys;
};

/** Adds one inserted revision around `nodes`, preserving exact owner depths. */
export const addCopyOriginAncestorsInNodes = (
  nodes: readonly InlineNode[],
  copiedBoundaries: Set<InlineNode>,
): InlineNode[] | undefined => {
  const visit = (
    current: readonly InlineNode[],
    addedDescendants: number,
  ): InlineNode[] | undefined => {
    const out: InlineNode[] = [];
    for (const node of current) {
      if (isCopyOriginBoundary(node)) {
        const origins = node.copyOriginAncestors ?? [];
        if (node.copyOriginAncestors !== undefined && !validOrdinals(origins))
          return undefined;
        const copied = copiedBoundaries.has(node);
        const nextOrigins = copied
          ? addCopyOriginAncestor(node.copyOriginAncestors, addedDescendants)
          : shiftCopyOriginAncestors(
              node.copyOriginAncestors,
              addedDescendants,
            );
        if (
          (copied || node.copyOriginAncestors !== undefined) &&
          nextOrigins === undefined
        )
          return undefined;
        if (
          nextOrigins === undefined ||
          (node.copyOriginAncestors !== undefined &&
            nextOrigins.length === node.copyOriginAncestors.length &&
            nextOrigins.every(
              (ordinal, index) => ordinal === node.copyOriginAncestors?.[index],
            ))
        ) {
          out.push(node);
          continue;
        }
        const copy = { ...node, copyOriginAncestors: nextOrigins };
        if (copiedBoundaries.delete(node)) copiedBoundaries.add(copy);
        out.push(copy);
        continue;
      }
      const children = childNodes(node);
      if (children === undefined) {
        out.push(node);
        continue;
      }
      const nextChildren = visit(
        children,
        addedDescendants + Number(isAddedWrapper(node)),
      );
      if (nextChildren === undefined) return undefined;
      if (
        nextChildren.length === children.length &&
        nextChildren.every((child, index) => child === children[index])
      ) {
        out.push(node);
        continue;
      }
      out.push(rebuildNode(node, nextChildren));
    }
    return out;
  };
  return visit(nodes, 0);
};

/** Records copied boundaries under an existing revision without shifting ancestry. */
export const addCopyOriginToExistingAncestor = (
  nodes: readonly InlineNode[],
  copiedBoundaries: Set<InlineNode>,
): InlineNode[] | undefined => {
  const visit = (
    current: readonly InlineNode[],
    nestedAddedAncestors: number,
  ): InlineNode[] | undefined => {
    const out: InlineNode[] = [];
    for (const node of current) {
      if (isCopyOriginBoundary(node)) {
        const origins = node.copyOriginAncestors ?? [];
        if (node.copyOriginAncestors !== undefined && !validOrdinals(origins))
          return undefined;
        if (
          !copiedBoundaries.has(node) ||
          origins.some((origin) => origin === nestedAddedAncestors)
        ) {
          out.push(node);
          continue;
        }
        const nextOrigins = [...origins, nestedAddedAncestors].sort(
          (left, right) => left - right,
        );
        if (nextOrigins.length > MAX_COPY_ORIGIN_ANCESTORS) return undefined;
        const copyOriginAncestors = nonEmptyOrdinals(nextOrigins);
        if (copyOriginAncestors === undefined) return undefined;
        const copy = { ...node, copyOriginAncestors };
        if (copiedBoundaries.delete(node)) copiedBoundaries.add(copy);
        out.push(copy);
        continue;
      }
      const children = childNodes(node);
      if (children === undefined) {
        out.push(node);
        continue;
      }
      const nextChildren = visit(
        children,
        nestedAddedAncestors + Number(isAddedWrapper(node)),
      );
      if (nextChildren === undefined) return undefined;
      out.push(
        nextChildren.length === children.length &&
          nextChildren.every((child, index) => child === children[index])
          ? node
          : rebuildNode(node, nextChildren),
      );
    }
    return out;
  };
  return visit(nodes, 0);
};

export const acceptCopyOriginAncestor = (
  nodes: readonly InlineNode[],
  copiedBoundaries?: Set<InlineNode>,
): InlineNode[] => {
  const visit = (
    current: readonly InlineNode[],
    addedDescendants: number,
  ): InlineNode[] => {
    const out: InlineNode[] = [];
    for (const node of current) {
      if (isCopyOriginBoundary(node)) {
        const origins = node.copyOriginAncestors;
        if (origins === undefined) {
          out.push(node);
          continue;
        }
        if (!validOrdinals(origins))
          return panic("A copied boundary has invalid ancestry.");
        const nextOrigins = acceptCopyOriginAncestors(
          origins,
          addedDescendants,
        );
        const next = { ...node };
        if (nextOrigins === undefined) delete next.copyOriginAncestors;
        else next.copyOriginAncestors = nextOrigins;
        if (copiedBoundaries?.delete(node)) copiedBoundaries.add(next);
        out.push(next);
        continue;
      }
      const children = childNodes(node);
      if (children === undefined) {
        out.push(node);
        continue;
      }
      const nextChildren = visit(
        children,
        addedDescendants + Number(isAddedWrapper(node)),
      );
      out.push(
        nextChildren.every((child, index) => child === children[index])
          ? node
          : rebuildNode(node, nextChildren),
      );
    }
    return out;
  };
  return visit(nodes, 0);
};

export const rejectCopyOriginAncestorsInNodes = (
  nodes: readonly InlineNode[],
  removedAddedAncestors: number,
): InlineNode[] => {
  const visit = (
    current: readonly InlineNode[],
    removedAncestors: number,
  ): InlineNode[] => {
    const out: InlineNode[] = [];
    for (const node of current) {
      if (isCopyOriginBoundary(node)) {
        const origins = node.copyOriginAncestors;
        if (origins === undefined) {
          out.push(node);
          continue;
        }
        if (!validOrdinals(origins))
          return panic("A copied boundary has invalid ancestry.");
        const result = rejectCopyOriginAncestors(origins, removedAncestors);
        if (result.type === "drop") continue;
        if (result.copyOriginAncestors === undefined) {
          out.push(node);
          continue;
        }
        const next = {
          ...node,
          copyOriginAncestors: result.copyOriginAncestors,
        };
        out.push(next);
        continue;
      }
      const children = childNodes(node);
      if (children === undefined) {
        out.push(node);
        continue;
      }
      const nextChildren = visit(
        children,
        removedAncestors + Number(isAddedWrapper(node)),
      );
      out.push(
        nextChildren.length === children.length &&
          nextChildren.every((child, index) => child === children[index])
          ? node
          : rebuildNode(node, nextChildren),
      );
    }
    return out;
  };
  return visit(nodes, removedAddedAncestors);
};
