/**
 * One range per comment id, for a paste or a drop into a ProseMirror session.
 *
 * A slice copied from the document repeats the anchors of the comments it
 * covers: their marks, references and range boundaries. Pasted as they are,
 * a comment would own two ranges. So a pasted anchor of a comment the
 * document already names is a copy: when the session holds that comment's
 * thread and the slice holds some of its range, the slice's anchors take a
 * fresh comment id and the paste copies the thread under it, as a paste
 * copies a comment; otherwise those anchors are left out. Anchors of a
 * comment the document does not name (a cut pasted back, a paste from
 * elsewhere) go in as they are.
 *
 * A drag may move its source, and whether the source still stands is known
 * once the drop is done: what was dropped is settled where it stands, so a
 * moved comment keeps its id. Plain text takes the comments around it as
 * typing does: it carries none, and extends their ranges.
 */
import {
  Fragment,
  type Mark,
  type Node as PMNode,
  Slice,
} from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import type { Comment } from "../types/content";
import {
  isCommentRangeBoundaryNode,
  readCommentRangeBoundaryAttrs,
} from "./commentRangeBoundaryAttrs";
import { isCommentRangeInertAtom } from "./commentRangeInlineAtoms";
import { readCommentReferenceAttrs } from "./commentReferenceAttrs";
import { COMMENT_REFERENCE_NODE_NAME } from "./commentReferenceIntegrity";
import type { RangeMarker } from "./emptyRangeAnchor";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { forEachAnnotationIdInDoc } from "./plugins/revisionIds";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";

/** What a paste does with the comment anchors its slice carries. */
export type PastedCommentSettlement = {
  slice: Slice;
  /** Copied threads the paste adds to the records, roots before replies. */
  copies: readonly Comment[];
};

const commentMarkId = (mark: Mark): number | undefined => {
  const id: unknown = mark.attrs["commentId"];
  return mark.type.name === "comment" && typeof id === "number"
    ? id
    : undefined;
};

/** The comment ids a node names itself, its marks aside. */
const nodeCommentIds = (node: PMNode): number[] => {
  if (node.type.name === COMMENT_REFERENCE_NODE_NAME) {
    const attrs = readCommentReferenceAttrs(node);
    return attrs.ok ? [attrs.value.commentId] : [];
  }
  if (isCommentRangeBoundaryNode(node)) {
    const attrs = readCommentRangeBoundaryAttrs(node);
    return attrs.ok ? [attrs.value.marker.id] : [];
  }
  if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
    const attrs = readRangeAnchorAttrs(node);
    return attrs.ok
      ? attrs.value.markers.flatMap((marker) =>
          marker.type === "commentRangeStart" ||
          marker.type === "commentRangeEnd"
            ? [marker.id]
            : [],
        )
      : [];
  }
  return [];
};

/** The comments whose ranges a node alone holds some of. */
const rangeCommentIds = (node: PMNode): number[] => {
  if (node.type.name === COMMENT_REFERENCE_NODE_NAME) return [];
  if (isCommentRangeInertAtom(node)) return nodeCommentIds(node);
  return [
    ...nodeCommentIds(node),
    ...node.marks.flatMap((mark) => {
      // A point comment's mark only shows it: its reference is its anchor.
      const id = mark.attrs["point"] === true ? undefined : commentMarkId(mark);
      return id === undefined ? [] : [id];
    }),
  ];
};

type SliceComments = {
  named: Set<number>;
  /** The comments the slice holds some of the range of. */
  ranged: Set<number>;
};

const sliceComments = (
  fragment: Fragment,
  out: SliceComments = { named: new Set(), ranged: new Set() },
): SliceComments => {
  fragment.forEach((node) => {
    for (const mark of node.marks) {
      const id = commentMarkId(mark);
      if (id !== undefined) out.named.add(id);
    }
    for (const id of nodeCommentIds(node)) out.named.add(id);
    for (const id of rangeCommentIds(node)) out.ranged.add(id);
    sliceComments(node.content, out);
  });
  return out;
};

/** Which comments' anchors take another id, and which are left out. */
type AnchorChanges = {
  renamed: ReadonlyMap<number, number>;
  dropped: ReadonlySet<number>;
};

/** A comment mark's replacement; `null` leaves it out. */
const settledMark = (
  mark: Mark,
  { renamed, dropped }: AnchorChanges,
): Mark | null => {
  const id = commentMarkId(mark);
  if (id === undefined) return mark;
  if (dropped.has(id)) return null;
  const commentId = renamed.get(id);
  return commentId === undefined
    ? mark
    : // What proposed the comment does not propose its copy.
      mark.type.create({ ...mark.attrs, commentId, suggestionId: null });
};

/** The attributes of a node naming comments; `null` leaves the node out. */
const settledAttrs = (
  node: PMNode,
  { renamed, dropped }: AnchorChanges,
): PMNode["attrs"] | null => {
  if (node.type.name === COMMENT_REFERENCE_NODE_NAME) {
    const id = nodeCommentIds(node)[0];
    if (id !== undefined && dropped.has(id)) return null;
    const next = id === undefined ? undefined : renamed.get(id);
    return next === undefined ? node.attrs : { ...node.attrs, commentId: next };
  }
  if (isCommentRangeBoundaryNode(node)) {
    const read = readCommentRangeBoundaryAttrs(node);
    if (!read.ok) return node.attrs;
    const { marker } = read.value;
    if (dropped.has(marker.id)) return null;
    const next = renamed.get(marker.id);
    return next === undefined
      ? node.attrs
      : { ...node.attrs, marker: { ...marker, id: next } };
  }
  if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
    const read = readRangeAnchorAttrs(node);
    if (!read.ok) return node.attrs;
    // A comment's markers here are whole pairs, so the rest stay paired.
    const markers = read.value.markers.flatMap((marker): RangeMarker[] => {
      const comment =
        marker.type === "commentRangeStart" ||
        marker.type === "commentRangeEnd";
      if (!comment) return [marker];
      if (dropped.has(marker.id)) return [];
      const next = renamed.get(marker.id);
      return [next === undefined ? marker : { ...marker, id: next }];
    });
    if (markers.length === 0) return null;
    return markers.length === read.value.markers.length &&
      markers.every((marker, index) => marker === read.value.markers[index])
      ? node.attrs
      : { ...node.attrs, markers };
  }
  return node.attrs;
};

/** Rename and drop comment anchors throughout a fragment. */
const settleFragment = (
  fragment: Fragment,
  changes: AnchorChanges,
): Fragment => {
  const nodes: PMNode[] = [];
  fragment.forEach((node) => {
    const marks = node.marks.flatMap(
      (mark) => settledMark(mark, changes) ?? [],
    );
    if (node.isText) {
      nodes.push(node.mark(marks));
      return;
    }
    const attrs = settledAttrs(node, changes);
    if (attrs === null) return;
    const content = settleFragment(node.content, changes);
    nodes.push(
      attrs === node.attrs && content === node.content
        ? node.mark(marks)
        : node.type.create(attrs, content, marks),
    );
  });
  return Fragment.fromArray(nodes);
};

/** A copy of a comment; its paragraph identities are minted on save. */
const copyOf = (comment: Comment, id: number): Comment => {
  const copy = structuredClone(comment);
  copy.id = id;
  // What proposed the comment does not propose its copy.
  delete copy.suggestionId;
  // The clone's paragraphs are the copy's own, so they drop their ids here.
  for (const paragraph of copy.content) {
    delete paragraph.paraId;
    delete paragraph.textId;
  }
  return copy;
};

/** What a paste settles its slice against: the session it goes into. */
export type PasteTarget = {
  doc: PMNode;
  /** The session's comment records; `null` when it holds none. */
  records: readonly Comment[] | null;
  /** The comment ids `doc` names through a mark, reference or boundary. */
  named: ReadonlySet<number>;
};

/** A comment and the replies whose parent chain reaches it, in record order. */
const threadOf = (records: readonly Comment[], id: number): Comment[] => {
  const ids = new Set([id]);
  let added = true;
  while (added) {
    added = false;
    for (const comment of records)
      if (
        !ids.has(comment.id) &&
        comment.parentId !== undefined &&
        ids.has(comment.parentId)
      ) {
        ids.add(comment.id);
        added = true;
      }
  }
  return records.filter((comment) => ids.has(comment.id));
};

type CopyThreadsOptions = {
  /** The comments to copy. */
  ids: Iterable<number>;
  records: readonly Comment[] | null;
  doc: PMNode;
  /** Ids a copy may not take, besides the records' and the document's. */
  taken: Iterable<number>;
};

/** Copies of the threads the session holds, under fresh ids. */
const copyThreads = ({
  ids: roots,
  records,
  doc,
  taken,
}: CopyThreadsOptions): {
  renamed: Map<number, number>;
  /** The comments without a thread to copy. */
  dropped: Set<number>;
  copies: Comment[];
} => {
  // Fresh ids sit above every id the records, the document and the slice
  // give meaning to, as the session's allocator seeds itself.
  let next = 0;
  for (const comment of records ?? [])
    next = Math.max(next, comment.id, comment.parentId ?? 0);
  forEachAnnotationIdInDoc(doc, (id) => {
    next = Math.max(next, id);
  });
  for (const id of taken) next = Math.max(next, id);
  const renamed = new Map<number, number>();
  const dropped = new Set<number>();
  const copies: Comment[] = [];
  for (const id of roots) {
    const thread =
      records?.some((comment) => comment.id === id) === true
        ? threadOf(records, id)
        : [];
    if (thread.length === 0 || next + thread.length > MAX_REVISION_ID) {
      dropped.add(id);
      continue;
    }
    const ids = new Map<number, number>();
    for (const comment of thread) ids.set(comment.id, ++next);
    for (const comment of thread) {
      const copy = copyOf(comment, ids.get(comment.id) ?? comment.id);
      if (comment.id === id) delete copy.parentId;
      else if (comment.parentId !== undefined)
        copy.parentId = ids.get(comment.parentId) ?? comment.parentId;
      copies.push(copy);
    }
    renamed.set(id, ids.get(id) ?? id);
  }
  return { renamed, dropped, copies };
};

/**
 * Settle the comment anchors a pasted `slice` carries against the session
 * it goes into; see the module comment.
 */
export const settlePastedComments = (
  slice: Slice,
  { doc, records, named }: PasteTarget,
): PastedCommentSettlement => {
  const pasted = sliceComments(slice.content);
  const repeated = [...pasted.named].filter((id) => named.has(id));
  if (repeated.length === 0) return { slice, copies: [] };
  const { renamed, dropped, copies } = copyThreads({
    ids: repeated.filter((id) => pasted.ranged.has(id)),
    records,
    doc,
    taken: pasted.named,
  });
  // A reference without its range would be a range covering nothing.
  for (const id of repeated) if (!pasted.ranged.has(id)) dropped.add(id);
  return {
    slice: new Slice(
      settleFragment(slice.content, { renamed, dropped }),
      slice.openStart,
      slice.openEnd,
    ),
    copies,
  };
};

/** Positions in a document, end exclusive. */
export type DroppedRange = { from: number; to: number };

/** The comment anchors on one side of a drop. */
type AnchoredComments = { ranged: Set<number>; referenced: Set<number> };

type DroppedAnchors = {
  dropped: AnchoredComments;
  elsewhere: AnchoredComments;
};

const droppedAnchors = (
  doc: PMNode,
  ranges: readonly DroppedRange[],
): DroppedAnchors => {
  const anchors: DroppedAnchors = {
    dropped: { ranged: new Set(), referenced: new Set() },
    elsewhere: { ranged: new Set(), referenced: new Set() },
  };
  doc.descendants((node, pos) => {
    const ranged = rangeCommentIds(node);
    const referenced =
      node.type.name === COMMENT_REFERENCE_NODE_NAME
        ? nodeCommentIds(node)
        : [];
    if (ranged.length === 0 && referenced.length === 0) return true;
    const end = pos + node.nodeSize;
    let covered = 0;
    for (const { from, to } of ranges)
      covered += Math.max(0, Math.min(to, end) - Math.max(from, pos));
    const sides = [
      ...(covered > 0 ? [anchors.dropped] : []),
      ...(covered < node.nodeSize ? [anchors.elsewhere] : []),
    ];
    for (const side of sides) {
      for (const id of ranged) side.ranged.add(id);
      for (const id of referenced) side.referenced.add(id);
    }
    return true;
  });
  return anchors;
};

type DroppedCommentsOptions = {
  /** Where the dropped content stands in the transaction's document. */
  ranges: readonly DroppedRange[];
  records: readonly Comment[] | null;
};

/**
 * Settle the comment anchors dropped into `transaction`'s document that its
 * comments have standing elsewhere in it too, and return the copied threads
 * to add to the records; `null` when none of them do.
 */
export const settleDroppedComments = (
  transaction: Transaction,
  { ranges, records }: DroppedCommentsOptions,
): readonly Comment[] | null => {
  const { doc } = transaction;
  const { dropped: inside, elsewhere } = droppedAnchors(doc, ranges);
  const copied = [...inside.ranged].filter((id) => elsewhere.ranged.has(id));
  const repeated = [...inside.referenced].filter(
    (id) => !copied.includes(id) && elsewhere.referenced.has(id),
  );
  if (copied.length === 0 && repeated.length === 0) return null;
  const { renamed, dropped, copies } = copyThreads({
    ids: copied,
    records,
    doc,
    taken: [],
  });
  for (const id of repeated) dropped.add(id);
  const changes: AnchorChanges = { renamed, dropped };
  const removed: DroppedRange[] = [];
  for (const { from, to } of ranges)
    doc.nodesBetween(from, to, (node, pos) => {
      for (const mark of node.marks) {
        const next = settledMark(mark, changes);
        if (next === mark) continue;
        const start = Math.max(from, pos);
        const end = Math.min(to, pos + node.nodeSize);
        transaction.removeMark(start, end, mark);
        if (next !== null) transaction.addMark(start, end, next);
      }
      const attrs = node.isText ? node.attrs : settledAttrs(node, changes);
      if (attrs === null) removed.push({ from: pos, to: pos + node.nodeSize });
      else
        for (const [name, value] of Object.entries(attrs))
          if (value !== node.attrs[name])
            transaction.setNodeAttribute(pos, name, value);
      return true;
    });
  // Marks and attributes move nothing: the removals go last, last first.
  for (const { from, to } of removed.toReversed()) transaction.delete(from, to);
  return copies;
};
