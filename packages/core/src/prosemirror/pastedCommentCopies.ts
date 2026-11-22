/**
 * One range per comment id, for a paste into a ProseMirror session.
 *
 * A slice copied from the document repeats the anchors of the comments it
 * covers: their marks, references and range boundaries. Pasted as they are,
 * a comment would own two ranges. So a pasted anchor of a comment the
 * document already names is a copy: when the session holds that comment's
 * thread, the slice's anchors take a fresh comment id and the paste copies
 * the thread under it, as a paste copies a comment; otherwise those anchors
 * are left out. Anchors of a comment the document does not name (a cut
 * pasted back, a paste from elsewhere) go in as they are.
 */
import {
  Fragment,
  type Mark,
  type Node as PMNode,
  Slice,
} from "prosemirror-model";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import type { Comment } from "../types/content";
import {
  isCommentRangeBoundaryNode,
  readCommentRangeBoundaryAttrs,
} from "./commentRangeBoundaryAttrs";
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

const sliceCommentIds = (fragment: Fragment, out = new Set<number>()) => {
  fragment.forEach((node) => {
    for (const mark of node.marks) {
      const id = commentMarkId(mark);
      if (id !== undefined) out.add(id);
    }
    for (const id of nodeCommentIds(node)) out.add(id);
    sliceCommentIds(node.content, out);
  });
  return out;
};

/** Rename and drop comment anchors throughout a fragment. */
const settleFragment = (
  fragment: Fragment,
  renamed: ReadonlyMap<number, number>,
  dropped: ReadonlySet<number>,
): Fragment => {
  const nodes: PMNode[] = [];
  fragment.forEach((node) => {
    const marks = node.marks.flatMap((mark) => {
      const id = commentMarkId(mark);
      if (id === undefined) return [mark];
      if (dropped.has(id)) return [];
      const next = renamed.get(id);
      return [
        next === undefined
          ? mark
          : mark.type.create({ ...mark.attrs, commentId: next }),
      ];
    });
    let attrs = node.attrs;
    if (node.type.name === COMMENT_REFERENCE_NODE_NAME) {
      const id = nodeCommentIds(node)[0];
      if (id !== undefined && dropped.has(id)) return;
      const next = id === undefined ? undefined : renamed.get(id);
      if (next !== undefined) attrs = { ...attrs, commentId: next };
    } else if (isCommentRangeBoundaryNode(node)) {
      const read = readCommentRangeBoundaryAttrs(node);
      if (read.ok) {
        const { marker } = read.value;
        if (dropped.has(marker.id)) return;
        const next = renamed.get(marker.id);
        if (next !== undefined)
          attrs = { ...attrs, marker: { ...marker, id: next } };
      }
    } else if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
      const read = readRangeAnchorAttrs(node);
      if (read.ok) {
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
        if (markers.length === 0) return;
        if (
          markers.length !== read.value.markers.length ||
          markers.some((marker, index) => marker !== read.value.markers[index])
        )
          attrs = { ...attrs, markers };
      }
    }
    if (node.isText) {
      nodes.push(node.mark(marks));
      return;
    }
    const content = settleFragment(node.content, renamed, dropped);
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

/**
 * Settle the comment anchors a pasted `slice` carries against the session
 * it goes into; see the module comment.
 */
export const settlePastedComments = (
  slice: Slice,
  { doc, records, named }: PasteTarget,
): PastedCommentSettlement => {
  const pasted = sliceCommentIds(slice.content);
  const repeated = [...pasted].filter((id) => named.has(id));
  if (repeated.length === 0) return { slice, copies: [] };
  // Fresh ids sit above every id the records, the document and the slice
  // give meaning to, as the session's allocator seeds itself.
  let next = 0;
  for (const comment of records ?? [])
    next = Math.max(next, comment.id, comment.parentId ?? 0);
  forEachAnnotationIdInDoc(doc, (id) => {
    next = Math.max(next, id);
  });
  for (const id of pasted) next = Math.max(next, id);
  const renamed = new Map<number, number>();
  const dropped = new Set<number>();
  const copies: Comment[] = [];
  for (const id of repeated) {
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
  return {
    slice: new Slice(
      settleFragment(slice.content, renamed, dropped),
      slice.openStart,
      slice.openEnd,
    ),
    copies,
  };
};
