/**
 * Realm-wide comment ID allocation, shared by live editors and headless reviewers.
 * Loaded and issued IDs stay reserved even when a comment is deleted.
 */

import { TaggedError } from "better-result";
import type { EditorView } from "prosemirror-view";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";
import type { Comment } from "../types/content";
import { maxAnnotationIdInDoc, seedRevisionIdsAbove } from "./plugins/revisionIds";

/** Sentinel ID for a comment that hasn't been persisted yet (anchored to selection). */
export const PENDING_COMMENT_ID = -1;

export class CommentIdSpaceExhaustedError extends TaggedError("CommentIdSpaceExhaustedError")<{
  message: string;
}> {}

let nextId = 1;
const reservedIds = new Set<number>();

/** Allocate a valid, unused comment ID in this realm. */
export function allocateCommentId(): number {
  if (reservedIds.size >= MAX_REVISION_ID) {
    throw new CommentIdSpaceExhaustedError({ message: "The OOXML comment ID space is exhausted." });
  }
  while (reservedIds.has(nextId)) {
    nextId = nextId === MAX_REVISION_ID ? 1 : nextId + 1;
  }
  const id = nextId;
  reservedIds.add(id);
  nextId = id === MAX_REVISION_ID ? 1 : id + 1;
  return id;
}

/** Reserve a loaded ID and advance past it when the upper bound permits. */
export function seedCommentIdAbove(id: number): void {
  if (!Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID) {
    return;
  }
  if (id > 0) reservedIds.add(id);
  if (id >= nextId) {
    nextId = id === MAX_REVISION_ID ? 1 : id + 1;
  }
}

export type CommentIdAllocator = {
  /** Allocate the next unused ID. */
  next(): number;
  /** Reserve a loaded ID and seed above it when possible. */
  seedAbove(id: number): void;
};

/** Every handle uses the same counter, including separate editor instances. */
export function createCommentIdAllocator(): CommentIdAllocator {
  return { next: allocateCommentId, seedAbove: seedCommentIdAbove };
}

/** Seed from comments, replies, and the document's conservative annotation maximum. */
export function seedCommentAllocator(
  allocator: CommentIdAllocator,
  comments: Comment[] | undefined,
  view: Pick<EditorView, "state"> | null,
): void {
  let max = 0;
  for (const comment of comments ?? []) {
    allocator.seedAbove(comment.id);
    if (Number.isInteger(comment.id) && comment.id > max && comment.id <= MAX_REVISION_ID) {
      max = comment.id;
    }
  }
  if (view) {
    view.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        const id: unknown = mark.attrs["commentId"];
        if (typeof id === "number") allocator.seedAbove(id);
      }
      if (node.type.name === "commentReference") {
        const id: unknown = node.attrs["commentId"];
        if (typeof id === "number") allocator.seedAbove(id);
      }
      if (node.type.name === "rangeAnchor") {
        const attrs = readRangeAnchorAttrs(node);
        if (attrs.ok && attrs.value.start.type === "commentRangeStart") {
          allocator.seedAbove(attrs.value.start.id);
        }
      }
    });
    max = Math.max(max, maxAnnotationIdInDoc(view.state.doc));
  }
  allocator.seedAbove(max);
  seedRevisionIdsAbove(max);
}
