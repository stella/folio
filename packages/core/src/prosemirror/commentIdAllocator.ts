/**
 * Per-document comment ID allocation for live editors and headless reviewers.
 * Loaded and issued IDs stay reserved even when a comment is deleted.
 */

import { TaggedError, type TaggedErrorClass } from "better-result";
import type { EditorView } from "prosemirror-view";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

import type { Comment } from "../types/content";
import {
  expectCommentRangeBoundaryAttrs,
  isCommentRangeBoundaryNode,
} from "./commentRangeBoundaryAttrs";
import {
  forEachAnnotationIdInDoc,
  seedRevisionIdsAbove,
} from "./plugins/revisionIds";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";

/** Sentinel ID for a comment that hasn't been persisted yet (anchored to selection). */
export const PENDING_COMMENT_ID = -1;

const CommentIdSpaceExhaustedErrorBase: TaggedErrorClass<"CommentIdSpaceExhaustedError"> =
  TaggedError("CommentIdSpaceExhaustedError");

export class CommentIdSpaceExhaustedError extends CommentIdSpaceExhaustedErrorBase<{
  message: string;
}> {}

export type CommentIdAllocator = {
  /** Allocate the next unused ID. */
  next(): number;
  /** Reserve a loaded ID and seed above it when possible. */
  seedAbove(id: number): void;
  /** An allocator holding this one's reservations, for a copy of its document. */
  clone(): CommentIdAllocator;
};

type AllocatorState = { nextId: number; reservedIds: Set<number> };

const allocatorOver = (state: AllocatorState): CommentIdAllocator => ({
  next: () => {
    if (state.reservedIds.size >= MAX_REVISION_ID) {
      throw new CommentIdSpaceExhaustedError({
        message: "The OOXML comment ID space is exhausted.",
      });
    }
    while (state.reservedIds.has(state.nextId)) {
      state.nextId = state.nextId === MAX_REVISION_ID ? 1 : state.nextId + 1;
    }
    const id = state.nextId;
    state.reservedIds.add(id);
    state.nextId = id === MAX_REVISION_ID ? 1 : id + 1;
    return id;
  },
  seedAbove: (id) => {
    if (!Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID) return;
    if (id > 0) state.reservedIds.add(id);
    if (id >= state.nextId) {
      state.nextId = id === MAX_REVISION_ID ? 1 : id + 1;
    }
  },
  clone: () =>
    allocatorOver({
      nextId: state.nextId,
      reservedIds: new Set(state.reservedIds),
    }),
});

/**
 * Comment IDs for one document: seeded from its comments and annotation ids,
 * then never reissued, also after a deletion. Each document holds its own, so
 * an ID depends only on that document and the edits made to it.
 */
export const createCommentIdAllocator = (): CommentIdAllocator =>
  allocatorOver({ nextId: 1, reservedIds: new Set() });

/** Seed from comments, replies, and the document's conservative annotation maximum. */
export function seedCommentAllocator(
  allocator: CommentIdAllocator,
  comments: Comment[] | undefined,
  view: Pick<EditorView, "state"> | null,
): void {
  let max = 0;
  for (const comment of comments ?? []) {
    allocator.seedAbove(comment.id);
    if (
      Number.isInteger(comment.id) &&
      comment.id > max &&
      comment.id <= MAX_REVISION_ID
    ) {
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
        if (attrs.ok)
          for (const marker of attrs.value.markers)
            if (marker.type === "commentRangeStart")
              allocator.seedAbove(marker.id);
      }
      if (isCommentRangeBoundaryNode(node)) {
        allocator.seedAbove(expectCommentRangeBoundaryAttrs(node).marker.id);
      }
    });
    // Every id the annotation space holds stays reserved, so an allocation
    // that wraps past the maximum never reuses one.
    forEachAnnotationIdInDoc(view.state.doc, (id) => {
      allocator.seedAbove(id);
      max = Math.max(max, id);
    });
  }
  allocator.seedAbove(max);
  seedRevisionIdsAbove(max);
}
