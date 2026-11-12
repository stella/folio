import { Result } from "better-result";

import type { Document, TrackedChangeInfo } from "../model/document";
import { applyDocumentOp } from "./apply";
import type { DocumentOpRefusal } from "./refusal";
import { documentStories, storyBody } from "./stories";
import type { OpStory } from "./types";
import { DOCUMENT_OP_TYPES, REVISION_DECISIONS } from "./types";

/**
 * The ids of the revisions in a value that belong to a pending suggestion:
 * tracked insertions and deletions, moves, property changes and paragraph
 * mark revisions, at any depth.
 */
const suggestedRevisionIdsIn = (value: unknown, out: Set<number>): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) suggestedRevisionIdsIn(item, out);
    return;
  }
  const record = value as Record<string, unknown>;
  const info = record["info"];
  if (
    typeof info === "object" &&
    info !== null &&
    typeof (info as Record<string, unknown>)["suggestionId"] === "string" &&
    typeof (info as Record<string, unknown>)["id"] === "number"
  )
    out.add((info as { id: number }).id);
  for (const child of Object.values(record)) suggestedRevisionIdsIn(child, out);
};

/**
 * The ids of the comments pending suggestions propose: every one, or those of
 * one suggestion. A reply is removed with the comment it answers.
 */
export const suggestedCommentIds = (
  document: Document,
  suggestionId?: string,
): number[] =>
  (document.package.document.comments ?? []).flatMap((comment) =>
    comment.suggestionId !== undefined &&
    (suggestionId === undefined || comment.suggestionId === suggestionId)
      ? [comment.id]
      : [],
  );

/** Remove proposed comments with their ranges and replies. */
export const withoutSuggestedComments = (
  document: Document,
  ids: readonly number[],
): Result<Document, DocumentOpRefusal> => {
  let current = document;
  for (const id of ids) {
    const removed = applyDocumentOp(current, {
      type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
      id,
      scope: "thread",
    });
    if (removed.isErr()) return Result.err(removed.error);
    current = removed.value.document;
  }
  return Result.ok(current);
};

/**
 * The document as a save writes it: every pending suggestion rejected. A
 * suggestion is a proposal the host holds, not a change in the file, so the
 * file reads as if it had never been made; ordinary tracked changes stay. A
 * comment a suggestion proposes is left out with it.
 */
export const withoutPendingSuggestions = (
  document: Document,
): Result<Document, DocumentOpRefusal> => {
  const uncommented = withoutSuggestedComments(
    document,
    suggestedCommentIds(document),
  );
  if (uncommented.isErr()) return uncommented;
  let current = uncommented.value;
  for (const story of documentStories(current)) {
    const ids = new Set<number>();
    suggestedRevisionIdsIn(storyBody(current, story), ids);
    if (ids.size === 0) continue;
    const rejected = applyDocumentOp(current, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story,
      revisionIds: [...ids],
      decision: REVISION_DECISIONS.REJECT,
    });
    if (rejected.isErr()) return Result.err(rejected.error);
    current = rejected.value.document;
  }
  return Result.ok(current);
};

/**
 * The revisions of one pending suggestion in a story, one info per revision
 * id, in document order.
 */
export const suggestionRevisionInfos = (
  document: Document,
  story: OpStory,
  suggestionId: string,
): TrackedChangeInfo[] => {
  const out = new Map<number, TrackedChangeInfo>();
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    const info = record["info"] as Partial<TrackedChangeInfo> | undefined;
    if (
      typeof info === "object" &&
      info !== null &&
      info.suggestionId === suggestionId &&
      typeof info.id === "number" &&
      !out.has(info.id)
    )
      out.set(info.id, info as TrackedChangeInfo);
    for (const child of Object.values(record)) visit(child);
  };
  visit(storyBody(document, story).content);
  return [...out.values()];
};

/**
 * The ids of every pending suggestion in a document, in every story, and of
 * every suggestion that proposes a comment.
 */
export const pendingSuggestionIdsIn = (document: Document): Set<string> => {
  const out = new Set<string>();
  for (const comment of document.package.document.comments ?? [])
    if (comment.suggestionId !== undefined) out.add(comment.suggestionId);
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    const record = value as Record<string, unknown>;
    const info = record["info"] as { suggestionId?: unknown } | undefined;
    if (
      typeof info === "object" &&
      info !== null &&
      typeof info.suggestionId === "string"
    )
      out.add(info.suggestionId);
    for (const child of Object.values(record)) visit(child);
  };
  for (const story of documentStories(document))
    visit(storyBody(document, story).content);
  return out;
};
