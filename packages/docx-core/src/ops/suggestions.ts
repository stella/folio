import { Result } from "better-result";

import type {
  Document,
  DocumentBody,
  TrackedChangeInfo,
} from "../model/document";
import { applyDocumentOp } from "./apply";
import type { DocumentOpRefusal } from "./refusal";
import { documentStories, findStoryBody, storyBody } from "./stories";
import type { OpStory, ParagraphAbsorption } from "./types";
import { DOCUMENT_OP_TYPES, REVISION_DECISIONS } from "./types";

type SuggestionRevision = TrackedChangeInfo & { suggestionId: string };

const visitSuggestionRevisions = (
  value: unknown,
  visit: (info: SuggestionRevision) => void,
): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) visitSuggestionRevisions(item, visit);
    return;
  }
  const record = value as Record<string, unknown>;
  const info = record["info"] as Partial<SuggestionRevision> | undefined;
  if (
    typeof info === "object" &&
    info !== null &&
    typeof info.suggestionId === "string" &&
    typeof info.id === "number"
  )
    visit(info as SuggestionRevision);
  for (const child of Object.values(record))
    visitSuggestionRevisions(child, visit);
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

type PendingSuggestionProjectionOptions = {
  /** Host receipts also identify comments belonging to a pending proposal. */
  commentIds?: ReadonlySet<number>;
  onAbsorptions?: (
    story: OpStory,
    absorptions: readonly ParagraphAbsorption[],
  ) => void;
};

/**
 * The document as a save writes it: every pending suggestion rejected. A
 * suggestion is a proposal the host holds, not a change in the file, so the
 * file reads as if it had never been made; ordinary tracked changes stay. A
 * comment a suggestion proposes is left out with it.
 */
export const withoutPendingSuggestions = (
  document: Document,
  { commentIds, onAbsorptions }: PendingSuggestionProjectionOptions = {},
): Result<Document, DocumentOpRefusal> => {
  const proposedComments = new Set(suggestedCommentIds(document));
  if (commentIds !== undefined) {
    for (const { id } of document.package.document.comments ?? []) {
      if (commentIds.has(id)) proposedComments.add(id);
    }
  }
  const uncommented = withoutSuggestedComments(document, [...proposedComments]);
  if (uncommented.isErr()) return uncommented;
  let current = uncommented.value;
  const receipts: {
    story: OpStory;
    absorbed: readonly ParagraphAbsorption[];
  }[] = [];
  for (const story of documentStories(current)) {
    const ids = new Set<number>();
    visitSuggestionRevisions(storyBody(current, story), ({ id }) =>
      ids.add(id),
    );
    if (ids.size === 0) continue;
    const rejected = applyDocumentOp(current, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story,
      revisionIds: [...ids],
      decision: REVISION_DECISIONS.REJECT,
    });
    if (rejected.isErr()) return Result.err(rejected.error);
    const absorbed = rejected.value.touched.absorbed;
    if (
      onAbsorptions !== undefined &&
      absorbed !== undefined &&
      absorbed.length > 0
    )
      receipts.push({ story, absorbed });
    current = rejected.value.document;
  }
  for (const { story, absorbed } of receipts) onAbsorptions?.(story, absorbed);
  return Result.ok(current);
};

const visitStorySuggestionRevisions = (
  body: DocumentBody,
  visit: (info: TrackedChangeInfo, suggestionId: string) => void,
): void => {
  visitSuggestionRevisions(body.content, (info) =>
    visit(info, info.suggestionId),
  );
  for (const { info } of body.finalSectionProperties?.propertyChanges ?? [])
    if (info.suggestionId !== undefined) visit(info, info.suggestionId);
};

/** Revisions of one pending suggestion, one info per ID in document order. */
export const suggestionRevisionInfos = (
  document: Document,
  story: OpStory,
  suggestionId: string,
): TrackedChangeInfo[] => {
  const out = new Map<number, TrackedChangeInfo>();
  visitStorySuggestionRevisions(storyBody(document, story), (info, owner) => {
    if (owner === suggestionId && !out.has(info.id)) out.set(info.id, info);
  });
  return [...out.values()];
};

/** The first proposal owning each physical revision ID in this story. */
export const suggestionOwnersInStory = (
  document: Document,
  story: OpStory,
): ReadonlyMap<number, string> => {
  const owners = new Map<number, string>();
  const body = findStoryBody(document, story);
  if (body === undefined) return owners;
  visitStorySuggestionRevisions(body, ({ id }, suggestionId) => {
    if (!owners.has(id)) owners.set(id, suggestionId);
  });
  return owners;
};

/**
 * The ids of every pending suggestion in a document, in every story, and of
 * every suggestion that proposes a comment.
 */
export const pendingSuggestionIdsIn = (document: Document): Set<string> => {
  const out = new Set<string>();
  for (const comment of document.package.document.comments ?? [])
    if (comment.suggestionId !== undefined) out.add(comment.suggestionId);
  for (const story of documentStories(document)) {
    visitStorySuggestionRevisions(
      storyBody(document, story),
      (_info, suggestionId) => out.add(suggestionId),
    );
  }
  return out;
};
