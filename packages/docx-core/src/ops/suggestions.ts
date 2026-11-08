import { Result } from "better-result";

import type { Document } from "../model/document";
import { applyDocumentOp } from "./apply";
import type { DocumentOpRefusal } from "./refusal";
import { documentStories, storyBody } from "./stories";
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
 * The document as a save writes it: every pending suggestion rejected. A
 * suggestion is a proposal the host holds, not a change in the file, so the
 * file reads as if it had never been made; ordinary tracked changes stay.
 */
export const withoutPendingSuggestions = (
  document: Document,
): Result<Document, DocumentOpRefusal> => {
  let current = document;
  for (const story of documentStories(document)) {
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
