/**
 * The accepted or the rejected view of a document: every revision of every
 * story resolved one way, through the same resolution an Accept All or a
 * Reject All applies. Readers that show a document as if its changes were
 * accepted or rejected read it from here, so they and a reviewer resolving
 * the document agree on every paragraph, table and note it keeps.
 */

import type { Document } from "../model/document";
import { applyDocumentOps } from "./apply";
import { IDENTITY_SPACES, identityKeysIn } from "./ids";
import { documentStories, findStoryBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type OpStory,
  type RevisionDecision,
} from "./types";

export type ResolvedDocumentView = {
  document: Document;
  /**
   * The stories whose resolution refused, left as they were: a reader shows
   * them with their changes still in place and says so.
   */
  unresolved: readonly OpStory[];
};

const REVISION_KEY_PREFIX = `${IDENTITY_SPACES.REVISION}:`;

/** Every revision id a story's content carries, in document order. */
const storyRevisionIds = (document: Document, story: OpStory): number[] => {
  const body = findStoryBody(document, story);
  if (body === undefined) return [];
  return [
    ...new Set(
      identityKeysIn(body.content).flatMap((key) =>
        key.startsWith(REVISION_KEY_PREFIX)
          ? [Number(key.slice(REVISION_KEY_PREFIX.length))]
          : [],
      ),
    ),
  ];
};

/**
 * `document` with every revision resolved by `decision`, story by story. A
 * story the resolution of an earlier one removed (a note whose every
 * reference went) is not read again.
 */
export const resolvedDocumentView = (
  document: Document,
  decision: RevisionDecision,
): ResolvedDocumentView => {
  let current = document;
  const unresolved: OpStory[] = [];
  for (const story of documentStories(document)) {
    const revisionIds = storyRevisionIds(current, story);
    if (revisionIds.length === 0) continue;
    const resolved = applyDocumentOps(current, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds,
        decision,
      },
    ]);
    if (resolved.isErr()) unresolved.push(story);
    else current = resolved.value.document;
  }
  return { document: current, unresolved };
};
