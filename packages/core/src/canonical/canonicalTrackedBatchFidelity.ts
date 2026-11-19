import { Result } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  findStoryBody,
  type OpStory,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import { comparisonApplying } from "../ai-edits/note-references";
import {
  getTrackedChangeGroupsForIds,
  getTrackedChangeStatsFromDoc,
} from "../ai-edits/read";
import {
  missingProtectedBlockKind,
  sameOriginalBlocks,
} from "../ai-edits/tracked-batch-fidelity";
import {
  findSuggestionRange,
  getSuggestions,
} from "../prosemirror/commands/comments";
import { lockedControlSpans } from "../prosemirror/lockedControls";
import type { Document } from "../types/document";
import { lockViolationBetween } from "./canonicalLockGuard";
import { canonicalRevisionLocks } from "./canonicalRevisionLocks";
import { changedWindow } from "./canonicalSeedContent";

type StoryVersion = {
  document: Document;
  projection: PMNode;
};

type RevisionIds = {
  all: readonly number[];
  ordinary: readonly number[];
};

const revisionIdsOf = (doc: PMNode): RevisionIds => {
  const { ids } = getTrackedChangeStatsFromDoc(doc);
  if (ids.length === 0) return { all: ids, ordinary: ids };
  const state = EditorState.create({ doc });
  const suggested = new Set(
    getSuggestions(state).flatMap(
      ({ suggestionId }) =>
        findSuggestionRange(state, suggestionId)?.revisionIds ?? [],
    ),
  );
  return {
    all: ids,
    ordinary:
      suggested.size === 0 ? ids : ids.filter((id) => !suggested.has(id)),
  };
};

type ResolveStoryOptions = {
  version: StoryVersion;
  story: OpStory;
  ids: readonly number[];
  decision: RevisionDecision;
};

const resolveStory = ({
  version: { document, projection },
  story,
  ids,
  decision,
}: ResolveStoryOptions) => {
  if (ids.length === 0) return findStoryBody(document, story)?.content;
  let revisionIds = ids;
  if (lockedControlSpans(projection).length > 0) {
    const selected = new Set(ids);
    const groups = getTrackedChangeGroupsForIds(projection, ids).map((group) =>
      group.filter((id) => selected.has(id)),
    );
    revisionIds = canonicalRevisionLocks({
      document,
      story,
      groups,
      decision,
    }).unlocked.flat();
  }
  if (revisionIds.length === 0) return findStoryBody(document, story)?.content;
  const resolved = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story,
      revisionIds,
      decision,
    },
  ]);
  if (
    resolved.isErr() ||
    lockViolationBetween(document, resolved.value.document) !== null
  ) {
    return undefined;
  }
  return findStoryBody(resolved.value.document, story)?.content;
};

type CanonicalTrackedBatchFidelityOptions = {
  story: OpStory;
  before: StoryVersion;
  after: StoryVersion;
  createdCommentIds: ReadonlySet<number>;
};

const checkCanonicalTrackedBatchFidelity = ({
  story,
  before,
  after,
  createdCommentIds,
}: CanonicalTrackedBatchFidelityOptions): string | null => {
  if (comparisonApplying()) return null;
  const beforeStory = findStoryBody(before.document, story);
  const afterStory = findStoryBody(after.document, story);
  if (beforeStory === undefined || afterStory === undefined)
    return "The tracked result could not be verified.";
  if (beforeStory.content === afterStory.content) return null;

  // Resolve the full changed story so joins, range repairs and lock checks see
  // the same neighbors and package records as canonical review.
  const beforeIds = revisionIdsOf(before.projection);
  const afterIds = revisionIdsOf(after.projection);
  const rejectedBefore = resolveStory({
    version: before,
    story,
    ids: beforeIds.all,
    decision: "reject",
  });
  const rejectedAfter = resolveStory({
    version: after,
    story,
    ids: afterIds.all,
    decision: "reject",
  });
  if (rejectedBefore === undefined || rejectedAfter === undefined)
    return "Reject All could not resolve the tracked result.";
  if (
    rejectedBefore !== rejectedAfter &&
    !sameOriginalBlocks(rejectedBefore, rejectedAfter, createdCommentIds)
  ) {
    return "Reject All would change the original content or formatting.";
  }

  const acceptedBefore = resolveStory({
    version: before,
    story,
    ids: beforeIds.ordinary,
    decision: "accept",
  });
  const acceptedAfter = resolveStory({
    version: after,
    story,
    ids: afterIds.ordinary,
    decision: "accept",
  });
  if (acceptedBefore === undefined || acceptedAfter === undefined)
    return "Accept All could not resolve the tracked result.";
  if (acceptedBefore === acceptedAfter) return null;
  const { removed, added } = changedWindow(acceptedBefore, acceptedAfter);
  let lost = missingProtectedBlockKind(removed, added);
  // A note or comment can stay reachable through an unchanged block.
  if (lost === "note" || lost === "comment")
    lost = missingProtectedBlockKind(acceptedBefore, acceptedAfter);
  return lost === null ? null : `Accept All would lose a ${lost}.`;
};

/** Check canonical review through the package model before publishing. */
export const canonicalTrackedBatchFidelityError = (
  options: CanonicalTrackedBatchFidelityOptions,
): string | null => {
  const verdict = Result.try({
    try: () => checkCanonicalTrackedBatchFidelity(options),
    catch: () => "The tracked result could not be verified.",
  });
  return verdict.isOk() ? verdict.value : verdict.error;
};
