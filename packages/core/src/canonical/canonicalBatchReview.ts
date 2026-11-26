import type { Node as PMNode } from "prosemirror-model";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  type DocumentOpRefusal,
  findStoryBody,
  type OpStory,
  type RevisionDecision,
  suggestionOwnersInStory,
} from "@stll/docx-core/ops";

import { anchoredCommentIdsInBlocks } from "../ai-edits/comment-lifecycle";
import {
  getTrackedChangeGroupsForIds,
  getTrackedChangeStatsFromDoc,
} from "../ai-edits/read";
import { lockedControlSpans } from "../prosemirror/lockedControls";
import type { BlockContent, Document } from "../types/document";
import { lockViolationBetween } from "./canonicalLockGuard";
import { canonicalRevisionLocks } from "./canonicalRevisionLocks";
import { preservePropertySources } from "./propertySourcePreservation";

type StoryVersion = {
  document: Document;
  projection: PMNode;
};

type RevisionIds = {
  all: readonly number[];
  ordinary: readonly number[];
};

const revisionIdsOf = (version: StoryVersion, story: OpStory): RevisionIds => {
  const { ids } = getTrackedChangeStatsFromDoc(version.projection);
  if (ids.length === 0) return { all: ids, ordinary: ids };
  const suggested = suggestionOwnersInStory(version.document, story);
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
  /**
   * The document whose paragraphs a rejection gives back: their property
   * sources carry over by paragraph id, as an undo carries them, so a
   * paragraph the rejection rebuilt reads its properties as the original
   * states them rather than as list numbering projects them.
   */
  restores?: Document;
};

type StoryResolution =
  | { type: "resolved"; content: readonly BlockContent[] }
  | { type: "refused"; error: DocumentOpRefusal };

const resolveStory = ({
  version: { document, projection },
  story,
  ids,
  decision,
  restores,
}: ResolveStoryOptions): StoryResolution | undefined => {
  const read = (result: Document): StoryResolution | undefined => {
    const content = findStoryBody(result, story)?.content;
    return content === undefined ? undefined : { type: "resolved", content };
  };
  if (ids.length === 0) return read(document);
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
  if (revisionIds.length === 0) return read(document);
  const resolved = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story,
      revisionIds,
      decision,
    },
  ]);
  if (resolved.isErr()) return { type: "refused", error: resolved.error };
  if (lockViolationBetween(document, resolved.value.document) !== null)
    return undefined;
  return read(
    restores === undefined
      ? resolved.value.document
      : preservePropertySources({
          target: resolved.value.document,
          source: restores,
          stories: [story],
          indentation: "restore",
        }),
  );
};

type StoryPairOptions = {
  before: { version: StoryVersion; ids: readonly number[] };
  after: { version: StoryVersion; ids: readonly number[] };
  story: OpStory;
  decision: RevisionDecision;
  /** The edited side's paragraphs read their properties as this states them. */
  restores?: Document;
};

/**
 * Both sides of a batch resolved the same way. A revision the document held
 * before the batch that blocks its own resolution, alone and the same way on
 * both sides, stays pending on both: the proof is about what the batch
 * changes, and a blocker the batch did not make must not refuse every edit
 * of a document that holds one.
 */
const resolveStoryPair = ({
  before,
  after,
  story,
  decision,
  restores,
}: StoryPairOptions) => {
  let beforeIds = before.ids;
  let afterIds = after.ids;
  while (true) {
    const original = resolveStory({
      version: before.version,
      story,
      ids: beforeIds,
      decision,
    });
    const edited = resolveStory({
      version: after.version,
      story,
      ids: afterIds,
      decision,
      ...(restores === undefined ? {} : { restores }),
    });
    if (original?.type === "resolved" && edited?.type === "resolved")
      return { original: original.content, edited: edited.content };
    if (original?.type !== "refused" || edited?.type !== "refused")
      return undefined;

    // Keep identical, pre-existing standalone blockers pending on both sides.
    // Each retry removes an ID from both selections before resolving again.
    const id = original.error.blockedParagraphMarkRevisionId;
    if (
      id === undefined ||
      edited.error.blockedParagraphMarkRevisionId !== id ||
      edited.error.reason !== original.error.reason ||
      edited.error.message !== original.error.message ||
      !beforeIds.includes(id) ||
      !afterIds.includes(id) ||
      getTrackedChangeGroupsForIds(before.version.projection, [id]).at(0)
        ?.length !== 1 ||
      getTrackedChangeGroupsForIds(after.version.projection, [id]).at(0)
        ?.length !== 1
    )
      return undefined;
    beforeIds = beforeIds.filter((revisionId) => revisionId !== id);
    afterIds = afterIds.filter((revisionId) => revisionId !== id);
  }
};

type CanonicalBatchReviewOptions = {
  before: StoryVersion;
  after: StoryVersion;
  story: OpStory;
};

/** Share the same native acceptance proof with receipts and fidelity. */
export const createCanonicalBatchReview = ({
  before,
  after,
  story,
}: CanonicalBatchReviewOptions) => {
  let beforeIds: RevisionIds | undefined;
  let afterIds: RevisionIds | undefined;
  const revisions = (side: "before" | "after") =>
    side === "before"
      ? (beforeIds ??= revisionIdsOf(before, story))
      : (afterIds ??= revisionIdsOf(after, story));

  let accepted: ReturnType<typeof resolveStoryPair> | null;
  const accept = () => {
    if (accepted !== undefined) return accepted;
    accepted =
      resolveStoryPair({
        story,
        decision: "accept",
        before: { version: before, ids: revisions("before").ordinary },
        after: { version: after, ids: revisions("after").ordinary },
      }) ?? null;
    return accepted;
  };
  let rejected: ReturnType<typeof resolveStoryPair> | null;
  const reject = () => {
    if (rejected !== undefined) return rejected;
    rejected =
      resolveStoryPair({
        story,
        decision: "reject",
        before: { version: before, ids: revisions("before").all },
        after: { version: after, ids: revisions("after").all },
        restores: before.document,
      }) ?? null;
    return rejected;
  };
  return {
    story,
    before: before.document,
    after: after.document,
    accept,
    reject,
    acceptedCommentIds(side: "before" | "after") {
      const pair = accept();
      return pair === null
        ? null
        : anchoredCommentIdsInBlocks(
            side === "before" ? pair.original : pair.edited,
          );
    },
  };
};
