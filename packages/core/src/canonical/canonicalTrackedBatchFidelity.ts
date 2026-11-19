import { Result } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  type DocumentOpRefusal,
  findStoryBody,
  OP_STORIES,
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
import type { BlockContent, Document, Paragraph } from "../types/document";
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

type StoryResolution =
  | { type: "resolved"; content: readonly BlockContent[] }
  | { type: "refused"; error: DocumentOpRefusal };

const resolveStory = ({
  version: { document, projection },
  story,
  ids,
  decision,
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
  return read(resolved.value.document);
};

type StoryAcceptanceOptions = {
  before: { version: StoryVersion; ids: readonly number[] };
  after: { version: StoryVersion; ids: readonly number[] };
  story: OpStory;
};

const acceptStoryPair = ({ before, after, story }: StoryAcceptanceOptions) => {
  const accept = ({ version, ids }: StoryAcceptanceOptions["before"]) =>
    resolveStory({ version, story, ids, decision: "accept" });
  const original = accept(before);
  const edited = accept(after);
  if (original?.type === "resolved" && edited?.type === "resolved")
    return { original: original.content, edited: edited.content };

  // An imported join can already be unresolvable. Keep that single record
  // pending in both versions and still verify all other revisions together.
  if (original?.type !== "refused" || edited?.type !== "refused")
    return undefined;
  const id = original.error.blockedParagraphMarkRevisionId;
  if (
    id === undefined ||
    edited.error.blockedParagraphMarkRevisionId !== id ||
    edited.error.message !== original.error.message ||
    !before.ids.includes(id) ||
    !after.ids.includes(id)
  )
    return undefined;
  const acceptRemaining = ({
    version,
    ids,
  }: StoryAcceptanceOptions["before"]) => {
    const group = getTrackedChangeGroupsForIds(version.projection, [id]).at(0);
    if (group?.length !== 1) return undefined;
    return accept({
      version,
      ids: ids.filter((revisionId) => revisionId !== id),
    });
  };
  const baseline = acceptRemaining(before);
  const proposed = acceptRemaining(after);
  return baseline?.type === "resolved" && proposed?.type === "resolved"
    ? { original: baseline.content, edited: proposed.content }
    : undefined;
};

type CanonicalTrackedBatchFidelityOptions = {
  story: OpStory;
  before: StoryVersion;
  after: StoryVersion;
  createdCommentIds: ReadonlySet<number>;
  retiredCommentIds?: readonly number[];
};

const plainParagraphKeys: ReadonlySet<string> = new Set([
  "type",
  "paraId",
  "textId",
  "formatting",
  "content",
  "listRendering",
  "renderedPageBreakBefore",
]);

const isPlainParagraph = (
  block: BlockContent | undefined,
): block is Paragraph =>
  block?.type === "paragraph" &&
  block.paraId !== undefined &&
  Object.keys(block).every((key) => plainParagraphKeys.has(key));

const isPlainTextRun = (item: Paragraph["content"][number]): boolean =>
  item.type === "run" &&
  item.propertyChanges === undefined &&
  item.preservedAttributes === undefined &&
  item.content.every((piece) => piece.type === "text");

const isStableNeighbor = (block: BlockContent | undefined): boolean =>
  isPlainParagraph(block) && block.content.every(isPlainTextRun);

type InlineSeam = Extract<
  Paragraph["content"][number],
  { type: "insertion" }
>["resolutionJoins"];

const plainSeamKeys: ReadonlySet<string> = new Set([
  "acceptance",
  "before",
  "after",
  "remove",
  "plainSides",
]);

// One level rejoins a split text run within this paragraph.
const isLocalPlainSeam = (seam: InlineSeam): boolean =>
  seam === undefined ||
  ([seam.before, seam.after, seam.remove].every(
    (depth) => depth === 0 || depth === 1,
  ) &&
    Object.keys(seam).every((key) => plainSeamKeys.has(key)));

const localRevisionIds = (paragraph: Paragraph): number[] | undefined => {
  const ids: number[] = [];
  for (const item of paragraph.content) {
    if (isPlainTextRun(item)) continue;
    if (
      (item.type !== "insertion" && item.type !== "deletion") ||
      item.info.suggestionId !== undefined ||
      !isLocalPlainSeam(item.resolutionJoins) ||
      !item.content.every(isPlainTextRun)
    ) {
      return undefined;
    }
    ids.push(item.info.id);
  }
  return ids;
};

type IsolatedTextChange = {
  before: Paragraph;
  after: Paragraph;
  beforeIds: number[];
  afterIds: number[];
};

const isolatedTextChanges = (
  before: BlockContent[],
  after: BlockContent[],
): IsolatedTextChange[] | undefined => {
  if (before.length !== after.length) return undefined;
  const changes: IsolatedTextChange[] = [];
  for (const [index, oldBlock] of before.entries()) {
    const newBlock = after.at(index);
    if (oldBlock === newBlock) continue;
    if (
      !isPlainParagraph(oldBlock) ||
      !isPlainParagraph(newBlock) ||
      oldBlock.paraId !== newBlock.paraId
    ) {
      return undefined;
    }
    if (
      index > 0 &&
      (before.at(index - 1) !== after.at(index - 1) ||
        !isStableNeighbor(before.at(index - 1)))
    ) {
      return undefined;
    }
    if (
      index + 1 < before.length &&
      (before.at(index + 1) !== after.at(index + 1) ||
        !isStableNeighbor(before.at(index + 1)))
    ) {
      return undefined;
    }
    const beforeIds = localRevisionIds(oldBlock);
    const afterIds = localRevisionIds(newBlock);
    if (beforeIds === undefined || afterIds === undefined) return undefined;
    changes.push({ before: oldBlock, after: newBlock, beforeIds, afterIds });
  }
  return changes;
};

const resolveIsolatedParagraph = (
  paragraph: Paragraph,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) => {
  const document: Document = {
    package: { document: { content: [paragraph] } },
  };
  if (revisionIds.length === 0) return document.package.document.content;
  const resolved = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);
  return resolved.isOk()
    ? resolved.value.document.package.document.content
    : undefined;
};

/** Undefined leaves verification with the complete native story resolver. */
const checkIsolatedTextChanges = (
  changes: readonly IsolatedTextChange[],
): string | null | undefined => {
  for (const { before, after, beforeIds, afterIds } of changes) {
    const original = resolveIsolatedParagraph(before, beforeIds, "reject");
    const restored = resolveIsolatedParagraph(after, afterIds, "reject");
    if (original === undefined || restored === undefined) return undefined;
    if (!sameOriginalBlocks(original, restored, new Set())) {
      return "Reject All would change the original content or formatting.";
    }
    if (
      resolveIsolatedParagraph(before, beforeIds, "accept") === undefined ||
      resolveIsolatedParagraph(after, afterIds, "accept") === undefined
    ) {
      return undefined;
    }
  }
  // Every changed run contains text only, so it cannot own protected content.
  return null;
};

const checkCanonicalTrackedBatchFidelity = ({
  story,
  before,
  after,
  createdCommentIds,
  retiredCommentIds,
}: CanonicalTrackedBatchFidelityOptions): string | null => {
  if (comparisonApplying()) return null;
  const beforeStory = findStoryBody(before.document, story);
  const afterStory = findStoryBody(after.document, story);
  if (beforeStory === undefined || afterStory === undefined)
    return "The tracked result could not be verified.";
  if (beforeStory.content === afterStory.content) return null;
  if (
    createdCommentIds.size === 0 &&
    beforeStory.finalSectionProperties?.propertyChanges === undefined &&
    afterStory.finalSectionProperties?.propertyChanges === undefined
  ) {
    const changes = isolatedTextChanges(
      beforeStory.content,
      afterStory.content,
    );
    if (changes !== undefined) {
      const verdict = checkIsolatedTextChanges(changes);
      if (verdict !== undefined) return verdict;
    }
  }

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
  if (rejectedBefore?.type !== "resolved" || rejectedAfter?.type !== "resolved")
    return "Reject All could not resolve the tracked result.";
  if (
    rejectedBefore.content !== rejectedAfter.content &&
    !sameOriginalBlocks(
      rejectedBefore.content,
      rejectedAfter.content,
      createdCommentIds,
    )
  ) {
    return "Reject All would change the original content or formatting.";
  }

  const accepted = acceptStoryPair({
    before: { version: before, ids: beforeIds.ordinary },
    after: { version: after, ids: afterIds.ordinary },
    story,
  });
  if (accepted === undefined)
    return "Accept All could not resolve the tracked result.";
  const { original: acceptedBefore, edited: acceptedAfter } = accepted;
  if (acceptedBefore === acceptedAfter) return null;
  const { removed, added } = changedWindow(acceptedBefore, acceptedAfter);
  let lost = missingProtectedBlockKind(removed, added, retiredCommentIds);
  // A note or comment can stay reachable through an unchanged block.
  if (lost === "note" || lost === "comment")
    lost = missingProtectedBlockKind(
      acceptedBefore,
      acceptedAfter,
      retiredCommentIds,
    );
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
