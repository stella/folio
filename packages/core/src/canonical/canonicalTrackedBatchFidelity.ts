import { Result } from "better-result";

import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  findStoryBody,
  OP_STORIES,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import { comparisonApplying } from "../ai-edits/note-references";
import {
  missingProtectedBlockKind,
  newlyDeletedProtectedBlocks,
  sameOriginalBlocks,
} from "../ai-edits/tracked-batch-fidelity";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { itemBefore } from "../utils/listItems";
import type { createCanonicalBatchReview } from "./canonicalBatchReview";
import { changedWindow } from "./canonicalSeedContent";
import {
  type IsolatedTextChange,
  verifyWholeTextBatch,
} from "./canonicalWholeTextFidelity";

type CanonicalTrackedBatchFidelityOptions = {
  review: ReturnType<typeof createCanonicalBatchReview>;
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
      (itemBefore(before, index) !== itemBefore(after, index) ||
        !isStableNeighbor(itemBefore(before, index)))
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

const resolveIsolatedParagraphs = (
  paragraphs: Paragraph[],
  revisionIds: readonly number[],
  decision: RevisionDecision,
) => {
  const document: Document = {
    package: { document: { content: paragraphs } },
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
  if (
    verifyWholeTextBatch({ changes, resolve: resolveIsolatedParagraphs })
      .type === "verified"
  )
    return null;
  for (const { before, after, beforeIds, afterIds } of changes) {
    const original = resolveIsolatedParagraphs([before], beforeIds, "reject");
    const restored = resolveIsolatedParagraphs([after], afterIds, "reject");
    if (original === undefined || restored === undefined) return undefined;
    if (!sameOriginalBlocks(original, restored, new Set())) {
      return "Reject All would change the original content or formatting.";
    }
    if (
      resolveIsolatedParagraphs([before], beforeIds, "accept") === undefined ||
      resolveIsolatedParagraphs([after], afterIds, "accept") === undefined
    ) {
      return undefined;
    }
  }
  // Every changed run contains text only, so it cannot own protected content.
  return null;
};

const checkCanonicalTrackedBatchFidelity = ({
  createdCommentIds,
  retiredCommentIds,
  review,
}: CanonicalTrackedBatchFidelityOptions): string | null => {
  if (comparisonApplying()) return null;
  const { story, before, after } = review;
  const beforeStory = findStoryBody(before, story);
  const afterStory = findStoryBody(after, story);
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
  const rejected = review.reject();
  if (rejected === null)
    return "Reject All could not resolve the tracked result.";
  const { original: rejectedBefore, edited: rejectedAfter } = rejected;
  if (
    rejectedBefore !== rejectedAfter &&
    !sameOriginalBlocks(rejectedBefore, rejectedAfter, createdCommentIds)
  )
    return "Reject All would change the original content or formatting.";

  const accepted = review.accept();
  if (accepted === null)
    return "Accept All could not resolve the tracked result.";
  const { original: acceptedBefore, edited: acceptedAfter } = accepted;
  if (acceptedBefore === acceptedAfter) return null;
  const { removed, added } = changedWindow(acceptedBefore, acceptedAfter);
  // What the batch shows as deleted, Accept All may remove.
  const deleted = newlyDeletedProtectedBlocks(
    beforeStory.content,
    afterStory.content,
  );
  let lost = missingProtectedBlockKind(
    removed,
    added,
    retiredCommentIds,
    deleted,
  );
  // A note or comment can stay reachable through an unchanged block.
  if (lost === "note" || lost === "comment")
    lost = missingProtectedBlockKind(
      acceptedBefore,
      acceptedAfter,
      retiredCommentIds,
      deleted,
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
