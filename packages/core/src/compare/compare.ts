/**
 * Deterministic `.docx` compare: two packages in, one redlined package plus a
 * JSON change list out.
 *
 * ## Determinism contract
 *
 * `compareDocx(base, target, options)` is a pure function of its three
 * arguments. It reads no clock and no randomness: revision dates come from
 * `options.timestamp`, and revision ids from a seed derived from the base
 * document's own highest existing revision id, so two runs over the same
 * inputs produce byte-identical buffers and deeply equal change lists.
 *
 * ## Round-trip contract
 *
 * Accepting every tracked change in the result yields the target's content;
 * rejecting every one yields the base's. Anything the comparison does not
 * cover is reported in `unsupported`, or fails the call, rather than being
 * silently dropped.
 *
 * @packageDocumentation
 */

import { Result } from "better-result";

import {
  FolioDocxReviewer,
  type FolioDocumentStoryHandle,
  type FolioRevisionStamp,
} from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { pairFolioDocumentStories } from "../document-stories";
import { planStoryCompare } from "./plan";
import {
  CompareDocxApplyError,
  CompareDocxOperationLimitError,
  CompareDocxParseError,
  CompareDocxSerializeError,
  type CompareChange,
  type CompareDocxError,
  type CompareDocxOptions,
  type CompareResult,
  type CompareUnsupportedPart,
} from "./types";

/**
 * Cap on operations one comparison generates. Both inputs are untrusted
 * documents, and each operation costs a document walk plus revision ids.
 */
export const MAX_COMPARE_OPERATIONS = 10_000;

const parseSide = async (
  buffer: ArrayBuffer,
  side: "base" | "target",
  author: string,
): Promise<Result<FolioDocxReviewer, CompareDocxParseError>> =>
  await Result.tryPromise({
    try: async () => await FolioDocxReviewer.fromBuffer(buffer, { author }),
    catch: (cause) =>
      new CompareDocxParseError({
        message: `The ${side} document could not be parsed.`,
        side,
        cause,
      }),
  });

/**
 * One past the highest revision id the base package already uses, across every
 * story. Seeding there keeps generated ids from colliding with revisions the
 * base already carries, while staying a pure function of the base bytes.
 */
const revisionIdSeedFor = (reviewer: FolioDocxReviewer): number => {
  let highest = 0;
  for (const { handle } of reviewer.listStories()) {
    const story = reviewer.readReviewedStory({ story: handle, view: "current-markup" });
    for (const change of story?.changes ?? []) {
      highest = Math.max(highest, change.id);
    }
  }
  return highest + 1;
};

const isMainStory = (story: FolioDocumentStoryHandle): boolean => story.type === "main";

/**
 * Compare `base` against `target` and return `base` carrying the tracked
 * changes that turn it into `target`, alongside the change list describing
 * them.
 */
export const compareDocx = async (
  base: ArrayBuffer,
  target: ArrayBuffer,
  options: CompareDocxOptions,
): Promise<Result<CompareResult, CompareDocxError>> => {
  const baseParse = await parseSide(base, "base", options.author);
  if (baseParse.isErr()) {
    return Result.err(baseParse.error);
  }
  const targetParse = await parseSide(target, "target", options.author);
  if (targetParse.isErr()) {
    return Result.err(targetParse.error);
  }

  const reviewer = baseParse.value;
  const targetReviewer = targetParse.value;
  const revisionStamp: FolioRevisionStamp = {
    date: options.timestamp,
    idSeed: revisionIdSeedFor(reviewer),
  };

  const baseStories = reviewer.listStories().map(({ handle }) => handle);
  const targetStories = targetReviewer.listStories().map(({ handle }) => handle);
  const changes: CompareChange[] = [];
  const unsupported: CompareUnsupportedPart[] = [];

  for (const { baseStory, revisedStory: targetStory } of pairFolioDocumentStories(
    baseStories,
    targetStories,
  )) {
    if (!baseStory) {
      unsupported.push({ reason: "story-missing-in-base", baseStory: null, targetStory });
      continue;
    }
    if (!targetStory) {
      unsupported.push({ reason: "story-missing-in-target", baseStory, targetStory: null });
      continue;
    }
    const baseSnapshot = isMainStory(baseStory) ? reviewer.snapshotStory(baseStory) : null;
    const targetSnapshot = isMainStory(targetStory)
      ? targetReviewer.snapshotStory(targetStory)
      : null;
    if (!baseSnapshot || !targetSnapshot) {
      unsupported.push({ reason: "secondary-story", baseStory, targetStory });
      continue;
    }

    const plan = planStoryCompare({
      story: baseStory,
      baseSnapshot,
      targetBlocks: targetSnapshot.blocks,
      maxOperations: MAX_COMPARE_OPERATIONS,
    });
    if (plan === null) {
      return Result.err(
        new CompareDocxOperationLimitError({
          message: "The comparison needs more operations than the engine generates.",
          limit: MAX_COMPARE_OPERATIONS,
        }),
      );
    }
    changes.push(...plan.changes);
    if (plan.operations.length === 0) {
      continue;
    }

    const { skipped } = reviewer.applyDocumentOperationsToStory({
      story: baseStory,
      snapshot: baseSnapshot,
      revisionStamp,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: plan.operations,
      },
    });
    if (skipped.length > 0) {
      return Result.err(
        new CompareDocxApplyError({
          message:
            "Some derived operations were refused, so the result would not match the target.",
          skipped,
        }),
      );
    }
  }

  const serialized = await Result.tryPromise({
    try: async () => await reviewer.toBuffer(),
    catch: (cause) =>
      new CompareDocxSerializeError({
        message: "The compared document could not be serialized.",
        cause,
      }),
  });
  if (serialized.isErr()) {
    return Result.err(serialized.error);
  }
  return Result.ok({ buffer: serialized.value, changes, unsupported });
};
