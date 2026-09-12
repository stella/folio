import { panic, Result } from "better-result";

import type { FolioDocumentStoryHandle } from "../../ai-edits/headless";
import {
  type FolioContentComparison,
  type FolioContentComparisonError,
  type FolioContentComparisonSessionError,
  type FolioContentComparisonWorkSession,
} from "../../compare/content";
import {
  resolvedDocxContentSnapshot,
  resolvedDocxStoryHandle,
  type ResolvedDocxStorySnapshot,
} from "./resolved-docx-story-snapshot";

const RESOLVED_DOCX_STORY_PAIR_BRAND: unique symbol = Symbol("resolved-docx-story-pair");
const RESOLVED_DOCX_STORY_COMPARISON_BRAND: unique symbol = Symbol(
  "resolved-docx-story-comparison",
);

/** Two exact DOCX story projections bound before semantic comparison begins. */
export type ResolvedDocxStoryPair = {
  readonly [RESOLVED_DOCX_STORY_PAIR_BRAND]: true;
};

/** The canonical neutral comparison owned by one exact DOCX story pair. */
export type ResolvedDocxStoryComparison = {
  readonly [RESOLVED_DOCX_STORY_COMPARISON_BRAND]: true;
};

export type ResolvedDocxStoryPairPayload = {
  readonly baseStory: FolioDocumentStoryHandle;
  readonly targetStory: FolioDocumentStoryHandle;
  readonly baseSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
};

export type ResolvedDocxStoryComparisonPayload = ResolvedDocxStoryPairPayload & {
  readonly pair: ResolvedDocxStoryPair;
  readonly comparison: FolioContentComparison;
};

const payloadByPair = new WeakMap<ResolvedDocxStoryPair, ResolvedDocxStoryPairPayload>();
const payloadByComparison = new WeakMap<
  ResolvedDocxStoryComparison,
  ResolvedDocxStoryComparisonPayload
>();

/** Bind two genuine snapshots before they can enter comparison or transport lowering. */
export const createResolvedDocxStoryPair = ({
  baseSnapshot,
  targetSnapshot,
}: {
  readonly baseSnapshot: ResolvedDocxStorySnapshot;
  readonly targetSnapshot: ResolvedDocxStorySnapshot;
}): ResolvedDocxStoryPair => {
  const baseStory = resolvedDocxStoryHandle(baseSnapshot);
  const targetStory = resolvedDocxStoryHandle(targetSnapshot);
  if (baseStory.type !== targetStory.type) {
    return panic("A resolved DOCX story pair must contain compatible story kinds", {
      baseKind: baseStory.type,
      targetKind: targetStory.type,
    });
  }
  const payload = Object.freeze({
    baseStory,
    targetStory,
    baseSnapshot,
    targetSnapshot,
  });
  const pair = Object.freeze({ [RESOLVED_DOCX_STORY_PAIR_BRAND]: true as const });
  payloadByPair.set(pair, payload);
  return pair;
};

/** Resolve only a pair issued by the factory in this module. */
export const resolvedDocxStoryPairPayload = (
  pair: ResolvedDocxStoryPair,
): ResolvedDocxStoryPairPayload =>
  payloadByPair.get(pair) ?? panic("A resolved DOCX story pair was not created by Folio");

/**
 * Run the neutral engine for one exact pair. The returned capsule is the only
 * value the DOCX planner accepts, so semantic output cannot be joined later to
 * structurally similar or independently supplied snapshots.
 */
export const compareResolvedDocxStoryPair = ({
  pair,
  workSession,
}: {
  readonly pair: ResolvedDocxStoryPair;
  readonly workSession: FolioContentComparisonWorkSession;
}): Result<
  ResolvedDocxStoryComparison,
  FolioContentComparisonError | FolioContentComparisonSessionError
> => {
  const pairPayload = resolvedDocxStoryPairPayload(pair);
  const captured = workSession.captureComparison({
    base: resolvedDocxContentSnapshot(pairPayload.baseSnapshot),
    revised: resolvedDocxContentSnapshot(pairPayload.targetSnapshot),
  });
  if (captured.isErr()) return Result.err(captured.error);
  const compared = captured.value.compare();
  if (compared.isErr()) return Result.err(compared.error);

  const comparison = Object.freeze({
    [RESOLVED_DOCX_STORY_COMPARISON_BRAND]: true as const,
  });
  payloadByComparison.set(
    comparison,
    Object.freeze({ ...pairPayload, pair, comparison: compared.value }),
  );
  return Result.ok(comparison);
};

/** Resolve only semantic output issued for a genuine exact story pair. */
export const resolvedDocxStoryComparisonPayload = (
  comparison: ResolvedDocxStoryComparison,
): ResolvedDocxStoryComparisonPayload =>
  payloadByComparison.get(comparison) ??
  panic("A resolved DOCX story comparison was not created by Folio");
