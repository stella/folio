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
 * ## Stages
 *
 * The call is four named steps: {@link parseComparison},
 * {@link planComparison}, {@link applyComparison}, {@link serializeComparison}.
 * `compareDocx` is their composition and nothing else, so the benchmark can
 * time the stages separately without keeping a second copy of the pipeline
 * that would drift from this one.
 *
 * @packageDocumentation
 */

import { Result } from "better-result";

import {
  FolioDocxReviewer,
  type FolioDocumentStoryHandle,
  type FolioRevisionStamp,
} from "../ai-edits/headless";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { pairFolioDocumentStories } from "../document-stories";
import { planStoryCompare, type CompareStoryPlan } from "./plan";
import { withFixedPackageDates } from "./reproducible-package";
import {
  CompareDocxApplyError,
  CompareDocxOperationLimitError,
  CompareDocxParseError,
  CompareDocxRoundTripError,
  CompareDocxSerializeError,
  InvalidCompareDocxOptionsError,
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
 * One story's text-and-structure projection: every block's text tagged with
 * the table cell it sits in. The tag is what makes the self-check below see a
 * paragraph that landed beside a table instead of inside it.
 */
const projectStory = (reviewer: FolioDocxReviewer, story: FolioDocumentStoryHandle): string[] => {
  const blocks = reviewer.readReviewedStory({ story, view: "final" })?.snapshot.blocks ?? [];
  return blocks.map(({ text, table }) => {
    const container = table
      ? `t${String(table.tableIndex)}r${String(table.rowIndex)}c${String(table.cellIndex)}p${String(table.paragraphIndex)}`
      : "body";
    return `${container}|${text}`;
  });
};

/** Two stories the comparison will align against one another. */
export type ComparedStoryPair = {
  baseStory: FolioDocumentStoryHandle;
  targetStory: FolioDocumentStoryHandle;
  baseSnapshot: FolioAIEditSnapshot;
  targetSnapshot: FolioAIEditSnapshot;
};

/** Everything the later stages need, and nothing they have to re-derive. */
export type ParsedComparison = {
  reviewer: FolioDocxReviewer;
  targetReviewer: FolioDocxReviewer;
  revisionStamp: FolioRevisionStamp;
  packageDate: Date;
  pairs: readonly ComparedStoryPair[];
  unsupported: readonly CompareUnsupportedPart[];
};

/** Stage 1: both packages to editor models, paired story by story. */
export const parseComparison = async (
  base: ArrayBuffer,
  target: ArrayBuffer,
  options: CompareDocxOptions,
): Promise<Result<ParsedComparison, CompareDocxParseError | InvalidCompareDocxOptionsError>> => {
  const packageDate = new Date(options.timestamp);
  if (Number.isNaN(packageDate.getTime())) {
    return Result.err(
      new InvalidCompareDocxOptionsError({
        message: "timestamp must be a date the package can be stamped with.",
        option: "timestamp",
        receivedValue: options.timestamp,
      }),
    );
  }

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
  const pairs: ComparedStoryPair[] = [];
  const unsupported: CompareUnsupportedPart[] = [];

  for (const { baseStory, revisedStory: targetStory } of pairFolioDocumentStories(
    reviewer.listStories().map(({ handle }) => handle),
    targetReviewer.listStories().map(({ handle }) => handle),
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
    pairs.push({ baseStory, targetStory, baseSnapshot, targetSnapshot });
  }

  return Result.ok({
    reviewer,
    targetReviewer,
    revisionStamp: { date: options.timestamp, idSeed: revisionIdSeedFor(reviewer) },
    packageDate,
    pairs,
    unsupported,
  });
};

/** One story's plan, kept with the pair it belongs to. */
export type PlannedStoryComparison = { pair: ComparedStoryPair; plan: CompareStoryPlan };

/**
 * Stage 2: align every paired story and derive its operations. Pure — no
 * parsing, no serialization, no clock.
 */
export const planComparison = ({
  pairs,
}: ParsedComparison): Result<readonly PlannedStoryComparison[], CompareDocxOperationLimitError> => {
  const planned: PlannedStoryComparison[] = [];
  for (const pair of pairs) {
    const plan = planStoryCompare({
      story: pair.baseStory,
      baseSnapshot: pair.baseSnapshot,
      targetBlocks: pair.targetSnapshot.blocks,
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
    planned.push({ pair, plan });
  }
  return Result.ok(planned);
};

/**
 * Stage 3: write the planned operations into the base document as tracked
 * changes, then check the work rather than trust it: accepting the story's
 * generated revisions must reproduce the target, structure included. A
 * difference the operation vocabulary cannot express would otherwise leave a
 * redline that reads plausibly and is wrong.
 */
export const applyComparison = (
  { reviewer, targetReviewer, revisionStamp }: ParsedComparison,
  planned: readonly PlannedStoryComparison[],
): Result<readonly CompareChange[], CompareDocxApplyError | CompareDocxRoundTripError> => {
  const changes: CompareChange[] = [];
  for (const { pair, plan } of planned) {
    changes.push(...plan.changes);
    if (plan.operations.length === 0) {
      continue;
    }

    const { skipped } = reviewer.applyDocumentOperationsToStory({
      story: pair.baseStory,
      snapshot: pair.baseSnapshot,
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

    const accepted = projectStory(reviewer, pair.baseStory);
    const expected = projectStory(targetReviewer, pair.targetStory);
    if (accepted.join(" ") !== expected.join(" ")) {
      return Result.err(
        new CompareDocxRoundTripError({
          message: "Accepting the generated tracked changes does not reproduce the target.",
          story: pair.baseStory,
          acceptedText: accepted,
          targetText: expected,
        }),
      );
    }
  }
  return Result.ok(changes);
};

/** Stage 4: the redlined package, with every ZIP entry date pinned. */
export const serializeComparison = async ({
  reviewer,
  packageDate,
}: ParsedComparison): Promise<Result<ArrayBuffer, CompareDocxSerializeError>> =>
  await Result.tryPromise({
    try: async () => await withFixedPackageDates(await reviewer.toBuffer(), packageDate),
    catch: (cause) =>
      new CompareDocxSerializeError({
        message: "The compared document could not be serialized.",
        cause,
      }),
  });

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
  const parsed = await parseComparison(base, target, options);
  if (parsed.isErr()) {
    return Result.err(parsed.error);
  }
  const planned = planComparison(parsed.value);
  if (planned.isErr()) {
    return Result.err(planned.error);
  }
  const changes = applyComparison(parsed.value, planned.value);
  if (changes.isErr()) {
    return Result.err(changes.error);
  }
  const serialized = await serializeComparison(parsed.value);
  if (serialized.isErr()) {
    return Result.err(serialized.error);
  }
  return Result.ok({
    buffer: serialized.value,
    changes: changes.value,
    unsupported: parsed.value.unsupported,
  });
};
