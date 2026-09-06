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

import { panic, Result } from "better-result";

import {
  FolioDocxReviewer,
  type FolioDocumentStoryHandle,
  type FolioNumberingLevel,
  type FolioRevisionStamp,
} from "../ai-edits/headless";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import type { WordDiffGranularity } from "../ai-edits/word-diff";
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

type ExistingRevisions = {
  /**
   * One past the highest revision id the base package already uses, across
   * every story. Seeding there keeps generated ids from colliding with
   * revisions the base already carries, while staying a pure function of the
   * base bytes.
   */
  idSeed: number;
  /**
   * Whether the base arrived carrying unresolved revisions. When it did, the
   * compared base is its accepted view rather than the package as stored, and
   * the result must be serialized even if nothing else changed.
   */
  present: boolean;
};

/** Read before either side is resolved, so it describes the package as it arrived. */
const existingRevisionsOf = (reviewer: FolioDocxReviewer): ExistingRevisions => {
  let highest = 0;
  let present = false;
  for (const { handle } of reviewer.listStories()) {
    const story = reviewer.readReviewedStory({ story: handle, view: "current-markup" });
    for (const change of story?.changes ?? []) {
      highest = Math.max(highest, change.id);
      present = true;
    }
  }
  return { idSeed: highest + 1, present };
};

/**
 * One story's text-and-structure projection: every block's text tagged with
 * the table cell it sits in. The tag is what makes the self-check below see a
 * paragraph that landed beside a table instead of inside it.
 */
const projectStory = (reviewer: FolioDocxReviewer, story: FolioDocumentStoryHandle): string[] => {
  const blocks = reviewer.readReviewedStory({ story, view: "final" })?.snapshot.blocks ?? [];
  return blocks.map(({ text, table, styleId, listLevel }) => {
    const container = table
      ? `t${String(table.tableIndex)}r${String(table.rowIndex)}c${String(table.cellIndex)}p${String(table.paragraphIndex)}`
      : "body";
    // The properties the comparison claims to compare are in the projection
    // too, or the self-check would pass a redline that reproduces every word
    // and leaves a list item at the wrong level.
    return `${container}|${styleId ?? ""}|${listLevel ?? ""}|${text}`;
  });
};

const numberingKey = ({ numId, level }: FolioNumberingLevel): string =>
  `${String(numId)}:${String(level)}`;

const sameNumbering = (left: FolioNumberingLevel, right: FolioNumberingLevel): boolean =>
  left.format === right.format && left.levelText === right.levelText && left.start === right.start;

/**
 * Numbering levels that differ between the two packages, in `numId` then
 * level order.
 *
 * A list renumbered BY an edit needs no entry: labels come from these
 * definitions rather than from the paragraphs, so inserting an item already
 * renumbers the ones below it as-if-accepted. A definition that itself
 * changed — decimal to lower-roman, a different level template, a different
 * start — changes every label in the list and nothing in any block's text,
 * which is exactly the difference a caller would otherwise never hear about.
 */
type NumberingChange = Extract<CompareChange, { kind: "numbering" }>;

const compareNumbering = (
  base: FolioDocxReviewer,
  target: FolioDocxReviewer,
): NumberingChange[] => {
  const baseLevels = new Map(
    base.readNumberingDefinitions().map((level) => [numberingKey(level), level]),
  );
  const targetLevels = new Map(
    target.readNumberingDefinitions().map((level) => [numberingKey(level), level]),
  );
  const changes: NumberingChange[] = [];
  for (const [key, before] of baseLevels) {
    const after = targetLevels.get(key) ?? null;
    if (after === null || !sameNumbering(before, after)) {
      changes.push({ kind: "numbering", numId: before.numId, level: before.level, before, after });
    }
  }
  for (const [key, after] of targetLevels) {
    if (!baseLevels.has(key)) {
      changes.push({
        kind: "numbering",
        numId: after.numId,
        level: after.level,
        before: null,
        after,
      });
    }
  }
  return changes.toSorted((left, right) => left.numId - right.numId || left.level - right.level);
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
  /** Token size a changed paragraph's redline is cut at. */
  granularity: WordDiffGranularity;
  /**
   * The base package as it arrived. It is the result when nothing changed, but
   * only when it carried no revisions of its own: otherwise the compared base
   * is its accepted view and these bytes are a different document.
   */
  baseBuffer: ArrayBuffer;
  baseCarriedRevisions: boolean;
  reviewer: FolioDocxReviewer;
  targetReviewer: FolioDocxReviewer;
  revisionStamp: FolioRevisionStamp;
  packageDate: Date;
  pairs: readonly ComparedStoryPair[];
  /** Package-level numbering differences, which belong to no story. */
  numberingChanges: readonly CompareChange[];
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
  const existing = existingRevisionsOf(reviewer);
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
    // Compare the accepted view of both sides. An input that already carries
    // revisions otherwise makes the result unreadable: the redline would layer
    // this comparison's marks on top of someone else's, and rejecting them all
    // would land on neither document. Accepting first states one question --
    // how does the base as it stands differ from the target as it stands --
    // and leaves the answer as the only redline in the package.
    reviewer.resolveReviewedStory({ story: baseStory, view: "final" });
    targetReviewer.resolveReviewedStory({ story: targetStory, view: "final" });

    const baseSnapshot = reviewer.snapshotStory(baseStory);
    const targetSnapshot = targetReviewer.snapshotStory(targetStory);
    if (!baseSnapshot || !targetSnapshot) {
      unsupported.push({ reason: "story-not-editable", baseStory, targetStory });
      continue;
    }
    pairs.push({ baseStory, targetStory, baseSnapshot, targetSnapshot });
  }

  return Result.ok({
    granularity: options.granularity ?? "word",
    baseBuffer: base,
    baseCarriedRevisions: existing.present,
    reviewer,
    targetReviewer,
    revisionStamp: { date: options.timestamp, idSeed: existing.idSeed },
    packageDate,
    pairs,
    numberingChanges: compareNumbering(reviewer, targetReviewer),
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
  { reviewer, targetReviewer, revisionStamp, granularity, numberingChanges }: ParsedComparison,
  planned: readonly PlannedStoryComparison[],
): Result<readonly CompareChange[], CompareDocxApplyError | CompareDocxRoundTripError> => {
  const changes: CompareChange[] = [...numberingChanges];
  // Each story gets the range that starts where the previous story's ended.
  // A revision `w:id` is scoped to the package, not the part, so two stories
  // seeded alike would let a reader resolving a header revision resolve a
  // body revision with it.
  let idSeed = revisionStamp.idSeed;
  for (const { pair, plan } of planned) {
    changes.push(...plan.changes);
    if (plan.operations.length === 0) {
      continue;
    }

    const { skipped, nextRevisionId } = reviewer.applyDocumentOperationsToStory({
      story: pair.baseStory,
      snapshot: pair.baseSnapshot,
      revisionStamp: { date: revisionStamp.date, idSeed },
      wordDiff: { granularity },
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: plan.operations,
      },
    });
    if (nextRevisionId === undefined) {
      // Only a host bridge that does not allocate ids itself omits this, and
      // the comparison drives the in-process applier.
      panic("The applier did not report where it left the revision-id counter", {
        story: pair.baseStory,
      });
    }
    idSeed = nextRevisionId;
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

/**
 * Stage 4: the result package, with every ZIP entry date pinned.
 *
 * A comparison that found nothing returns the base bytes as they arrived. A
 * change is only ever reported alongside the operations that realize it, so no
 * operations means no changes, and re-serializing then rewrites a document
 * nobody edited: on a 2,200-block pair that was a second of work to reproduce
 * the input.
 *
 * Unless the base carried revisions of its own. Then the compared base was its
 * accepted view, the arriving bytes are a different document, and handing them
 * back would make rejecting the result land before the previous reviewer's
 * edits rather than after them.
 *
 * The short-circuit lives here rather than in {@link compareDocx} so that
 * every caller of the stages sees the same decision. Putting it in the
 * composition let the benchmark's own composition disagree with the shipped
 * one within a single run.
 */
export const serializeComparison = async (
  { baseBuffer, baseCarriedRevisions, reviewer, packageDate }: ParsedComparison,
  planned: readonly PlannedStoryComparison[],
): Promise<Result<ArrayBuffer, CompareDocxSerializeError>> => {
  if (!baseCarriedRevisions && planned.every(({ plan }) => plan.operations.length === 0)) {
    return Result.ok(baseBuffer);
  }
  return await Result.tryPromise({
    try: async () => await withFixedPackageDates(await reviewer.toBuffer(), packageDate),
    catch: (cause) =>
      new CompareDocxSerializeError({
        message: "The compared document could not be serialized.",
        cause,
      }),
  });
};

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
  const serialized = await serializeComparison(parsed.value, planned.value);
  if (serialized.isErr()) {
    return Result.err(serialized.error);
  }
  return Result.ok({
    buffer: serialized.value,
    changes: changes.value,
    unsupported: parsed.value.unsupported,
  });
};
