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
 * rejecting every one yields the base's. Both directions are checked before
 * the call returns, and the verdict travels with the result as
 * `verification`. Anything the comparison does not cover is reported in
 * `unsupported`, or fails the call, rather than being silently dropped.
 *
 * An unproven redline is refused by default. `mode: "bestEffort"` returns it
 * anyway, with every invariant that did not hold named: a caller that would
 * rather show its best attempt and say what is missing can, and one that wants
 * a redline it can stand behind still gets nothing else.
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
  getFolioDocxComparisonAccess,
  type FolioDocumentStoryHandle,
  type FolioNumberingLevel,
  type FolioRevisionStamp,
} from "../ai-edits/headless";
import { projectTableGeometry } from "../internal/compare/table-geometry-program";
import { storyTablesOf } from "../ai-edits/snapshot";
import {
  resolvedDocxContentBlocks,
  resolvedDocxContentSnapshot,
  resolvedDocxNumberingReferenceKeys,
  resolvedDocxOperationSnapshot,
  type ResolvedDocxStorySnapshot,
} from "../internal/compare/resolved-docx-story-snapshot";
import type { WordDiffGranularity } from "./text-diff";
import { pairFolioDocumentStories } from "../document-stories";
import {
  compareContentStories,
  createContentComparisonWorkSession,
  type FolioContentComparison,
  FolioContentComparisonSessionError,
} from "./content";
import { planStoryCompare, type CompareStoryPlan } from "./plan";
import { withFixedPackageDates } from "./reproducible-package";
import {
  CompareDocxApplyError,
  CompareDocxContentComparisonError,
  CompareDocxFinalParagraphMarkError,
  CompareDocxLoweringError,
  CompareDocxOperationLimitError,
  CompareDocxParseError,
  CompareDocxRoundTripError,
  CompareDocxSerializeError,
  CompareDocxUnsupportedError,
  InvalidCompareDocxOptionsError,
  type CompareChange,
  type CompareDocxError,
  type CompareDocxOptions,
  type CompareResult,
  type CompareUnsupportedPart,
} from "./types";
import {
  classifyContentProjectionMismatch,
  classifyGeometryMismatch,
  revisedFinalParagraphMarks,
  type CompareVerification,
  type CompareVerificationFailure,
} from "./verification";

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

const storyKey = (story: FolioDocumentStoryHandle): string => {
  switch (story.type) {
    case "main":
      return "main";
    case "header":
    case "footer":
      return `${story.type}:${story.relationshipId}`;
    case "footnote":
    case "endnote":
      return `${story.type}:${String(story.noteId)}`;
    default: {
      const exhaustive: never = story;
      return exhaustive;
    }
  }
};

const numberingKey = ({ numId, level }: Pick<FolioNumberingLevel, "numId" | "level">): string =>
  `${String(numId)}:${String(level)}`;

const sameNumbering = (left: FolioNumberingLevel, right: FolioNumberingLevel): boolean =>
  left.format === right.format && left.levelText === right.levelText && left.start === right.start;

/**
 * Referenced numbering levels that differ between the two packages, in
 * `numId` then level order.
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
  referenced: ReadonlySet<string>,
): NumberingChange[] => {
  const baseLevels = new Map(
    base.readNumberingDefinitions().map((level) => [numberingKey(level), level]),
  );
  const targetLevels = new Map(
    target.readNumberingDefinitions().map((level) => [numberingKey(level), level]),
  );
  const changes: NumberingChange[] = [];
  for (const [key, before] of baseLevels) {
    if (!referenced.has(key)) {
      continue;
    }
    const after = targetLevels.get(key) ?? null;
    if (after === null || !sameNumbering(before, after)) {
      changes.push({ kind: "numbering", numId: before.numId, level: before.level, before, after });
    }
  }
  for (const [key, after] of targetLevels) {
    if (referenced.has(key) && !baseLevels.has(key)) {
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
  baseSnapshot: ResolvedDocxStorySnapshot;
  targetSnapshot: ResolvedDocxStorySnapshot;
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
  // Compare the accepted view of both sides. An input that already carries
  // revisions otherwise makes the result unreadable: the redline would layer
  // this comparison's marks on top of someone else's, and rejecting them all
  // would land on neither document. Accepting first states one question --
  // how does the base as it stands differ from the target as it stands --
  // and leaves the answer as the only redline in the package.
  //
  // EVERY story, not only the paired ones. A note or a header the other side
  // does not have is reported rather than compared, but the base's copy of it
  // still ships in the result, so it ships accepted like the rest -- and an
  // unresolvable mark it carried, on the paragraph a note ends with, would
  // otherwise fail the structural guard on bytes this comparison never wrote.
  const baseProjection =
    getFolioDocxComparisonAccess(reviewer).projectStories("with-revision-census");
  const targetProjection =
    getFolioDocxComparisonAccess(targetReviewer).projectStories("without-revision-census");
  const baseStories: FolioDocumentStoryHandle[] = [];
  const targetStories: FolioDocumentStoryHandle[] = [];
  const baseSnapshots = new Map<FolioDocumentStoryHandle, ResolvedDocxStorySnapshot | null>();
  const targetSnapshots = new Map<FolioDocumentStoryHandle, ResolvedDocxStorySnapshot | null>();
  for (const { handle, snapshot } of baseProjection.stories) {
    baseStories.push(handle);
    baseSnapshots.set(handle, snapshot);
  }
  for (const { handle, snapshot } of targetProjection.stories) {
    targetStories.push(handle);
    targetSnapshots.set(handle, snapshot);
  }
  const pairs: ComparedStoryPair[] = [];
  const unsupported: CompareUnsupportedPart[] = [];
  const referencedNumberingLevels = new Set<string>();
  const collectNumberingReferences = (
    snapshot: ResolvedDocxStorySnapshot | null | undefined,
  ): void => {
    if (!snapshot) {
      return;
    }
    for (const referenceKey of resolvedDocxNumberingReferenceKeys(snapshot)) {
      referencedNumberingLevels.add(referenceKey);
    }
  };

  for (const { baseStory, revisedStory: targetStory } of pairFolioDocumentStories(
    baseStories,
    targetStories,
  )) {
    if (!baseStory) {
      if (!targetStory) {
        panic("A story pair contained neither a base nor a target story");
      }
      collectNumberingReferences(targetSnapshots.get(targetStory));
      unsupported.push({ reason: "story-missing-in-base", baseStory: null, targetStory });
      continue;
    }
    if (!targetStory) {
      collectNumberingReferences(baseSnapshots.get(baseStory));
      unsupported.push({ reason: "story-missing-in-target", baseStory, targetStory: null });
      continue;
    }
    const baseSnapshot = baseSnapshots.get(baseStory);
    const targetSnapshot = targetSnapshots.get(targetStory);
    collectNumberingReferences(baseSnapshot);
    collectNumberingReferences(targetSnapshot);
    if (!baseSnapshot || !targetSnapshot) {
      unsupported.push({ reason: "story-not-editable", baseStory, targetStory });
      continue;
    }
    pairs.push({ baseStory, targetStory, baseSnapshot, targetSnapshot });
  }

  return Result.ok({
    granularity: options.granularity ?? "word",
    baseBuffer: base,
    baseCarriedRevisions: baseProjection.revisions.present,
    reviewer,
    revisionStamp: {
      date: options.timestamp,
      idSeed: baseProjection.revisions.highestId + 1,
    },
    packageDate,
    pairs,
    numberingChanges: compareNumbering(reviewer, targetReviewer, referencedNumberingLevels),
    unsupported,
  });
};

/** One story's canonical comparison and its transport lowering. */
export type PlannedStoryComparison = {
  pair: ComparedStoryPair;
  comparison: FolioContentComparison;
  plan: CompareStoryPlan;
};

/**
 * Stage 2: align every paired story and derive its operations. Pure — no
 * parsing, no serialization, no clock.
 */
export const planComparison = ({
  pairs,
  granularity,
}: ParsedComparison): Result<
  readonly PlannedStoryComparison[],
  CompareDocxContentComparisonError | CompareDocxLoweringError | CompareDocxOperationLimitError
> => {
  const planned: PlannedStoryComparison[] = [];
  const workSession = createContentComparisonWorkSession({ granularity });
  const compared = compareContentStories({
    workSession,
    stories: pairs.map((pair) => ({
      key: pair,
      base: resolvedDocxContentSnapshot(pair.baseSnapshot),
      revised: resolvedDocxContentSnapshot(pair.targetSnapshot),
    })),
  });
  if (compared.isErr()) {
    if (compared.error.cause instanceof FolioContentComparisonSessionError) {
      return panic("The DOCX comparison misused its private content work session", {
        cause: compared.error.cause,
      });
    }
    const pair = pairs[compared.error.storyIndex];
    if (!pair) {
      return panic("A content comparison failure named a missing story pair", {
        storyIndex: compared.error.storyIndex,
      });
    }
    return Result.err(
      new CompareDocxContentComparisonError({
        message: "A DOCX story did not satisfy the bounded content comparison contract.",
        story: pair.baseStory,
        cause: compared.error.cause,
      }),
    );
  }
  let remainingOperations = MAX_COMPARE_OPERATIONS;
  for (const { key: pair, comparison } of compared.value) {
    const result = planStoryCompare({
      story: pair.baseStory,
      baseSnapshot: pair.baseSnapshot,
      targetSnapshot: pair.targetSnapshot,
      comparison,
      maxOperations: remainingOperations,
    });
    if (result.isErr()) return Result.err(result.error);
    const plan = result.value;
    planned.push({ pair, comparison, plan });
    remainingOperations -= plan.program.size;
  }
  return Result.ok(planned);
};

/** What stage 3 produced: the change list, whether it was proven, and whether it wrote anything. */
export type AppliedComparison = {
  changes: readonly CompareChange[];
  verification: CompareVerification;
  unsupported: readonly CompareUnsupportedPart[];
  /**
   * Whether the stage wrote into the base document at all. A comparison that
   * found nothing writes nothing, and the serialize stage then hands the
   * arriving bytes back rather than reproducing them.
   */
  documentChanged: boolean;
};

/**
 * Stage 3: write the planned operations into the base document as tracked
 * changes, then check the work rather than trust it. Both directions of the
 * round trip are checked, structure and table geometry included: accepting the
 * story's generated revisions must reproduce the target, and rejecting them
 * must reproduce the base it was compared from. A difference the operation
 * vocabulary cannot express would otherwise leave a redline that reads
 * plausibly and is wrong.
 *
 * The check reports rather than throws. {@link compareDocx} decides what to do
 * with an unverified result, because "give me your best attempt and tell me
 * what you could not represent" and "give me nothing unless you can prove it"
 * are both legitimate asks and only the caller knows which one it is making.
 */
export const applyComparison = (
  { reviewer, revisionStamp, numberingChanges }: ParsedComparison,
  planned: readonly PlannedStoryComparison[],
  {
    mode,
    unsupported,
  }: {
    mode: "strict" | "bestEffort";
    unsupported: readonly CompareUnsupportedPart[];
  },
): Result<AppliedComparison, CompareDocxApplyError | CompareDocxUnsupportedError> => {
  const comparisonAccess = getFolioDocxComparisonAccess(reviewer);
  const changes: CompareChange[] = [...numberingChanges];
  const failures: CompareVerificationFailure[] = [];
  const preparedStories = planned.map(({ pair, comparison, plan }) => ({
    pair,
    comparison,
    plan,
    prepared: comparisonAccess.prepareStoryProgram({
      story: pair.baseStory,
      snapshot: pair.baseSnapshot,
      target: pair.targetSnapshot,
      program: plan.program,
    }),
  }));
  const transportUnsupported: CompareUnsupportedPart[] = preparedStories.flatMap(
    ({ pair, prepared }) =>
      prepared.issues.map(({ instructionIndex, reason, blockId }) => ({
        reason: "transport-preflight" as const,
        story: pair.baseStory,
        instructionIndex,
        detail: reason,
        ...(blockId !== undefined && { blockId }),
      })),
  );
  const allUnsupported = Object.freeze([...unsupported, ...transportUnsupported]);
  if (mode === "strict" && allUnsupported.length > 0) {
    return Result.err(
      new CompareDocxUnsupportedError({
        message: "The comparison contains differences with no proved tracked-document lowering.",
        unsupported: allUnsupported,
      }),
    );
  }
  for (const omitted of allUnsupported) {
    let story: FolioDocumentStoryHandle;
    switch (omitted.reason) {
      case "story-missing-in-base":
      case "story-missing-in-target":
      case "story-not-editable":
        story =
          omitted.baseStory ??
          omitted.targetStory ??
          panic("An unsupported story has neither a base nor a target handle");
        break;
      case "numbering-definition":
        story =
          planned.at(0)?.pair.baseStory ??
          panic("A referenced numbering change has no paired story");
        break;
      default:
        story = omitted.story;
        break;
    }
    failures.push({
      invariant: "accept-reproduces-target",
      cause: "unsupported",
      story,
      detail: `the ${omitted.reason} difference has no proved tracked-document instruction`,
    });
  }
  // Each story gets the range that starts where the previous story's ended.
  // A revision `w:id` is scoped to the package, not the part, so two stories
  // seeded alike would let a reader resolving a header revision resolve a
  // body revision with it.
  let idSeed = revisionStamp.idSeed;
  let documentChanged = false;
  for (const { pair, plan, prepared } of preparedStories) {
    changes.push(...plan.changes);
    const executed = comparisonAccess.commitStoryProgram({
      story: pair.baseStory,
      revisionStamp: { date: revisionStamp.date, idSeed },
      prepared,
    });
    if (executed.status === "unsupported") {
      return Result.err(
        new CompareDocxApplyError({
          message: "A preflighted story could not complete its atomic comparison transaction.",
          story: pair.baseStory,
          reason: executed.issue.reason,
        }),
      );
    }
    idSeed = executed.receipt.nextRevisionId;
    documentChanged ||= executed.receipt.transaction.docChanged;
  }

  // Verification is two package projections, independent of story count.
  // Projecting inside the loop above would rebuild the whole live Document
  // twice per story and turn multi-part comparisons quadratic in practice.
  const acceptedByStory = new Map(
    comparisonAccess
      .projectReviewedStories("final")
      .stories.map(({ handle, snapshot }) => [storyKey(handle), snapshot] as const),
  );
  const rejectedByStory = new Map(
    comparisonAccess
      .projectReviewedStories("original")
      .stories.map(({ handle, snapshot }) => [storyKey(handle), snapshot] as const),
  );
  for (const { pair } of preparedStories) {
    const baseBefore = resolvedDocxContentBlocks(pair.baseSnapshot);
    const baseBeforeGeometry = projectTableGeometry(
      storyTablesOf(resolvedDocxOperationSnapshot(pair.baseSnapshot)),
    );
    const acceptedSnapshot = acceptedByStory.get(storyKey(pair.baseStory)) ?? null;
    const acceptFailure = classifyContentProjectionMismatch({
      invariant: "accept-reproduces-target",
      story: pair.baseStory,
      actual: acceptedSnapshot ? resolvedDocxContentBlocks(acceptedSnapshot) : [],
      expected: resolvedDocxContentBlocks(pair.targetSnapshot),
    });
    if (acceptFailure) {
      failures.push(acceptFailure);
    }
    const rejectedSnapshot = rejectedByStory.get(storyKey(pair.baseStory)) ?? null;
    const rejectFailure = classifyContentProjectionMismatch({
      invariant: "reject-reproduces-base",
      story: pair.baseStory,
      actual: rejectedSnapshot ? resolvedDocxContentBlocks(rejectedSnapshot) : [],
      expected: baseBefore,
    });
    if (rejectFailure) {
      failures.push(rejectFailure);
    }
    // The block projection says which cell every paragraph landed in and
    // nothing about the cell. A table's own properties need their own
    // comparison, or a redline that reproduces every word and none of the
    // widths, spans, shading or borders passes the self-check. Checked after
    // the blocks, because a block in the wrong container is the finding a
    // reader needs first and a geometry difference follows from it.
    const geometryAcceptFailure = classifyGeometryMismatch({
      invariant: "accept-reproduces-target",
      story: pair.baseStory,
      actual: projectTableGeometry(
        acceptedSnapshot
          ? storyTablesOf(resolvedDocxOperationSnapshot(acceptedSnapshot))
          : [],
      ),
      expected: projectTableGeometry(
        storyTablesOf(resolvedDocxOperationSnapshot(pair.targetSnapshot)),
      ),
    });
    if (geometryAcceptFailure) {
      failures.push(geometryAcceptFailure);
    }
    const geometryRejectFailure = classifyGeometryMismatch({
      invariant: "reject-reproduces-base",
      story: pair.baseStory,
      actual: projectTableGeometry(
        rejectedSnapshot
          ? storyTablesOf(resolvedDocxOperationSnapshot(rejectedSnapshot))
          : [],
      ),
      expected: baseBeforeGeometry,
    });
    if (geometryRejectFailure) {
      failures.push(geometryRejectFailure);
    }
  }
  return Result.ok({
    changes,
    unsupported: allUnsupported,
    verification:
      failures.length === 0 ? { status: "verified" } : { status: "unverified", failures },
    documentChanged,
  });
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
  { baseBuffer, baseCarriedRevisions, reviewer, packageDate, revisionStamp }: ParsedComparison,
  { documentChanged }: AppliedComparison,
): Promise<Result<ArrayBuffer, CompareDocxSerializeError | CompareDocxFinalParagraphMarkError>> => {
  if (!baseCarriedRevisions && !documentChanged) {
    return Result.ok(baseBuffer);
  }
  const document = reviewer.toDocument();
  // A revision on the paragraph that ends a container asks a consumer to merge
  // it with a paragraph that is not there, or to close a break back over one,
  // and neither is an edit that can be carried out. Nothing downstream can
  // recover from that, so it is fatal in either comparison mode:
  // unlike an unproven redline there is no partial result worth handing back.
  //
  // Scoped to the revisions this comparison minted. A base can arrive carrying
  // one on a paragraph of a part no story mounts, so resolving to its accepted
  // view does not reach it; folio preserves what it parses, and refusing the
  // comparison would report the base's own bytes as this call's doing.
  const revisions = revisedFinalParagraphMarks(document, { since: revisionStamp.idSeed });
  const [first] = revisions;
  if (first !== undefined) {
    return Result.err(
      new CompareDocxFinalParagraphMarkError({
        message:
          `A container's final paragraph mark carries a ${first.kind}, which no ` +
          `consumer can resolve: ${first.container} paragraph ` +
          `${String(first.paragraphIndex)}.`,
        revisions,
      }),
    );
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
 *
 * The result is verified by default: a redline whose round trip cannot be
 * proven is refused rather than returned, because one that reads plausibly and
 * is wrong is worse than none. `mode: "bestEffort"` asks for the opposite
 * trade — the best redline available, plus the typed list of what could not be
 * represented — for a caller that would rather show something and say what is
 * missing.
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
  const unsupported = Object.freeze([
    ...parsed.value.unsupported,
    ...planned.value.flatMap(({ plan }) => plan.unsupported),
    ...parsed.value.numberingChanges.map(({ numId, level }) => ({
      reason: "numbering-definition" as const,
      numId,
      level,
    })),
  ]);
  const applied = applyComparison(parsed.value, planned.value, {
    mode: options.mode ?? "strict",
    unsupported,
  });
  if (applied.isErr()) {
    return Result.err(applied.error);
  }
  const { changes, verification } = applied.value;
  if (verification.status === "unverified" && (options.mode ?? "strict") === "strict") {
    const [firstFailure] = verification.failures;
    if (firstFailure === undefined) {
      panic("An unverified comparison reported no failing invariant");
    }
    return Result.err(
      new CompareDocxRoundTripError({
        message: `The generated tracked changes do not satisfy ${firstFailure.invariant}: ${firstFailure.detail}`,
        story: firstFailure.story,
        invariant: firstFailure.invariant,
        cause: firstFailure.cause,
        failures: verification.failures,
      }),
    );
  }
  const serialized = await serializeComparison(parsed.value, applied.value);
  if (serialized.isErr()) {
    return Result.err(serialized.error);
  }
  return Result.ok({
    buffer: serialized.value,
    changes,
    verification,
    unsupported: applied.value.unsupported,
  });
};
