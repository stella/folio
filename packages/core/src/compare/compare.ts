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
 * An unproven redline is refused by default. `onUnverified: "emit"` returns it
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
import type { Node as PMNode } from "prosemirror-model";

import {
  FolioDocxReviewer,
  type FolioDocumentStoryHandle,
  type FolioNumberingLevel,
  type FolioRevisionStamp,
} from "../ai-edits/headless";
import { projectTableGeometry } from "../ai-edits/table-geometry";
import type { FolioTableTemplates } from "../ai-edits/table-template";
import type { FolioAIBlock, FolioAIEditSkipReason, FolioAIEditSnapshot } from "../ai-edits/types";
import type { WordDiffGranularity } from "../ai-edits/word-diff";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { pairFolioDocumentStories } from "../document-stories";
import { planStoryCompare, type CompareStoryPlan, type CompareTableTemplateRequest } from "./plan";
import { withFixedPackageDates } from "./reproducible-package";
import {
  CompareDocxApplyError,
  CompareDocxFinalParagraphMarkError,
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
import {
  classifyGeometryMismatch,
  classifyProjectionMismatch,
  revisedFinalParagraphMarks,
  projectSupportedInlineFormatting,
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
const projectBlocks = (blocks: readonly FolioAIBlock[]): string[] => {
  return blocks.map(({ text, table, styleId, listLevel }) => {
    const container = table
      ? `t${String(table.tableIndex)}r${String(table.rowIndex)}c${String(table.cellIndex)}g${String(table.gridColumnIndex)}x${String(table.columnSpan)}y${String(table.rowSpan)}p${String(table.paragraphIndex)}`
      : "body";
    // The properties the comparison claims to compare are in the projection
    // too, or the self-check would pass a redline that reproduces every word
    // and leaves a list item at the wrong level.
    return `${container}|${styleId ?? ""}|${listLevel ?? ""}|${text}`;
  });
};

type FormattingRoundTripFailureOptions = {
  invariant: CompareVerificationFailure["invariant"];
  story: FolioDocumentStoryHandle;
  changes: readonly CompareChange[];
  actualBlocks: readonly FolioAIBlock[];
  expectedBlocks: readonly FolioAIBlock[];
  expectedBlockId: (change: Extract<CompareChange, { kind: "format" }>) => string;
};

/** Verify formatting only where the plan claims a text-equal formatting change. */
const formattingRoundTripFailure = ({
  invariant,
  story,
  changes,
  actualBlocks,
  expectedBlocks,
  expectedBlockId,
}: FormattingRoundTripFailureOptions): CompareVerificationFailure | null => {
  const expectedIndexById = new Map(expectedBlocks.map(({ id }, index) => [id, index]));
  const checkedExpectedIds = new Set<string>();
  for (const change of changes) {
    if (change.kind !== "format") {
      continue;
    }
    const expectedId = expectedBlockId(change);
    if (checkedExpectedIds.has(expectedId)) {
      continue;
    }
    checkedExpectedIds.add(expectedId);
    const expectedIndex = expectedIndexById.get(expectedId) ?? -1;
    const expected = expectedBlocks.at(expectedIndex);
    const actual = actualBlocks.at(expectedIndex);
    if (expectedIndex === -1 || !actual || !expected) {
      return {
        invariant,
        cause: "inline-formatting",
        story,
        detail: "a text-equal aligned block could not be projected for formatting verification",
      };
    }
    if (projectSupportedInlineFormatting(actual) !== projectSupportedInlineFormatting(expected)) {
      return {
        invariant,
        cause: "inline-formatting",
        story,
        detail: "supported inline formatting differs in a text-equal aligned block",
      };
    }
  }
  return null;
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
  for (const { handle } of reviewer.listStories()) {
    reviewer.resolveReviewedStory({ story: handle, view: "final" });
  }
  for (const { handle } of targetReviewer.listStories()) {
    targetReviewer.resolveReviewedStory({ story: handle, view: "final" });
  }
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
      targetSnapshot: pair.targetSnapshot,
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
 * What a skipped operation says about the plan that derived it.
 *
 * Two different things go wrong at apply time, and only one of them leaves the
 * result unusable. A reason that says the plan did not match the document it
 * was planned against is an engine defect — the operations came from this very
 * snapshot moments earlier, so nothing should have moved under them, and a
 * redline built on the rest is built on a document the plan no longer
 * describes. A reason that says the applier had nothing to write, or could not
 * write that shape THERE, leaves the redline standing and turns the question
 * into "was anything lost", which is what the round-trip check answers a few
 * lines below. Refusing on those instead trades a partial answer for none, and
 * refuses a whole document because one paragraph sits inside a structure the
 * block snapshot does not model — a text box, a content control — where a
 * paragraph mark has nowhere to go.
 */
const COMPARE_SKIP_DISPOSITION = {
  missingBlock: "fatal",
  changedBlock: "fatal",
  ambiguousFind: "fatal",
  missingFind: "fatal",
  unsupportedBlock: "unwritable",
  unsupportedMode: "fatal",
  atomicBatchRejected: "fatal",
  preconditionFailed: "fatal",
  staleRange: "fatal",
  emptyOperation: "unwritable",
  noopOperation: "unwritable",
  documentVersionMismatch: "fatal",
  documentNotEditable: "fatal",
} as const satisfies Record<FolioAIEditSkipReason, "fatal" | "unwritable">;

/** What stage 3 produced: the change list, whether it was proven, and whether it wrote anything. */
export type AppliedComparison = {
  changes: readonly CompareChange[];
  verification: CompareVerification;
  /**
   * Whether the stage wrote into the base document at all. A comparison that
   * found nothing writes nothing, and the serialize stage then hands the
   * arriving bytes back rather than reproducing them.
   */
  documentChanged: boolean;
};

/**
 * The table each `insertTable` / `insertTableRow` in the plan should place,
 * resolved against the target document the plan named it in.
 *
 * The plan is pure and names a table by index; the node lives in the other
 * package, which only this stage holds. A request naming a table the target
 * does not have resolves to nothing and the operation falls back to its cell
 * texts, which is the same redline the comparison produced before.
 */
const resolveTableTemplates = (
  targetTables: ReadonlyMap<number, PMNode>,
  requests: readonly CompareTableTemplateRequest[],
): FolioTableTemplates => {
  const templates = new Map<string, PMNode>();
  for (const { operationId, targetTableIndex, targetRowIndex } of requests) {
    const table = targetTables.get(targetTableIndex);
    if (!table) {
      continue;
    }
    if (targetRowIndex === undefined) {
      templates.set(operationId, table);
      continue;
    }
    const row = table.maybeChild(targetRowIndex);
    if (row) {
      templates.set(operationId, row);
    }
  }
  return templates;
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
  { reviewer, targetReviewer, revisionStamp, granularity, numberingChanges }: ParsedComparison,
  planned: readonly PlannedStoryComparison[],
): Result<AppliedComparison, CompareDocxApplyError> => {
  const changes: CompareChange[] = [...numberingChanges];
  const failures: CompareVerificationFailure[] = [];
  // Each story gets the range that starts where the previous story's ended.
  // A revision `w:id` is scoped to the package, not the part, so two stories
  // seeded alike would let a reader resolving a header revision resolve a
  // body revision with it.
  let idSeed = revisionStamp.idSeed;
  let documentChanged = false;
  for (const { pair, plan } of planned) {
    changes.push(...plan.changes);
    if (plan.operations.length === 0 && plan.tableGeometryPairings.length === 0) {
      continue;
    }

    // Read before anything lands: this is the document the redline is written
    // against, and rejecting every revision has to return to it.
    const baseBeforeBlocks =
      reviewer.readReviewedStory({ story: pair.baseStory, view: "final" })?.snapshot.blocks ?? [];
    const baseBefore = projectBlocks(baseBeforeBlocks);
    const baseBeforeGeometry = projectTableGeometry(
      reviewer.storyTables({ story: pair.baseStory }),
    );
    const targetTables = new Map(
      targetReviewer
        .storyTables({ story: pair.targetStory })
        .map(({ index, node }) => [index, node] as const),
    );

    // Table properties first, while the base's table indices still describe
    // the document the plan was written against: the operations below add and
    // remove tables, and every index past the first one would then name a
    // different table.
    const afterGeometry = reviewer.matchStoryTableGeometry({
      story: pair.baseStory,
      targetTables,
      pairings: plan.tableGeometryPairings,
      revisionStamp: { date: revisionStamp.date, idSeed },
    });
    const geometryChanged = afterGeometry > idSeed;
    documentChanged ||= geometryChanged;
    idSeed = afterGeometry;

    if (plan.operations.length === 0 && !geometryChanged) {
      continue;
    }

    if (plan.operations.length > 0) {
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
        tableTemplates: resolveTableTemplates(targetTables, plan.tableTemplates),
      });
      if (nextRevisionId === undefined) {
        // Only a host bridge that does not allocate ids itself omits this, and
        // the comparison drives the in-process applier.
        panic("The applier did not report where it left the revision-id counter", {
          story: pair.baseStory,
        });
      }
      idSeed = nextRevisionId;
      documentChanged = true;
      const refused = skipped.filter(({ reason }) => COMPARE_SKIP_DISPOSITION[reason] === "fatal");
      if (refused.length > 0) {
        return Result.err(
          new CompareDocxApplyError({
            message:
              "Some derived operations were refused, so the result would not match the target.",
            skipped: refused,
          }),
        );
      }
    }

    const acceptedStory = reviewer.readReviewedStory({ story: pair.baseStory, view: "final" });
    const acceptFailure = classifyProjectionMismatch({
      invariant: "accept-reproduces-target",
      story: pair.baseStory,
      actual: projectBlocks(acceptedStory?.snapshot.blocks ?? []),
      expected: projectBlocks(pair.targetSnapshot.blocks),
    });
    if (acceptFailure) {
      failures.push(acceptFailure);
    } else {
      const formattingFailure = formattingRoundTripFailure({
        invariant: "accept-reproduces-target",
        story: pair.baseStory,
        changes: plan.changes,
        actualBlocks: acceptedStory?.snapshot.blocks ?? [],
        expectedBlocks: pair.targetSnapshot.blocks,
        expectedBlockId: ({ targetBlockId }) => targetBlockId,
      });
      if (formattingFailure) {
        failures.push(formattingFailure);
      }
    }
    const rejectedStory = reviewer.readReviewedStory({ story: pair.baseStory, view: "original" });
    const rejectFailure = classifyProjectionMismatch({
      invariant: "reject-reproduces-base",
      story: pair.baseStory,
      actual: projectBlocks(rejectedStory?.snapshot.blocks ?? []),
      expected: baseBefore,
    });
    if (rejectFailure) {
      failures.push(rejectFailure);
    } else {
      const formattingFailure = formattingRoundTripFailure({
        invariant: "reject-reproduces-base",
        story: pair.baseStory,
        changes: plan.changes,
        actualBlocks: rejectedStory?.snapshot.blocks ?? [],
        expectedBlocks: pair.baseSnapshot.blocks,
        expectedBlockId: ({ baseBlockId }) => baseBlockId,
      });
      if (formattingFailure) {
        failures.push(formattingFailure);
      }
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
      actual: projectTableGeometry(reviewer.storyTables({ story: pair.baseStory, view: "final" })),
      expected: projectTableGeometry(targetReviewer.storyTables({ story: pair.targetStory })),
    });
    if (geometryAcceptFailure) {
      failures.push(geometryAcceptFailure);
    }
    const geometryRejectFailure = classifyGeometryMismatch({
      invariant: "reject-reproduces-base",
      story: pair.baseStory,
      actual: projectTableGeometry(
        reviewer.storyTables({ story: pair.baseStory, view: "original" }),
      ),
      expected: baseBeforeGeometry,
    });
    if (geometryRejectFailure) {
      failures.push(geometryRejectFailure);
    }
  }
  return Result.ok({
    changes,
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
  { baseBuffer, baseCarriedRevisions, reviewer, packageDate }: ParsedComparison,
  { documentChanged }: AppliedComparison,
): Promise<Result<ArrayBuffer, CompareDocxSerializeError | CompareDocxFinalParagraphMarkError>> => {
  if (!baseCarriedRevisions && !documentChanged) {
    return Result.ok(baseBuffer);
  }
  const document = reviewer.toDocument();
  // A revision on the paragraph that ends a container asks a consumer to merge
  // it with a paragraph that is not there, or to close a break back over one,
  // and neither is an edit that can be carried out. Nothing downstream can
  // recover from that, so it is fatal under either `onUnverified` setting:
  // unlike an unproven redline there is no partial result worth handing back.
  const revisions = revisedFinalParagraphMarks(document);
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
 * is wrong is worse than none. `onUnverified: "emit"` asks for the opposite
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
  const applied = applyComparison(parsed.value, planned.value);
  if (applied.isErr()) {
    return Result.err(applied.error);
  }
  const { changes, verification } = applied.value;
  if (verification.status === "unverified" && (options.onUnverified ?? "refuse") === "refuse") {
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
    unsupported: parsed.value.unsupported,
  });
};
