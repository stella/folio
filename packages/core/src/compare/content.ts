/**
 * Pure comparison of two ordered, representation-neutral content snapshots.
 *
 * Every event carries owned blocks or canonical base/revised range relations.
 * Projecting those discriminated events reconstructs either snapshot without
 * another alignment, text diff, or formatting comparison.
 */

import { panic, Result, TaggedError } from "better-result";
import {
  createWordDiffSession,
  WORD_DIFF_GRANULARITIES,
  type WordDiffGranularity,
  type WordDiffSegment,
} from "./text-diff";
import { pairedInlineFormattingSegments } from "./formatting";
import { changedFolioContentProperties } from "./content-properties";
import {
  alignFolioContentStructure,
  contentBlocksShareContainer,
  createFolioContentAlignmentWorkSession,
  folioContentIdentityPairDisposition,
  type FolioContentAlignmentStep,
  type FolioContentAlignmentWorkSession,
} from "./content-alignment";
import type {
  FolioContentBlock,
  FolioContentIdentity,
  FolioContentInlineFormattingChange,
  FolioContentParagraphFormatting,
  FolioContentPropertyChange,
  FolioContentPropertyInput,
  FolioContentPropertySet,
  FolioContentPropertyValue,
  FolioContentRun,
  FolioContentSnapshot,
} from "./content-types";
import {
  FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_IDENTITY_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_IDENTITY_SEMANTICS,
  FOLIO_CONTENT_PARAGRAPH_FORMATTING_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_PROPERTY_ARRAY_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_PROPERTY_ENTRY_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_PROPERTY_OBJECT_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_RUN_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_SNAPSHOT_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS,
} from "./content-types";

/** Hard resource ceilings for one representation-neutral comparison work session. */
export const FOLIO_CONTENT_COMPARISON_LIMITS = Object.freeze({
  storiesPerSession: 4_096,
  blocksPerSnapshot: 100_000,
  events: 200_000,
  changes: 10_000,
  formattingRanges: 10_000,
  structuralMembers: 100_000,
  blockCodeUnits: 1_048_576,
  textCodeUnitsPerSnapshot: 8_000_000,
  runsPerBlock: 65_536,
  runsPerSnapshot: 1_000_000,
  structuralBoundariesPerBlock: 65_536,
  structuralBoundariesPerSnapshot: 1_000_000,
  containerDepth: 64,
  containerEntriesPerSnapshot: 1_000_000,
  propertyDepth: 32,
  propertyEntriesPerContainer: 65_536,
  propertyNodesPerSnapshot: 2_000_000,
  attributeCodeUnits: 16_384,
  attributeCodeUnitsPerSnapshot: 8_000_000,
} as const);

export type FolioContentComparisonLimit = keyof typeof FOLIO_CONTENT_COMPARISON_LIMITS;

const DENSE_ARRAY_CAPTURE_LIMITS = Object.freeze({
  blocksPerSnapshot: FOLIO_CONTENT_COMPARISON_LIMITS.blocksPerSnapshot,
  runsPerBlock: FOLIO_CONTENT_COMPARISON_LIMITS.runsPerBlock,
  structuralBoundariesPerBlock:
    FOLIO_CONTENT_COMPARISON_LIMITS.structuralBoundariesPerBlock,
  containerDepth: FOLIO_CONTENT_COMPARISON_LIMITS.containerDepth,
  propertyEntriesPerContainer:
    FOLIO_CONTENT_COMPARISON_LIMITS.propertyEntriesPerContainer,
} as const);

type DenseArrayCaptureLimit = keyof typeof DENSE_ARRAY_CAPTURE_LIMITS;

/** Words a one-sided block needs before it may be classified as a move. */
const MOVE_MINIMUM_WORD_COUNT = 3;

/** Same-text move candidates retained for one repeated value. */
const MAX_MOVE_CANDIDATES_PER_TEXT = 64;

/** Pairwise comparisons allowed for edited-move discovery. */
const MAX_MOVE_SIMILARITY_COMPARISONS = 20_000;

/** Map lookups allowed across edited-move similarity scoring. */
const MAX_MOVE_SIMILARITY_TOKEN_LOOKUPS = 4_000_000;

/** Minimum multiset Dice similarity for an edited relocation. */
const MOVE_SIMILARITY_THRESHOLD = 0.8;

/** Bound token material retained for one edited-move candidate. */
const MAX_MOVE_PROFILE_TOKENS = 16_384;

/** Bound source code units copied into one edited-move token profile. */
const MAX_MOVE_PROFILE_CODE_UNITS = 1_048_576;

export class InvalidFolioContentComparisonError extends TaggedError(
  "InvalidFolioContentComparisonError",
)<{
  message: string;
  input: "options" | "base" | "revised";
  blockIndex?: number;
  field: string;
}> {}

export class FolioContentComparisonLimitError extends TaggedError(
  "FolioContentComparisonLimitError",
)<{
  message: string;
  input: "base" | "revised" | "result" | "session";
  limit: FolioContentComparisonLimit;
  maximum: number;
  actual: number;
  blockIndex?: number;
  field?: string;
}> {}

export class FolioContentComparisonSessionError extends TaggedError(
  "FolioContentComparisonSessionError",
)<{
  message: string;
  reason: "operation-consumed" | "operation-active" | "session-poisoned";
}> {}

/** Errors returned by {@link compareContent}. */
export type FolioContentComparisonError =
  | InvalidFolioContentComparisonError
  | FolioContentComparisonLimitError;

/** One Folio word-diff segment with JavaScript-slice-compatible offsets. */
export type FolioContentTextSegment = {
  readonly type: "equal" | "del" | "ins";
  readonly text: string;
  readonly baseStart: number;
  readonly baseEnd: number;
  readonly revisedStart: number;
  readonly revisedEnd: number;
};

/** Presentation differences for one text-aligned block pair. */
export type FolioContentFormatRange = {
  readonly baseStart: number;
  readonly baseEnd: number;
  readonly revisedStart: number;
  readonly revisedEnd: number;
  readonly formatting: FolioContentInlineFormattingChange;
};

/** Authored and effective paragraph-property deltas for one paired block. */
export type FolioContentParagraphFormattingChange = FolioContentInlineFormattingChange;

/** Presentation differences for one text-aligned block pair. */
export type FolioContentFormattingChange = {
  readonly paragraph: FolioContentParagraphFormattingChange;
  readonly ranges: readonly FolioContentFormatRange[];
};

type BlockFieldDescriptor =
  (typeof FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS)[keyof typeof FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS];

export type FolioContentBlockProperty = Extract<
  BlockFieldDescriptor,
  { role: "kind" | "block-property" | "structure" }
>["field"];

export type FolioContentValuePresence<Value> =
  | { readonly type: "absent" }
  | { readonly type: "present"; readonly value: Value };

/** Exact non-presentation differences for one paired block relation. */
export type FolioContentBlockChange =
  | { readonly field: "kind"; readonly base: string; readonly revised: string }
  | { readonly field: "blockProperties"; readonly changes: readonly FolioContentPropertyChange[] }
  | {
      readonly field: "structuralBoundaries";
      readonly base: readonly FolioContentBlock["structuralBoundaries"][number][];
      readonly revised: readonly FolioContentBlock["structuralBoundaries"][number][];
    }
  | {
      readonly field: "table";
      readonly base: FolioContentValuePresence<NonNullable<FolioContentBlock["table"]>>;
      readonly revised: FolioContentValuePresence<NonNullable<FolioContentBlock["table"]>>;
    }
  | {
      readonly field: "containerPath";
      readonly base: FolioContentBlock["containerPath"];
      readonly revised: FolioContentBlock["containerPath"];
    };

/** An owned absolute UTF-16 range within one compared block. */
export type FolioContentBlockRange = {
  readonly block: FolioContentBlock;
  readonly startOffset: number;
  readonly endOffset: number;
};

/** Complete semantics for one surviving base/revised block-range pairing. */
type FolioContentPairRelationFields = {
  readonly base: FolioContentBlockRange;
  readonly revised: FolioContentBlockRange;
  readonly segments: readonly FolioContentTextSegment[];
};

type FolioContentSemanticPairRelationFields = FolioContentPairRelationFields & {
  readonly blockChanges: readonly FolioContentBlockChange[];
  readonly formatting: FolioContentFormattingChange | null;
};

/** A surviving pairing covering both complete blocks. */
export type FolioContentWholePairRelation = FolioContentSemanticPairRelationFields & {
  readonly relationType: "whole";
};

/** One surviving correspondence inside a split or merge. */
export type FolioContentRangePairRelation = FolioContentSemanticPairRelationFields & {
  readonly relationType: "range";
};

/** Text present only between split/merge survivors on one side. */
export type FolioContentSeparatorRelation = FolioContentPairRelationFields & {
  readonly relationType: "separator";
  readonly blockChanges: readonly [];
  readonly formatting: null;
};

export type FolioContentPairRelation =
  | FolioContentWholePairRelation
  | FolioContentRangePairRelation
  | FolioContentSeparatorRelation;

/** One relocation shared by its source and destination stream positions. */
export type FolioContentMove = {
  readonly id: number;
  readonly relation: FolioContentWholePairRelation;
};

export type FolioContentBlockGroup = readonly [FolioContentBlock, ...FolioContentBlock[]];

/** One structural change owning its non-empty ordered member set. */
export type FolioContentStructuralChange =
  | {
      readonly type: "table-delete";
      readonly tableIndex: number;
      readonly blocks: FolioContentBlockGroup;
    }
  | {
      readonly type: "table-insert";
      readonly tableIndex: number;
      readonly blocks: FolioContentBlockGroup;
    }
  | {
      readonly type: "table-row-delete";
      readonly tableIndex: number;
      readonly rowIndex: number;
      readonly blocks: FolioContentBlockGroup;
    }
  | {
      readonly type: "table-row-insert";
      readonly tableIndex: number;
      readonly rowIndex: number;
      readonly blocks: FolioContentBlockGroup;
    }
  | {
      readonly type: "table-column-delete";
      readonly tableIndex: number;
      readonly columnIndex: number;
      readonly blocks: FolioContentBlockGroup;
    }
  | {
      readonly type: "table-column-insert";
      readonly tableIndex: number;
      readonly columnIndex: number;
      readonly blocks: FolioContentBlockGroup;
      readonly anchor: { readonly blockId: string; readonly position: "after" | "before" };
    };

/** One row-major stream position belonging to a shared structural change. */
export type FolioContentStructuralEvent = {
  readonly type: "structural";
  readonly change: FolioContentStructuralChange;
  /** Index into the change-owned member tuple at this row-major stream position. */
  readonly memberIndex: number;
};

/**
 * One structurally incompatible table pair. The nested comparison is computed
 * in the same bounded semantic pass and is available only as a lowering
 * refinement; the outer event remains the sole document-order occurrence.
 */
export type FolioContentTableReplacement = {
  readonly baseBlocks: FolioContentBlockGroup;
  readonly revisedBlocks: FolioContentBlockGroup;
  readonly baseTableIndex: number;
  readonly revisedTableIndex: number;
  readonly refinement: FolioContentComparison;
};

/**
 * One item in the complete comparison stream. Every surviving pairing has
 * one canonical relation; one-sided events carry exactly one owned block.
 */
export type FolioContentComparisonEvent =
  | { readonly type: "unchanged"; readonly relation: FolioContentWholePairRelation }
  | { readonly type: "modified"; readonly relation: FolioContentWholePairRelation }
  | { readonly type: "formatting"; readonly relation: FolioContentWholePairRelation }
  | { readonly type: "inserted"; readonly block: FolioContentBlock }
  | { readonly type: "deleted"; readonly block: FolioContentBlock }
  | { readonly type: "movedFrom"; readonly move: FolioContentMove }
  | { readonly type: "movedTo"; readonly move: FolioContentMove }
  | {
      readonly type: "split";
      readonly relations: readonly [FolioContentRangePairRelation, FolioContentRangePairRelation];
      readonly separator: FolioContentSeparatorRelation;
    }
  | {
      readonly type: "merge";
      readonly relations: readonly [FolioContentRangePairRelation, FolioContentRangePairRelation];
      readonly separator: FolioContentSeparatorRelation;
    }
  | {
      readonly type: "tableReplacement";
      readonly replacement: FolioContentTableReplacement;
    }
  | FolioContentStructuralEvent;

const FOLIO_CONTENT_COMPARISON_BRAND: unique symbol = Symbol("FolioContentComparison");

/** Result of one representation-neutral story comparison. */
export type FolioContentComparison = {
  readonly events: readonly FolioContentComparisonEvent[];
  readonly [FOLIO_CONTENT_COMPARISON_BRAND]: true;
};

const completeContentComparison = (
  events: readonly FolioContentComparisonEvent[],
): FolioContentComparison => {
  const comparison = { events };
  Object.defineProperty(comparison, FOLIO_CONTENT_COMPARISON_BRAND, {
    value: true,
    enumerable: false,
    writable: false,
    configurable: false,
  });
  // SAFETY: defineProperty installs the private nominal brand immediately above.
  return Object.freeze(comparison) as FolioContentComparison;
};

/** Inputs to {@link compareContent}. */
export type CompareContentOptions = {
  base: FolioContentSnapshot;
  revised: FolioContentSnapshot;
  /** Token size for modified-block segments; defaults to `"word"`. */
  granularity?: WordDiffGranularity;
};

const COMPARE_CONTENT_OPTION_FIELD_DESCRIPTORS = Object.freeze({
  base: Object.freeze({ field: "base" }),
  revised: Object.freeze({ field: "revised" }),
  granularity: Object.freeze({ field: "granularity" }),
} as const satisfies {
  [Field in keyof CompareContentOptions]-?: { readonly field: Field };
});

const SNAPSHOT_RESOURCE_DESCRIPTORS = [
  { resource: "blocks", limit: "blocksPerSnapshot" },
  { resource: "textCodeUnits", limit: "textCodeUnitsPerSnapshot" },
  { resource: "runs", limit: "runsPerSnapshot" },
  { resource: "structuralBoundaries", limit: "structuralBoundariesPerSnapshot" },
  { resource: "containerEntries", limit: "containerEntriesPerSnapshot" },
  { resource: "propertyNodes", limit: "propertyNodesPerSnapshot" },
  { resource: "attributeCodeUnits", limit: "attributeCodeUnitsPerSnapshot" },
] as const satisfies readonly {
  resource: string;
  limit: FolioContentComparisonLimit;
}[];

type SnapshotResource = (typeof SNAPSHOT_RESOURCE_DESCRIPTORS)[number]["resource"];
type SnapshotResourceUsage = Record<SnapshotResource, number>;

type ContentComparisonResourceUsage = {
  base: SnapshotResourceUsage;
  revised: SnapshotResourceUsage;
  stories: number;
  changes: number;
  formattingRanges: number;
  structuralMembers: number;
  events: number;
};

const emptySnapshotResourceUsage = (): SnapshotResourceUsage => ({
  blocks: 0,
  textCodeUnits: 0,
  runs: 0,
  structuralBoundaries: 0,
  containerEntries: 0,
  propertyNodes: 0,
  attributeCodeUnits: 0,
});

type CapturedContentComparison = {
  readonly base: FolioContentSnapshot;
  readonly revised: FolioContentSnapshot;
};

type ContentComparisonExecution = {
  readonly comparison: FolioContentComparison;
  readonly changes: number;
  readonly formattingRanges: number;
  readonly structuralMembers: number;
  readonly events: number;
};

type ContentComparisonOperationState =
  | {
      status: "ready";
      run: () => Result<FolioContentComparison, FolioContentComparisonLimitError>;
    }
  | { status: "consumed" };

class FolioContentComparisonOperation {
  #state: ContentComparisonOperationState;

  private constructor(
    run: () => Result<FolioContentComparison, FolioContentComparisonLimitError>,
  ) {
    this.#state = { status: "ready", run };
  }

  static create(
    run: () => Result<FolioContentComparison, FolioContentComparisonLimitError>,
  ): FolioContentComparisonOperation {
    return new FolioContentComparisonOperation(run);
  }

  compare(): Result<
    FolioContentComparison,
    FolioContentComparisonLimitError | FolioContentComparisonSessionError
  > {
    if (this.#state.status === "consumed") {
      return Result.err(
        new FolioContentComparisonSessionError({
          message: "A content comparison operation can only be consumed once.",
          reason: "operation-consumed",
        }),
      );
    }
    const { run } = this.#state;
    this.#state = { status: "consumed" };
    return run();
  }
}

type ContentComparisonSessionState =
  | { status: "ready" }
  | { status: "captured"; operationId: symbol }
  | { status: "poisoned" };

type CaptureContentComparisonOptions = {
  base: unknown;
  revised: unknown;
};

type ContentComparisonWorkSessionOptions = {
  granularity?: WordDiffGranularity;
  /** @internal A controlled diff used by focused ownership/parity tests. */
  diffText?: (base: string, revised: string) => WordDiffSegment[];
  /** @internal Lower allowances used by focused boundary tests. */
  moveComparisons?: number;
  moveTokenLookups?: number;
};

type MoveSimilarityResult =
  | { status: "exhausted" }
  | { status: "unscored" }
  | { status: "scored"; similarity: number };

type ContentComparisonEngine = {
  alignContentStructure: <Block extends FolioContentBlock>(
    options: Omit<Parameters<typeof alignFolioContentStructure<Block>>[0], "workSession">,
  ) => FolioContentAlignmentStep<Block>[];
  diffText: (base: string, revised: string) => WordDiffSegment[];
  scoreMoveSimilarity: (base: TokenProfile, revised: TokenProfile) => MoveSimilarityResult;
};

const boundedSessionAllowance = (value: number | undefined, maximum: number, field: string): number => {
  const allowance = value ?? maximum;
  if (!Number.isSafeInteger(allowance) || allowance < 0 || allowance > maximum) {
    return panic("A content comparison allowance is outside the supported range", {
      field,
      allowance,
      maximum,
    });
  }
  return allowance;
};

export class FolioContentComparisonWorkSession {
  readonly #alignment: FolioContentAlignmentWorkSession;
  readonly #diffText: ReturnType<typeof createWordDiffSession>["diff"];
  #remainingMoveComparisons: number;
  #remainingMoveTokenLookups: number;
  #resourceUsage: ContentComparisonResourceUsage = {
    base: emptySnapshotResourceUsage(),
    revised: emptySnapshotResourceUsage(),
    stories: 0,
    changes: 0,
    formattingRanges: 0,
    structuralMembers: 0,
    events: 0,
  };
  #state: ContentComparisonSessionState = { status: "ready" };
  #executing = false;

  private constructor({
    granularity,
    diffText,
    moveComparisons,
    moveTokenLookups,
  }: ContentComparisonWorkSessionOptions = {}) {
    this.#alignment = createFolioContentAlignmentWorkSession();
    this.#diffText = diffText ?? createWordDiffSession({ ...(granularity && { granularity }) }).diff;
    this.#remainingMoveComparisons = boundedSessionAllowance(
      moveComparisons,
      MAX_MOVE_SIMILARITY_COMPARISONS,
      "moveComparisons",
    );
    this.#remainingMoveTokenLookups = boundedSessionAllowance(
      moveTokenLookups,
      MAX_MOVE_SIMILARITY_TOKEN_LOOKUPS,
      "moveTokenLookups",
    );
  }

  static create(options: ContentComparisonWorkSessionOptions = {}): FolioContentComparisonWorkSession {
    return new FolioContentComparisonWorkSession(options);
  }

  #assertBudgetUseAllowed(): void {
    if (this.#state.status !== "ready" && !this.#executing) {
      return panic("A content comparison session cannot perform work in its current state", {
        status: this.#state.status,
      });
    }
  }

  #alignContentStructure<Block extends FolioContentBlock>(
    options: Omit<Parameters<typeof alignFolioContentStructure<Block>>[0], "workSession">,
  ): FolioContentAlignmentStep<Block>[] {
    this.#assertBudgetUseAllowed();
    return alignFolioContentStructure({ ...options, workSession: this.#alignment });
  }

  #diff(base: string, revised: string): WordDiffSegment[] {
    this.#assertBudgetUseAllowed();
    return this.#diffText(base, revised);
  }

  #scoreMoveSimilarity(base: TokenProfile, revised: TokenProfile): MoveSimilarityResult {
    this.#assertBudgetUseAllowed();
    if (this.#remainingMoveComparisons <= 0) return { status: "exhausted" };
    this.#remainingMoveComparisons--;
    const scored = tokenSimilarity(base, revised, this.#remainingMoveTokenLookups);
    if (!scored) return { status: "unscored" };
    this.#remainingMoveTokenLookups -= scored.lookups;
    return { status: "scored", similarity: scored.similarity };
  }

  captureComparison({
    base,
    revised,
  }: CaptureContentComparisonOptions): Result<
    FolioContentComparisonOperation,
    FolioContentComparisonError | FolioContentComparisonSessionError
  > {
    if (this.#state.status !== "ready") {
      return Result.err(
        new FolioContentComparisonSessionError({
          message:
            this.#state.status === "poisoned"
              ? "A failed content comparison permanently closes its work session."
              : "The active content comparison operation must be consumed first.",
          reason: this.#state.status === "poisoned" ? "session-poisoned" : "operation-active",
        }),
      );
    }
    const storyCount = this.#resourceUsage.stories + 1;
    if (storyCount > FOLIO_CONTENT_COMPARISON_LIMITS.storiesPerSession) {
      return Result.err(
        limitExceeded({
          input: "session",
          limit: "storiesPerSession",
          maximum: FOLIO_CONTENT_COMPARISON_LIMITS.storiesPerSession,
          actual: storyCount,
        }),
      );
    }
    // Capture itself performs bounded work. Charge the attempt before touching
    // caller data so repeated invalid or near-limit stories cannot bypass the
    // aggregate session ceiling.
    this.#resourceUsage.stories = storyCount;
    const registry = createCaptureRegistry();
    const capturedBase = captureContentSnapshot(base, "base", registry);
    if (capturedBase.isErr()) return Result.err(capturedBase.error);
    const capturedRevised = captureContentSnapshot(revised, "revised", registry);
    if (capturedRevised.isErr()) return Result.err(capturedRevised.error);
    const aggregateError = claimSnapshotPairResources(
      this.#resourceUsage,
      capturedBase.value.usage,
      capturedRevised.value.usage,
    );
    if (aggregateError) return Result.err(aggregateError);
    const captured = Object.freeze({
      base: capturedBase.value.snapshot,
      revised: capturedRevised.value.snapshot,
    });
    const operationId = Symbol("content-comparison-operation");
    this.#state = { status: "captured", operationId };
    return Result.ok(
      FolioContentComparisonOperation.create(() => this.#consumeComparison(operationId, captured)),
    );
  }

  #consumeComparison(
    operationId: symbol,
    captured: CapturedContentComparison,
  ): Result<FolioContentComparison, FolioContentComparisonLimitError> {
    if (this.#state.status !== "captured" || this.#state.operationId !== operationId) {
      return panic("A content comparison operation does not own the active session capture");
    }

    const maximumChanges = FOLIO_CONTENT_COMPARISON_LIMITS.changes - this.#resourceUsage.changes;
    const maximumFormattingRanges =
      FOLIO_CONTENT_COMPARISON_LIMITS.formattingRanges - this.#resourceUsage.formattingRanges;
    const maximumStructuralMembers =
      FOLIO_CONTENT_COMPARISON_LIMITS.structuralMembers - this.#resourceUsage.structuralMembers;
    const maximumEvents = FOLIO_CONTENT_COMPARISON_LIMITS.events - this.#resourceUsage.events;
    // A claimed operation starts from a poisoned default. Only a complete,
    // successful semantic pass reopens the session for another story.
    this.#state = { status: "poisoned" };
    this.#executing = true;
    let executed: ReturnType<typeof executeCapturedContentComparison>;
    try {
      executed = executeCapturedContentComparison({
        captured,
        engine: {
          alignContentStructure: (options) => this.#alignContentStructure(options),
          diffText: (base, revised) => this.#diff(base, revised),
          scoreMoveSimilarity: (base, revised) => this.#scoreMoveSimilarity(base, revised),
        },
        maximumChanges,
        maximumFormattingRanges,
        maximumStructuralMembers,
        maximumEvents,
      });
    } finally {
      this.#executing = false;
    }
    if (executed.isErr()) return Result.err(executed.error);
    this.#resourceUsage.changes += executed.value.changes;
    this.#resourceUsage.formattingRanges += executed.value.formattingRanges;
    this.#resourceUsage.structuralMembers += executed.value.structuralMembers;
    this.#resourceUsage.events += executed.value.events;
    this.#state = { status: "ready" };
    return Result.ok(executed.value.comparison);
  }
}

export const createContentComparisonWorkSession = (
  options: ContentComparisonWorkSessionOptions = {},
): FolioContentComparisonWorkSession => FolioContentComparisonWorkSession.create(options);

type PairedContentStory<Key> = {
  readonly key: Key;
  readonly base: FolioContentSnapshot;
  readonly revised: FolioContentSnapshot;
};

type ComparedContentStory<Key> = {
  readonly key: Key;
  readonly comparison: FolioContentComparison;
};

type CompareContentStoriesError = {
  readonly storyIndex: number;
  readonly cause: FolioContentComparisonError | FolioContentComparisonSessionError;
};

/** Compare already-paired stories through one aggregate bounded work session. @internal */
export const compareContentStories = <Key>({
  stories,
  workSession,
}: {
  stories: readonly PairedContentStory<Key>[];
  workSession: FolioContentComparisonWorkSession;
}): Result<readonly ComparedContentStory<Key>[], CompareContentStoriesError> => {
  const compared: ComparedContentStory<Key>[] = [];
  for (const [storyIndex, story] of stories.entries()) {
    const operation = workSession.captureComparison({ base: story.base, revised: story.revised });
    if (operation.isErr()) {
      return Result.err({ storyIndex, cause: operation.error });
    }
    const result = operation.value.compare();
    if (result.isErr()) {
      return Result.err({ storyIndex, cause: result.error });
    }
    compared.push(Object.freeze({ key: story.key, comparison: result.value }));
  }
  return Result.ok(Object.freeze(compared));
};

const invalidInput = (
  input: "options" | "base" | "revised",
  field: string,
  message: string,
  blockIndex?: number,
): InvalidFolioContentComparisonError =>
  new InvalidFolioContentComparisonError({
    message,
    input,
    field,
    ...(blockIndex !== undefined && { blockIndex }),
  });

const limitExceeded = ({
  input,
  limit,
  maximum,
  actual,
  blockIndex,
  field,
}: {
  input: "base" | "revised" | "result" | "session";
  limit: FolioContentComparisonLimit;
  maximum: number;
  actual: number;
  blockIndex?: number;
  field?: string;
}): FolioContentComparisonLimitError =>
  new FolioContentComparisonLimitError({
    message: `The ${input} content exceeds the ${limit} comparison limit.`,
    input,
    limit,
    maximum,
    actual,
    ...(blockIndex !== undefined && { blockIndex }),
    ...(field !== undefined && { field }),
  });

const isFiniteInteger = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value);

const isFiniteNumber = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value);

const IDENTITY_SEMANTICS = new Set(FOLIO_CONTENT_IDENTITY_SEMANTICS);

type CapturedDataRecord<Field extends string> = ReadonlyMap<Field, unknown>;
type CapturedOwnProperty = Readonly<PropertyDescriptor> | null;
type CapturedPropertyDescriptors = Map<PropertyKey, CapturedOwnProperty>;
type CaptureRegistry = {
  readonly descriptors: WeakMap<object, CapturedPropertyDescriptors>;
  readonly denseArrays: WeakMap<object, readonly unknown[]>;
  readonly arrayIdentity: WeakMap<object, boolean>;
};

const createCaptureRegistry = (): CaptureRegistry => ({
  descriptors: new WeakMap(),
  denseArrays: new WeakMap(),
  arrayIdentity: new WeakMap(),
});

type CaptureLocation = {
  readonly side: "options" | "base" | "revised";
  readonly blockIndex?: number;
  readonly registry?: CaptureRegistry;
};

type CaptureContext = {
  side: "base" | "revised";
  blockIndex?: number;
  usage: SnapshotResourceUsage;
  registry: CaptureRegistry;
};

const captureOwnProperty = (
  input: object,
  key: PropertyKey,
  path: string,
  { side, blockIndex, registry }: CaptureLocation,
): Result<CapturedOwnProperty, InvalidFolioContentComparisonError> => {
  let retained = registry?.descriptors.get(input);
  if (retained?.has(key)) return Result.ok(retained.get(key) ?? null);
  const captured = Result.try({
    try: () => Object.getOwnPropertyDescriptor(input, key),
    catch: () =>
      invalidInput(
        side,
        path,
        "Content comparison values must expose stable own data properties.",
        blockIndex,
      ),
  });
  if (captured.isErr()) return Result.err(captured.error);
  const descriptor = captured.value === undefined ? null : Object.freeze(captured.value);
  if (registry) {
    retained ??= new Map();
    retained.set(key, descriptor);
    registry.descriptors.set(input, retained);
  }
  return Result.ok(descriptor);
};

const captureArrayIdentity = (
  input: object,
  path: string,
  { side, blockIndex, registry }: CaptureLocation,
): Result<boolean, InvalidFolioContentComparisonError> => {
  const retained = registry?.arrayIdentity;
  if (retained?.has(input)) return Result.ok(retained.get(input) ?? false);
  const captured = Result.try({
    try: () => Array.isArray(input),
    catch: () =>
      invalidInput(
        side,
        path,
        "Content comparison values must expose a stable array identity.",
        blockIndex,
      ),
  });
  if (captured.isOk()) retained?.set(input, captured.value);
  return captured;
};

/**
 * Capture the known fields of an object once at the untrusted-input boundary.
 * Validation and all later reads use the returned descriptor values, never
 * the caller object. Unrelated fields are deliberately ignored: JavaScript has
 * no bounded key-enumeration primitive, so rejecting arbitrary extra keys
 * would let an adversarial object force unbounded result allocation before a
 * size check.
 */
const captureKnownDataRecord = <const Field extends string>(
  input: unknown,
  fields: readonly Field[],
  path: string,
  location: CaptureLocation,
): Result<CapturedDataRecord<Field>, InvalidFolioContentComparisonError> => {
  const { side, blockIndex } = location;
  if (typeof input !== "object" || input === null) {
    return Result.err(
      invalidInput(side, path, "Content comparison records must be objects.", blockIndex),
    );
  }
  const arrayIdentity = captureArrayIdentity(input, path, location);
  if (arrayIdentity.isErr()) return Result.err(arrayIdentity.error);
  if (arrayIdentity.value) {
    return Result.err(
      invalidInput(side, path, "Content comparison records must be objects.", blockIndex),
    );
  }

  const values = new Map<Field, unknown>();
  for (const key of fields) {
    const captured = captureOwnProperty(input, key, `${path}.${key}`, location);
    if (captured.isErr()) return Result.err(captured.error);
    const descriptor = captured.value;
    if (descriptor === null) continue;
    if (!("value" in descriptor)) {
      return Result.err(
        invalidInput(
          side,
          `${path}.${key}`,
          "Content comparison inputs must contain own data properties.",
          blockIndex,
        ),
      );
    }
    values.set(key, descriptor.value);
  }
  return Result.ok(values);
};

/** Capture one dense array after checking its declared length against the cap. */
const captureBoundedDenseArray = (
  input: unknown,
  path: string,
  location: CaptureContext,
  limit: DenseArrayCaptureLimit,
): Result<
  readonly unknown[],
  InvalidFolioContentComparisonError | FolioContentComparisonLimitError
> => {
  const { side, blockIndex, registry } = location;
  const maximum = DENSE_ARRAY_CAPTURE_LIMITS[limit];
  if (typeof input === "object" && input !== null) {
    const retained = registry.denseArrays.get(input);
    if (retained) {
      if (retained.length > maximum) {
        return Result.err(
          limitExceeded({
            input: side,
            limit,
            maximum,
            actual: retained.length,
            ...(blockIndex !== undefined && { blockIndex }),
            field: path,
          }),
        );
      }
      return Result.ok(retained);
    }
  }
  if (typeof input !== "object" || input === null) {
    return Result.err(invalidInput(side, path, "Content comparison value must be an array.", blockIndex));
  }
  const arrayIdentity = captureArrayIdentity(input, path, location);
  if (arrayIdentity.isErr()) return Result.err(arrayIdentity.error);
  if (!arrayIdentity.value) {
    return Result.err(invalidInput(side, path, "Content comparison value must be an array.", blockIndex));
  }
  const capturedLength = captureOwnProperty(input, "length", `${path}.length`, location);
  if (capturedLength.isErr()) return Result.err(capturedLength.error);
  const lengthDescriptor = capturedLength.value;
  if (
    !lengthDescriptor ||
    !("value" in lengthDescriptor) ||
    !Number.isSafeInteger(lengthDescriptor.value) ||
    lengthDescriptor.value < 0
  ) {
    return Result.err(
      invalidInput(side, path, "Content comparison arrays must have a valid data length.", blockIndex),
    );
  }
  const length: number = lengthDescriptor.value;
  if (length > maximum) {
    return Result.err(
      limitExceeded({
        input: side,
        limit,
        maximum,
        actual: length,
        ...(blockIndex !== undefined && { blockIndex }),
        field: path,
      }),
    );
  }
  const values = new Array<unknown>(length);
  for (let index = 0; index < length; index++) {
    const key = String(index);
    const captured = captureOwnProperty(input, key, `${path}[${key}]`, location);
    if (captured.isErr()) return Result.err(captured.error);
    const descriptor = captured.value;
    if (!descriptor || !("value" in descriptor)) {
      return Result.err(
        invalidInput(
          side,
          `${path}[${key}]`,
          "Content comparison arrays must contain own data items.",
          blockIndex,
        ),
      );
    }
    values[index] = descriptor.value;
  }
  const captured = Object.freeze(values);
  registry.denseArrays.set(input, captured);
  return Result.ok(captured);
};

const recordHasExactly = <Field extends string>(
  record: CapturedDataRecord<Field>,
  fields: readonly Field[],
): boolean =>
  record.size === fields.length && fields.every((field) => record.has(field));

type TableCellRectangle = {
  blockIndex: number;
  left: number;
  right: number;
  top: number;
  bottom: number;
};

type TableRowCellPlacement = {
  cellIndex: number;
  gridEnd: number;
};

const overlappingTableCellBlockIndex = (cells: readonly TableCellRectangle[]): number | null => {
  if (cells.length < 2) return null;
  const columns = [...new Set(cells.flatMap(({ left, right }) => [left, right]))].toSorted(
    (left, right) => left - right,
  );
  if (columns.length < 2) return null;
  const columnIndex = new Map(columns.map((column, index) => [column, index] as const));
  const intervalCount = columns.length - 1;
  const maximum = new Int32Array(intervalCount * 4);
  const pending = new Int32Array(intervalCount * 4);
  const add = (
    node: number,
    nodeLeft: number,
    nodeRight: number,
    rangeLeft: number,
    rangeRight: number,
    amount: number,
  ): void => {
    if (rangeLeft <= nodeLeft && nodeRight <= rangeRight) {
      maximum[node] = (maximum[node] ?? 0) + amount;
      pending[node] = (pending[node] ?? 0) + amount;
      return;
    }
    const middle = nodeLeft + Math.floor((nodeRight - nodeLeft) / 2);
    if (rangeLeft <= middle) {
      add(node * 2, nodeLeft, middle, rangeLeft, rangeRight, amount);
    }
    if (rangeRight > middle) {
      add(node * 2 + 1, middle + 1, nodeRight, rangeLeft, rangeRight, amount);
    }
    maximum[node] =
      (pending[node] ?? 0) + Math.max(maximum[node * 2] ?? 0, maximum[node * 2 + 1] ?? 0);
  };
  const events = cells.flatMap((cell) => [
    { row: cell.top, amount: 1, cell },
    { row: cell.bottom, amount: -1, cell },
  ]);
  events.sort(
    (left, right) =>
      left.row - right.row ||
      left.amount - right.amount ||
      left.cell.left - right.cell.left ||
      left.cell.blockIndex - right.cell.blockIndex,
  );
  for (const { amount, cell } of events) {
    const left = columnIndex.get(cell.left);
    const right = columnIndex.get(cell.right);
    if (left === undefined || right === undefined || left >= right) {
      return panic("A validated table cell has no coordinate-compression interval");
    }
    add(1, 0, intervalCount - 1, left, right - 1, amount);
    if (amount > 0 && (maximum[1] ?? 0) > 1) {
      return cell.blockIndex;
    }
  }
  return null;
};

const chargeAttributeString = ({
  value,
  side,
  blockIndex,
  field,
  usage,
}: {
  value: string | null | undefined;
  side: "base" | "revised";
  blockIndex: number;
  field: string;
  usage: SnapshotResourceUsage;
}): FolioContentComparisonLimitError | null => {
  if (value === undefined || value === null) {
    return null;
  }
  if (value.length > FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnits) {
    return limitExceeded({
      input: side,
      limit: "attributeCodeUnits",
      maximum: FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnits,
      actual: value.length,
      blockIndex,
      field,
    });
  }
  usage.attributeCodeUnits += value.length;
  if (usage.attributeCodeUnits > FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnitsPerSnapshot) {
    return limitExceeded({
      input: side,
      limit: "attributeCodeUnitsPerSnapshot",
      maximum: FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnitsPerSnapshot,
      actual: usage.attributeCodeUnits,
      blockIndex,
      field,
    });
  }
  return null;
};

const chargeCapturedString = (
  value: string,
  path: string,
  context: CaptureContext,
): FolioContentComparisonLimitError | null =>
  chargeAttributeString({
    value,
    side: context.side,
    blockIndex: context.blockIndex,
    field: path,
    usage: context.usage,
  });

const claimPropertyNode = (
  path: string,
  context: CaptureContext,
): FolioContentComparisonLimitError | null => {
  context.usage.propertyNodes++;
  return context.usage.propertyNodes > FOLIO_CONTENT_COMPARISON_LIMITS.propertyNodesPerSnapshot
    ? limitExceeded({
        input: context.side,
        limit: "propertyNodesPerSnapshot",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.propertyNodesPerSnapshot,
        actual: context.usage.propertyNodes,
        blockIndex: context.blockIndex,
        field: path,
      })
    : null;
};

const capturePropertyValue = (
  input: unknown,
  path: string,
  context: CaptureContext,
  depth: number,
): Result<FolioContentPropertyValue, FolioContentComparisonError> => {
  if (depth > FOLIO_CONTENT_COMPARISON_LIMITS.propertyDepth) {
    return Result.err(
      limitExceeded({
        input: context.side,
        limit: "propertyDepth",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.propertyDepth,
        actual: depth,
        blockIndex: context.blockIndex,
        field: path,
      }),
    );
  }
  const nodeError = claimPropertyNode(path, context);
  if (nodeError) return Result.err(nodeError);
  if (input === null || typeof input === "boolean") return Result.ok(input);
  if (typeof input === "number") {
    if (!Number.isFinite(input)) {
      return Result.err(
        invalidInput(context.side, path, "Property numbers must be finite.", context.blockIndex),
      );
    }
    return Result.ok(Object.is(input, -0) ? 0 : input);
  }
  if (typeof input === "string") {
    const limit = chargeCapturedString(input, path, context);
    return limit ? Result.err(limit) : Result.ok(input);
  }
  const record = captureKnownDataRecord(
    input,
    [
      ...Object.values(FOLIO_CONTENT_PROPERTY_ARRAY_FIELD_DESCRIPTORS).map(({ field }) => field),
      ...Object.values(FOLIO_CONTENT_PROPERTY_OBJECT_FIELD_DESCRIPTORS).map(({ field }) => field),
    ],
    path,
    context,
  );
  if (record.isErr()) return Result.err(record.error);
  const type = record.value.get("type");
  if (type === "array") {
    const fields = Object.values(FOLIO_CONTENT_PROPERTY_ARRAY_FIELD_DESCRIPTORS).map(
      ({ field }) => field,
    );
    if (!recordHasExactly(record.value, fields)) {
      return Result.err(
        invalidInput(
          context.side,
          path,
          "Property arrays require exactly type and items.",
          context.blockIndex,
        ),
      );
    }
    const items = captureBoundedDenseArray(
      record.value.get("items"),
      `${path}.items`,
      context,
      "propertyEntriesPerContainer",
    );
    if (items.isErr()) return Result.err(items.error);
    const capturedItems: FolioContentPropertyValue[] = [];
    for (let index = 0; index < items.value.length; index++) {
      const itemPath = `${path}.items[${String(index)}]`;
      const captured = capturePropertyValue(items.value[index], itemPath, context, depth + 1);
      if (captured.isErr()) return Result.err(captured.error);
      capturedItems.push(captured.value);
    }
    return Result.ok(Object.freeze({ type: "array", items: Object.freeze(capturedItems) }));
  }
  if (type === "object") {
    const fields = Object.values(FOLIO_CONTENT_PROPERTY_OBJECT_FIELD_DESCRIPTORS).map(
      ({ field }) => field,
    );
    if (!recordHasExactly(record.value, fields)) {
      return Result.err(
        invalidInput(
          context.side,
          path,
          "Property objects require exactly type and entries.",
          context.blockIndex,
        ),
      );
    }
    const captured = capturePropertySet(
      record.value.get("entries"),
      `${path}.entries`,
      context,
      depth + 1,
    );
    return captured.isErr()
      ? Result.err(captured.error)
      : Result.ok(Object.freeze({ type: "object", entries: captured.value }));
  }
  return Result.err(
    invalidInput(
      context.side,
      `${path}.type`,
      "Property containers must declare type object or array.",
      context.blockIndex,
    ),
  );
};

const capturePropertySet = (
  input: unknown,
  path: string,
  context: CaptureContext,
  depth = 0,
): Result<FolioContentPropertySet, FolioContentComparisonError> => {
  const entries = captureBoundedDenseArray(
    input,
    path,
    context,
    "propertyEntriesPerContainer",
  );
  if (entries.isErr()) return Result.err(entries.error);
  const keys = new Set<string>();
  const captured: { key: string; value: FolioContentPropertyValue }[] = [];
  for (let index = 0; index < entries.value.length; index++) {
    const entryPath = `${path}[${String(index)}]`;
    const entry = captureKnownDataRecord(
      entries.value[index],
      Object.values(FOLIO_CONTENT_PROPERTY_ENTRY_FIELD_DESCRIPTORS).map(({ field }) => field),
      entryPath,
      context,
    );
    if (entry.isErr()) return Result.err(entry.error);
    const key = entry.value.get("key");
    if (typeof key !== "string" || key.length === 0 || keys.has(key)) {
      return Result.err(
        invalidInput(
          context.side,
          `${entryPath}.key`,
          "Property keys must be unique non-empty strings.",
          context.blockIndex,
        ),
      );
    }
    const keyLimit = chargeCapturedString(key, `${entryPath}.key`, context);
    if (keyLimit) return Result.err(keyLimit);
    keys.add(key);
    const value = entry.value.get("value");
    if (value === undefined) {
      return Result.err(
        invalidInput(
          context.side,
          `${entryPath}.value`,
          "Property values cannot be undefined.",
          context.blockIndex,
        ),
      );
    }
    const capturedValue = capturePropertyValue(
      value,
      `${entryPath}.value`,
      context,
      depth,
    );
    if (capturedValue.isErr()) return Result.err(capturedValue.error);
    captured.push(Object.freeze({ key, value: capturedValue.value }));
  }
  captured.sort(({ key: left }, { key: right }) => (left < right ? -1 : left > right ? 1 : 0));
  return Result.ok(Object.freeze(captured));
};

const EMPTY_PROPERTY_SET = Object.freeze([]) satisfies FolioContentPropertySet;

const captureParagraphFormatting = (
  input: unknown,
  path: string,
  context: CaptureContext,
): Result<FolioContentParagraphFormatting, FolioContentComparisonError> => {
  const record = captureKnownDataRecord(
    input,
    Object.values(FOLIO_CONTENT_PARAGRAPH_FORMATTING_FIELD_DESCRIPTORS).map(
      ({ field }) => field,
    ),
    path,
    context,
  );
  if (record.isErr()) return Result.err(record.error);
  let effective = EMPTY_PROPERTY_SET;
  let authored = EMPTY_PROPERTY_SET;
  for (const descriptor of Object.values(
    FOLIO_CONTENT_PARAGRAPH_FORMATTING_FIELD_DESCRIPTORS,
  )) {
    const value = record.value.get(descriptor.field);
    if (value === undefined) continue;
    const captured = capturePropertySet(value, `${path}.${descriptor.field}`, context);
    if (captured.isErr()) return Result.err(captured.error);
    if (descriptor.role === "effective-format") {
      effective = captured.value;
    } else {
      authored = captured.value;
    }
  }
  return Result.ok(Object.freeze({ effective, authored }));
};

const captureContentRun = (
  input: unknown,
  path: string,
  context: CaptureContext,
): Result<FolioContentRun, FolioContentComparisonError> => {
  const record = captureKnownDataRecord(
    input,
    Object.values(FOLIO_CONTENT_RUN_FIELD_DESCRIPTORS).map(({ field }) => field),
    path,
    context,
  );
  if (record.isErr()) return Result.err(record.error);
  const captured: {
    text: string;
    effectiveFormatting: FolioContentPropertySet;
    authoredFormatting: FolioContentPropertySet;
  } = { text: "", effectiveFormatting: Object.freeze([]), authoredFormatting: Object.freeze([]) };
  for (const descriptor of Object.values(FOLIO_CONTENT_RUN_FIELD_DESCRIPTORS)) {
    const fieldPath = `${path}.${descriptor.field}`;
    const value = record.value.get(descriptor.field);
    switch (descriptor.capture) {
      case "required-scalar":
        if (typeof value !== "string") {
          return Result.err(
            invalidInput(context.side, fieldPath, "Preview-run text must be a string.", context.blockIndex),
          );
        }
        captured.text = value;
        break;
      case "properties": {
        if (value === undefined) break;
        const formatting = capturePropertySet(value, fieldPath, context);
        if (formatting.isErr()) return Result.err(formatting.error);
        Reflect.set(captured, descriptor.field, formatting.value);
        break;
      }
      default: {
        const unreachable: never = descriptor;
        return panic("Unhandled preview-run capture", { descriptor: unreachable });
      }
    }
  }
  return Result.ok(Object.freeze(captured));
};

const captureTableLocation = (
  input: unknown,
  path: string,
  context: CaptureContext,
): Result<NonNullable<FolioContentBlock["table"]>, FolioContentComparisonError> => {
  const record = captureKnownDataRecord(
    input,
    Object.values(FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS).map(({ field }) => field),
    path,
    context,
  );
  if (record.isErr()) return Result.err(record.error);
  const captured: {
    outerTableIdentity?: FolioContentBlock["identity"];
    tableIdentity?: FolioContentBlock["identity"];
    rowIdentity?: FolioContentBlock["identity"];
    cellIdentity?: FolioContentBlock["identity"];
    outerTableIndex: number;
    tableIndex: number;
    rowIndex: number;
    cellIndex: number;
    gridColumnIndex: number;
    columnSpan: number;
    rowSpan: number;
    paragraphIndex: number;
  } = {
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex: 0,
    cellIndex: 0,
    gridColumnIndex: 0,
    columnSpan: 0,
    rowSpan: 0,
    paragraphIndex: 0,
  };
  for (const descriptor of Object.values(FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS)) {
    const fieldPath = `${path}.${descriptor.field}`;
    const value = record.value.get(descriptor.field);
    if (descriptor.validation === "identity") {
      const identity = captureContentIdentity(value, fieldPath, context);
      if (identity.isErr()) return Result.err(identity.error);
      Reflect.set(captured, descriptor.field, identity.value);
      continue;
    }
    const minimum = descriptor.validation === "span" ? 1 : 0;
    if (!isFiniteInteger(value) || value < minimum) {
      return Result.err(
        invalidInput(context.side, fieldPath, "Table coordinates must be supported integers.", context.blockIndex),
      );
    }
    captured[descriptor.field] = value;
  }
  if (
    captured.outerTableIdentity === undefined ||
    captured.tableIdentity === undefined ||
    captured.rowIdentity === undefined ||
    captured.cellIdentity === undefined ||
    !Number.isSafeInteger(captured.gridColumnIndex + captured.columnSpan) ||
    !Number.isSafeInteger(captured.rowIndex + captured.rowSpan)
  ) {
    return Result.err(
      invalidInput(context.side, path, "Table cell bounds must be safe integers.", context.blockIndex),
    );
  }
  return Result.ok(
    Object.freeze({
      outerTableIdentity: captured.outerTableIdentity,
      tableIdentity: captured.tableIdentity,
      rowIdentity: captured.rowIdentity,
      cellIdentity: captured.cellIdentity,
      outerTableIndex: captured.outerTableIndex,
      tableIndex: captured.tableIndex,
      rowIndex: captured.rowIndex,
      cellIndex: captured.cellIndex,
      gridColumnIndex: captured.gridColumnIndex,
      columnSpan: captured.columnSpan,
      rowSpan: captured.rowSpan,
      paragraphIndex: captured.paragraphIndex,
    }),
  );
};

const captureContainerPath = (
  input: unknown,
  path: string,
  context: CaptureContext,
): Result<NonNullable<FolioContentBlock["containerPath"]>, FolioContentComparisonError> => {
  const entries = captureBoundedDenseArray(input, path, context, "containerDepth");
  if (entries.isErr()) return Result.err(entries.error);
  context.usage.containerEntries += entries.value.length;
  if (context.usage.containerEntries > FOLIO_CONTENT_COMPARISON_LIMITS.containerEntriesPerSnapshot) {
    return Result.err(
      limitExceeded({
        input: context.side,
        limit: "containerEntriesPerSnapshot",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.containerEntriesPerSnapshot,
        actual: context.usage.containerEntries,
        blockIndex: context.blockIndex,
        field: path,
      }),
    );
  }
  const captured: { kind: string; identity: FolioContentBlock["identity"] }[] = [];
  for (let pathIndex = 0; pathIndex < entries.value.length; pathIndex++) {
    const entryPath = `${path}[${String(pathIndex)}]`;
    const item = captureKnownDataRecord(
      entries.value[pathIndex],
      Object.values(FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS).map(({ field }) => field),
      entryPath,
      context,
    );
    if (item.isErr()) return Result.err(item.error);
    const entry: { kind: string; identity?: FolioContentBlock["identity"] } = { kind: "" };
    for (const descriptor of Object.values(FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS)) {
      const fieldPath = `${entryPath}.${descriptor.field}`;
      const value = item.value.get(descriptor.field);
      if (descriptor.validation === "identity") {
        const identity = captureContentIdentity(value, fieldPath, context);
        if (identity.isErr()) return Result.err(identity.error);
        entry.identity = identity.value;
        continue;
      }
      if (typeof value !== "string" || value.length === 0) {
        return Result.err(
          invalidInput(context.side, fieldPath, "Container fields must be non-empty strings.", context.blockIndex),
        );
      }
      const limit = chargeCapturedString(value, fieldPath, context);
      if (limit) return Result.err(limit);
      entry[descriptor.field] = value;
    }
    if (entry.identity === undefined) {
      return panic("The total container descriptor did not capture identity");
    }
    captured.push(Object.freeze({ kind: entry.kind, identity: entry.identity }));
  }
  return Result.ok(Object.freeze(captured));
};

const captureStructuralBoundaries = (
  input: unknown,
  blockTextLength: number,
  path: string,
  context: CaptureContext,
): Result<NonNullable<FolioContentBlock["structuralBoundaries"]>, FolioContentComparisonError> => {
  const entries = captureBoundedDenseArray(
    input,
    path,
    context,
    "structuralBoundariesPerBlock",
  );
  if (entries.isErr()) return Result.err(entries.error);
  context.usage.structuralBoundaries += entries.value.length;
  if (
    context.usage.structuralBoundaries >
    FOLIO_CONTENT_COMPARISON_LIMITS.structuralBoundariesPerSnapshot
  ) {
    return Result.err(
      limitExceeded({
        input: context.side,
        limit: "structuralBoundariesPerSnapshot",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.structuralBoundariesPerSnapshot,
        actual: context.usage.structuralBoundaries,
        blockIndex: context.blockIndex,
        field: path,
      }),
    );
  }
  const captured: { type: "pageBreak"; offset: number; clear?: "all" | "left" | "right" | "none" }[] = [];
  let priorOffset = -1;
  for (let boundaryIndex = 0; boundaryIndex < entries.value.length; boundaryIndex++) {
    const boundaryPath = `${path}[${String(boundaryIndex)}]`;
    const item = captureKnownDataRecord(
      entries.value[boundaryIndex],
      Object.values(FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS).map(
        ({ field }) => field,
      ),
      boundaryPath,
      context,
    );
    if (item.isErr()) return Result.err(item.error);
    const boundary: { type: "pageBreak"; offset: number; clear?: "all" | "left" | "right" | "none" } = {
      type: "pageBreak",
      offset: 0,
    };
    for (const descriptor of Object.values(FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS)) {
      const fieldPath = `${boundaryPath}.${descriptor.field}`;
      const value = item.value.get(descriptor.field);
      if (descriptor.validation === "page-break") {
        if (value !== "pageBreak") {
          return Result.err(invalidInput(context.side, fieldPath, "Only page-break boundaries are supported.", context.blockIndex));
        }
      } else if (descriptor.validation === "offset") {
        if (!isFiniteInteger(value) || value < priorOffset || value > blockTextLength) {
          return Result.err(invalidInput(context.side, fieldPath, "Boundary offsets must be monotone UTF-16 positions.", context.blockIndex));
        }
      } else if (
        value !== undefined &&
        value !== "all" &&
        value !== "left" &&
        value !== "right" &&
        value !== "none"
      ) {
        return Result.err(invalidInput(context.side, fieldPath, "Page-break clear value is invalid.", context.blockIndex));
      }
      if (value !== undefined) Reflect.set(boundary, descriptor.field, value);
    }
    priorOffset = boundary.offset;
    captured.push(Object.freeze(boundary));
  }
  return Result.ok(Object.freeze(captured));
};

const captureRuns = (
  input: unknown,
  blockText: string,
  path: string,
  context: CaptureContext,
): Result<NonNullable<FolioContentBlock["runs"]>, FolioContentComparisonError> => {
  const entries = captureBoundedDenseArray(input, path, context, "runsPerBlock");
  if (entries.isErr()) return Result.err(entries.error);
  context.usage.runs += entries.value.length;
  if (context.usage.runs > FOLIO_CONTENT_COMPARISON_LIMITS.runsPerSnapshot) {
    return Result.err(
      limitExceeded({
        input: context.side,
        limit: "runsPerSnapshot",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.runsPerSnapshot,
        actual: context.usage.runs,
        blockIndex: context.blockIndex,
        field: path,
      }),
    );
  }
  if (entries.value.length === 0) {
    return Result.ok(Object.freeze([]));
  }
  const captured: FolioContentRun[] = [];
  let offset = 0;
  for (let runIndex = 0; runIndex < entries.value.length; runIndex++) {
    const runPath = `${path}[${String(runIndex)}]`;
    const run = captureContentRun(entries.value[runIndex], runPath, context);
    if (run.isErr()) return Result.err(run.error);
    if (!blockText.startsWith(run.value.text, offset)) {
      return Result.err(
        invalidInput(context.side, path, "Preview-run text must reconstruct block text exactly.", context.blockIndex),
      );
    }
    offset += run.value.text.length;
    captured.push(run.value);
  }
  if (offset !== blockText.length) {
    return Result.err(
      invalidInput(context.side, path, "Preview-run text must reconstruct block text exactly.", context.blockIndex),
    );
  }
  return Result.ok(Object.freeze(captured));
};

const captureContentIdentity = (
  input: unknown,
  path: string,
  context: CaptureContext,
): Result<FolioContentBlock["identity"], FolioContentComparisonError> => {
  const record = captureKnownDataRecord(
    input,
    Object.values(FOLIO_CONTENT_IDENTITY_FIELD_DESCRIPTORS).map(({ field }) => field),
    path,
    context,
  );
  if (record.isErr()) return Result.err(record.error);
  const type = record.value.get("type");
  const id = record.value.get("id");
  if (!IDENTITY_SEMANTICS.has(type) || typeof id !== "string" || id.length === 0) {
    return Result.err(
      invalidInput(
        context.side,
        path,
        "Block identity requires a supported type and a non-empty id.",
        context.blockIndex,
      ),
    );
  }
  const limit = chargeCapturedString(id, `${path}.id`, context);
  if (limit) return Result.err(limit);
  switch (type) {
    case "authoritative":
      return Result.ok(Object.freeze({ type: "authoritative", id }));
    case "persistent-hint":
      return Result.ok(Object.freeze({ type: "persistent-hint", id }));
    case "positional":
      return Result.ok(Object.freeze({ type: "positional", id }));
    default:
      return panic("Validated content identity has an unsupported type");
  }
};

const captureContentBlock = (
  input: unknown,
  side: "base" | "revised",
  blockIndex: number,
  usage: SnapshotResourceUsage,
  registry: CaptureRegistry,
): Result<FolioContentBlock, FolioContentComparisonError> => {
  const path = `blocks[${String(blockIndex)}]`;
  const context = { side, blockIndex, usage, registry } satisfies CaptureContext;
  const record = captureKnownDataRecord(
    input,
    Object.values(FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS).map(({ field }) => field),
    path,
    context,
  );
  if (record.isErr()) return Result.err(record.error);
  const captured: {
    identity?: FolioContentBlock["identity"];
    kind?: string;
    text?: string;
    blockProperties: FolioContentPropertySet;
    paragraphFormatting: FolioContentParagraphFormatting;
    runs: readonly FolioContentRun[];
    structuralBoundaries: FolioContentBlock["structuralBoundaries"];
    table?: FolioContentBlock["table"];
    containerPath: FolioContentBlock["containerPath"];
  } = {
    blockProperties: Object.freeze([]),
    paragraphFormatting: Object.freeze({
      effective: EMPTY_PROPERTY_SET,
      authored: EMPTY_PROPERTY_SET,
    }),
    runs: Object.freeze([]),
    structuralBoundaries: Object.freeze([]),
    containerPath: Object.freeze([]),
  };
  for (const descriptor of Object.values(FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS)) {
    const fieldPath = `${path}.${descriptor.field}`;
    const value = record.value.get(descriptor.field);
    switch (descriptor.capture) {
      case "identity": {
        const identity = captureContentIdentity(value, fieldPath, context);
        if (identity.isErr()) return Result.err(identity.error);
        captured.identity = identity.value;
        break;
      }
      case "required-scalar": {
        if (
          typeof value !== "string" ||
          (descriptor.validation === "nonempty-string" && value.length === 0)
        ) {
          return Result.err(
            invalidInput(side, fieldPath, "Content block field is invalid.", blockIndex),
          );
        }
        if (descriptor.validation === "text") {
          if (value.length > FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits) {
            return Result.err(
              limitExceeded({
                input: side,
                limit: "blockCodeUnits",
                maximum: FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits,
                actual: value.length,
                blockIndex,
                field: fieldPath,
              }),
            );
          }
          usage.textCodeUnits += value.length;
          if (usage.textCodeUnits > FOLIO_CONTENT_COMPARISON_LIMITS.textCodeUnitsPerSnapshot) {
            return Result.err(
              limitExceeded({
                input: side,
                limit: "textCodeUnitsPerSnapshot",
                maximum: FOLIO_CONTENT_COMPARISON_LIMITS.textCodeUnitsPerSnapshot,
                actual: usage.textCodeUnits,
                blockIndex,
                field: fieldPath,
              }),
            );
          }
        } else {
          const limit = chargeCapturedString(value, fieldPath, context);
          if (limit) return Result.err(limit);
        }
        Reflect.set(captured, descriptor.field, value);
        break;
      }
      case "properties": {
        if (value === undefined) break;
        const properties = capturePropertySet(value, fieldPath, context);
        if (properties.isErr()) return Result.err(properties.error);
        Reflect.set(captured, descriptor.field, properties.value);
        break;
      }
      case "paragraph-formatting": {
        if (value === undefined) break;
        const formatting = captureParagraphFormatting(value, fieldPath, context);
        if (formatting.isErr()) return Result.err(formatting.error);
        captured.paragraphFormatting = formatting.value;
        break;
      }
      case "runs": {
        if (value === undefined) break;
        if (captured.text === undefined) {
          return panic("The total block descriptor must capture text before preview runs");
        }
        const runs = captureRuns(value, captured.text, fieldPath, context);
        if (runs.isErr()) return Result.err(runs.error);
        Reflect.set(captured, descriptor.field, runs.value);
        break;
      }
      case "boundaries": {
        if (value === undefined) break;
        if (captured.text === undefined) {
          return panic("The total block descriptor must capture text before structural boundaries");
        }
        const boundaries = captureStructuralBoundaries(value, captured.text.length, fieldPath, context);
        if (boundaries.isErr()) return Result.err(boundaries.error);
        Reflect.set(captured, descriptor.field, boundaries.value);
        break;
      }
      case "table": {
        if (value === undefined) break;
        const table = captureTableLocation(value, fieldPath, context);
        if (table.isErr()) return Result.err(table.error);
        Reflect.set(captured, descriptor.field, table.value);
        break;
      }
      case "container": {
        if (value === undefined) break;
        const container = captureContainerPath(value, fieldPath, context);
        if (container.isErr()) return Result.err(container.error);
        Reflect.set(captured, descriptor.field, container.value);
        break;
      }
      default: {
        const unreachable: never = descriptor;
        return panic("Unhandled content block capture", { descriptor: unreachable });
      }
    }
  }
  if (captured.identity === undefined || captured.kind === undefined || captured.text === undefined) {
    return panic("The total block descriptor did not capture required neutral fields");
  }
  return Result.ok(
    Object.freeze({
      identity: captured.identity,
      kind: captured.kind,
      text: captured.text,
      blockProperties: captured.blockProperties,
      paragraphFormatting: captured.paragraphFormatting,
      runs: captured.runs,
      structuralBoundaries: captured.structuralBoundaries,
      ...(captured.table !== undefined && { table: captured.table }),
      containerPath: captured.containerPath,
    }),
  );
};

const captureValidatedSnapshotInto = (
  snapshot: unknown,
  side: "base" | "revised",
  usage: SnapshotResourceUsage,
  capturedBlocks: FolioContentBlock[],
  registry: CaptureRegistry,
): FolioContentComparisonError | null => {
  const snapshotContext = { side, usage, registry } satisfies CaptureContext;
  const record = captureKnownDataRecord(
    snapshot,
    Object.values(FOLIO_CONTENT_SNAPSHOT_FIELD_DESCRIPTORS).map(({ field }) => field),
    side,
    snapshotContext,
  );
  if (record.isErr()) return record.error;
  const blocks = captureBoundedDenseArray(
    record.value.get("blocks"),
    `${side}.blocks`,
    snapshotContext,
    "blocksPerSnapshot",
  );
  if (blocks.isErr()) return blocks.error;
  const inputBlocks = blocks.value;

  usage.blocks = inputBlocks.length;
  const ids = new Set<string>();
  const lastCoordinateByTable = new Map<string, readonly [number, number, number]>();
  const placementByRow = new Map<string, TableRowCellPlacement>();
  const geometryByCell = new Map<string, readonly [number, number, number]>();
  const cellsByTable = new Map<string, TableCellRectangle[]>();
  const outerTableByTableIndex = new Map<number, number>();
  const structuralIdentityIndexes = {
    outerTable: {
      identityByCoordinate: new Map<string, string>(),
      coordinateByIdentity: new Map<string, string>(),
    },
    table: {
      identityByCoordinate: new Map<string, string>(),
      coordinateByIdentity: new Map<string, string>(),
    },
    row: {
      identityByCoordinate: new Map<string, string>(),
      coordinateByIdentity: new Map<string, string>(),
    },
    cell: {
      identityByCoordinate: new Map<string, string>(),
      coordinateByIdentity: new Map<string, string>(),
    },
  };
  const claimStructuralIdentity = ({
    level,
    coordinate,
    identity,
    blockIndex,
  }: {
    level: keyof typeof structuralIdentityIndexes;
    coordinate: string;
    identity: FolioContentIdentity;
    blockIndex: number;
  }): InvalidFolioContentComparisonError | null => {
    const index = structuralIdentityIndexes[level];
    const identityKey = JSON.stringify([identity.type, identity.id]);
    const priorIdentity = index.identityByCoordinate.get(coordinate);
    const priorCoordinate = index.coordinateByIdentity.get(identityKey);
    if (
      (priorIdentity !== undefined && priorIdentity !== identityKey) ||
      (priorCoordinate !== undefined && priorCoordinate !== coordinate)
    ) {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].table.${level}Identity`,
        "A structural identity must name exactly one table coordinate.",
        blockIndex,
      );
    }
    index.identityByCoordinate.set(coordinate, identityKey);
    index.coordinateByIdentity.set(identityKey, coordinate);
    return null;
  };
  let lastOuterTableIndex = -1;
  let activeOuterTableIndex: number | null = null;
  for (let blockIndex = 0; blockIndex < inputBlocks.length; blockIndex++) {
    const capturedBlock = captureContentBlock(
      inputBlocks[blockIndex],
      side,
      blockIndex,
      usage,
      registry,
    );
    if (capturedBlock.isErr()) return capturedBlock.error;
    const block = capturedBlock.value;
    if (ids.has(block.identity.id)) {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].identity.id`,
        "Content block ids must be unique within a snapshot.",
        blockIndex,
      );
    }
    ids.add(block.identity.id);
    if (block.table !== undefined) {
      const table = block.table;
      if (table.tableIndex < table.outerTableIndex) {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].table.tableIndex`,
          "An inner table cannot precede its outer table in document order.",
          blockIndex,
        );
      }
      const knownOuterTableIndex = outerTableByTableIndex.get(table.tableIndex);
      if (knownOuterTableIndex !== undefined && knownOuterTableIndex !== table.outerTableIndex) {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].table.tableIndex`,
          "A table index must identify only one outer table.",
          blockIndex,
        );
      }
      outerTableByTableIndex.set(table.tableIndex, table.outerTableIndex);
      if (
        table.outerTableIndex < lastOuterTableIndex ||
        (table.outerTableIndex === lastOuterTableIndex &&
          activeOuterTableIndex !== table.outerTableIndex)
      ) {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].table`,
          "Each outer table must occupy one contiguous position in document order.",
          blockIndex,
        );
      }
      lastOuterTableIndex = table.outerTableIndex;
      activeOuterTableIndex = table.outerTableIndex;
      const tableKey = `${String(table.outerTableIndex)}:${String(table.tableIndex)}`;
      const rowKey = `${tableKey}:${String(table.rowIndex)}`;
      const cellKey = `${tableKey}:${String(table.rowIndex)}:${String(table.cellIndex)}`;
      const structuralIdentityError =
        claimStructuralIdentity({
          level: "outerTable",
          coordinate: String(table.outerTableIndex),
          identity: table.outerTableIdentity,
          blockIndex,
        }) ??
        claimStructuralIdentity({
          level: "table",
          coordinate: tableKey,
          identity: table.tableIdentity,
          blockIndex,
        }) ??
        claimStructuralIdentity({
          level: "row",
          coordinate: `${tableKey}:${String(table.rowIndex)}`,
          identity: table.rowIdentity,
          blockIndex,
        }) ??
        claimStructuralIdentity({
          level: "cell",
          coordinate: cellKey,
          identity: table.cellIdentity,
          blockIndex,
        });
      if (structuralIdentityError) return structuralIdentityError;
      const geometry = [table.gridColumnIndex, table.columnSpan, table.rowSpan] as const;
      const priorGeometry = geometryByCell.get(cellKey);
      if (
        priorGeometry !== undefined &&
        (geometry[0] !== priorGeometry[0] ||
          geometry[1] !== priorGeometry[1] ||
          geometry[2] !== priorGeometry[2])
      ) {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].table`,
          "Every block in one physical table cell must carry the same geometry.",
          blockIndex,
        );
      }
      geometryByCell.set(cellKey, geometry);
      if (priorGeometry === undefined) {
        const priorPlacement = placementByRow.get(rowKey);
        if (
          (priorPlacement === undefined && table.cellIndex !== 0) ||
          (priorPlacement !== undefined &&
            (table.cellIndex !== priorPlacement.cellIndex + 1 ||
              table.gridColumnIndex < priorPlacement.gridEnd))
        ) {
          return invalidInput(
            side,
            `blocks[${String(blockIndex)}].table`,
            "Physical cell order must agree with non-overlapping logical grid order.",
            blockIndex,
          );
        }
        placementByRow.set(rowKey, {
          cellIndex: table.cellIndex,
          gridEnd: table.gridColumnIndex + table.columnSpan,
        });
        const cells = cellsByTable.get(tableKey);
        const rectangle = {
          blockIndex,
          left: table.gridColumnIndex,
          right: table.gridColumnIndex + table.columnSpan,
          top: table.rowIndex,
          bottom: table.rowIndex + table.rowSpan,
        };
        if (cells) {
          cells.push(rectangle);
        } else {
          cellsByTable.set(tableKey, [rectangle]);
        }
      }
      const coordinate = [table.rowIndex, table.cellIndex, table.paragraphIndex] as const;
      const previous = lastCoordinateByTable.get(tableKey);
      if (
        previous &&
        (coordinate[0] < previous[0] ||
          (coordinate[0] === previous[0] && coordinate[1] < previous[1]) ||
          (coordinate[0] === previous[0] &&
            coordinate[1] === previous[1] &&
            coordinate[2] <= previous[2]))
      ) {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].table`,
          "Blocks in one table must use row-major coordinates.",
          blockIndex,
        );
      }
      lastCoordinateByTable.set(tableKey, coordinate);
    } else {
      activeOuterTableIndex = null;
    }
    capturedBlocks.push(block);
  }
  for (const cells of cellsByTable.values()) {
    const overlappingBlockIndex = overlappingTableCellBlockIndex(cells);
    if (overlappingBlockIndex !== null) {
      return invalidInput(
        side,
        `blocks[${String(overlappingBlockIndex)}].table`,
        "Physical table cells must not overlap.",
        overlappingBlockIndex,
      );
    }
  }
  return null;
};

type CapturedSnapshot = {
  snapshot: FolioContentSnapshot;
  usage: SnapshotResourceUsage;
};

const captureContentSnapshot = (
  snapshot: unknown,
  side: "base" | "revised",
  registry: CaptureRegistry,
): Result<CapturedSnapshot, FolioContentComparisonError> => {
  const usage = emptySnapshotResourceUsage();
  const blocks: FolioContentBlock[] = [];
  const error = captureValidatedSnapshotInto(snapshot, side, usage, blocks, registry);
  if (error) return Result.err(error);
  return Result.ok({
    snapshot: Object.freeze({ blocks: Object.freeze(blocks) }),
    usage,
  });
};

const aggregateSnapshotLimitError = (
  side: "base" | "revised",
  retained: SnapshotResourceUsage,
  incoming: SnapshotResourceUsage,
): FolioContentComparisonLimitError | null => {
  for (const { resource, limit } of SNAPSHOT_RESOURCE_DESCRIPTORS) {
    const maximum = FOLIO_CONTENT_COMPARISON_LIMITS[limit];
    const actual = retained[resource] + incoming[resource];
    if (actual > maximum) {
      return limitExceeded({ input: side, limit, maximum, actual });
    }
  }
  return null;
};

const claimSnapshotPairResources = (
  resourceUsage: ContentComparisonResourceUsage,
  base: SnapshotResourceUsage,
  revised: SnapshotResourceUsage,
): FolioContentComparisonLimitError | null => {
  const baseError = aggregateSnapshotLimitError("base", resourceUsage.base, base);
  if (baseError) return baseError;
  const revisedError = aggregateSnapshotLimitError("revised", resourceUsage.revised, revised);
  if (revisedError) return revisedError;
  for (const { resource } of SNAPSHOT_RESOURCE_DESCRIPTORS) {
    resourceUsage.base[resource] += base[resource];
    resourceUsage.revised[resource] += revised[resource];
  }
  return null;
};

const withTextOffsets = (
  segments: readonly WordDiffSegment[],
  baseStart = 0,
  revisedStart = 0,
): FolioContentTextSegment[] => {
  const positioned: FolioContentTextSegment[] = [];
  let baseOffset = baseStart;
  let revisedOffset = revisedStart;
  for (const segment of segments) {
    const baseLength = segment.type === "ins" ? 0 : segment.text.length;
    const revisedLength = segment.type === "del" ? 0 : segment.text.length;
    positioned.push(
      Object.freeze({
        ...segment,
        baseStart: baseOffset,
        baseEnd: baseOffset + baseLength,
        revisedStart: revisedOffset,
        revisedEnd: revisedOffset + revisedLength,
      }),
    );
    baseOffset += baseLength;
    revisedOffset += revisedLength;
  }
  return positioned;
};

const tableLocationEqual = (
  base: FolioContentBlock["table"],
  revised: FolioContentBlock["table"],
): boolean => {
  if (base === undefined || revised === undefined) return base === revised;
  return Object.values(FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS).every(
    ({ field, validation }) => {
      if (validation !== "identity") return base[field] === revised[field];
      const baseIdentity = base[field];
      const revisedIdentity = revised[field];
      return (
        baseIdentity !== undefined &&
        revisedIdentity !== undefined &&
        baseIdentity.type === revisedIdentity.type &&
        baseIdentity.id === revisedIdentity.id
      );
    },
  );
};

const containerPathEqual = (
  base: FolioContentBlock["containerPath"],
  revised: FolioContentBlock["containerPath"],
): boolean => {
  if ((base?.length ?? 0) !== (revised?.length ?? 0)) return false;
  return (base ?? []).every((entry, index) =>
    Object.values(FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS).every(
      ({ field, validation }) => {
        const counterpart = revised?.[index];
        if (counterpart === undefined) return false;
        if (validation !== "identity") return entry[field] === counterpart[field];
        return (
          entry.identity.type === counterpart.identity.type &&
          entry.identity.id === counterpart.identity.id
        );
      },
    ),
  );
};

const structuralBoundariesEqual = (
  base: FolioContentBlock["structuralBoundaries"],
  revised: FolioContentBlock["structuralBoundaries"],
): boolean => {
  if ((base?.length ?? 0) !== (revised?.length ?? 0)) return false;
  return (base ?? []).every((boundary, index) =>
    Object.values(FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS).every(
      ({ field }) => boundary[field] === revised?.[index]?.[field],
    ),
  );
};

export const changedFolioContentParagraphFormatting = (
  base: FolioContentBlock,
  revised: FolioContentBlock,
): FolioContentParagraphFormattingChange =>
  Object.freeze({
    authored: changedFolioContentProperties(
      base.paragraphFormatting.authored,
      revised.paragraphFormatting.authored,
    ),
    effective: changedFolioContentProperties(
      base.paragraphFormatting.effective,
      revised.paragraphFormatting.effective,
    ),
  });

const changedBlockProperties = (
  base: FolioContentBlock,
  revised: FolioContentBlock,
): readonly FolioContentBlockChange[] => {
  const changed: FolioContentBlockChange[] = [];
  for (const descriptor of Object.values(FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS)) {
    switch (descriptor.comparison) {
      case "none":
      case "paragraph":
      case "inline":
        continue;
      case "kind":
        if (base.kind !== revised.kind) {
          changed.push(Object.freeze({ field: "kind", base: base.kind, revised: revised.kind }));
        }
        break;
      case "properties": {
        const changes = changedFolioContentProperties(base.blockProperties, revised.blockProperties);
        if (changes.length > 0) {
          changed.push(Object.freeze({ field: "blockProperties", changes }));
        }
        break;
      }
      case "table":
        if (!tableLocationEqual(base.table, revised.table)) {
          changed.push(
            Object.freeze({
              field: "table",
              base:
                base.table === undefined
                  ? Object.freeze({ type: "absent" })
                  : Object.freeze({ type: "present", value: base.table }),
              revised:
                revised.table === undefined
                  ? Object.freeze({ type: "absent" })
                  : Object.freeze({ type: "present", value: revised.table }),
            }),
          );
        }
        break;
      case "container":
        if (!containerPathEqual(base.containerPath, revised.containerPath)) {
          changed.push(
            Object.freeze({
              field: "containerPath",
              base: base.containerPath,
              revised: revised.containerPath,
            }),
          );
        }
        break;
      case "structural-boundaries":
        if (!structuralBoundariesEqual(base.structuralBoundaries, revised.structuralBoundaries)) {
          changed.push(
            Object.freeze({
              field: "structuralBoundaries",
              base: base.structuralBoundaries,
              revised: revised.structuralBoundaries,
            }),
          );
        }
        break;
      default: {
        const unreachable: never = descriptor;
        return panic("Unhandled block-property comparison", { descriptor: unreachable });
      }
    }
  }
  return Object.freeze(changed);
};

type ParagraphMarkPlan<Block extends FolioContentBlock> =
  | {
      type: "split";
      baseBlock: Block;
      revisedBlocks: readonly [Block, Block];
      offset: number;
      separator: string;
    }
  | {
      type: "merge";
      baseBlocks: readonly [Block, Block];
      revisedBlock: Block;
      separator: string;
    };

const separatorBetween = (whole: string, head: string, tail: string): string | null => {
  if (head.length === 0 || tail.length === 0 || whole.length < head.length + tail.length) {
    return null;
  }
  if (!whole.startsWith(head) || !whole.endsWith(tail)) {
    return null;
  }
  const separator = whole.slice(head.length, whole.length - tail.length);
  return separator.length === 0 || /^\s+$/u.test(separator) ? separator : null;
};

export const detectFolioContentParagraphMarkPlans = <Block extends FolioContentBlock>(
  steps: readonly FolioContentAlignmentStep<Block>[],
): ReadonlyMap<number, ParagraphMarkPlan<Block>> => {
  const plans = new Map<number, ParagraphMarkPlan<Block>>();
  for (let index = 0; index < steps.length - 1; index++) {
    const step = steps[index];
    const next = steps[index + 1];
    if (step === undefined || next === undefined) continue;

    if (step.type === "pair" && next.type === "revisedOnly") {
      const separator = separatorBetween(
        step.baseBlock.text,
        step.revisedBlock.text,
        next.block.text,
      );
      if (separator !== null && contentBlocksShareContainer(step.revisedBlock, next.block)) {
        plans.set(index, {
          type: "split",
          baseBlock: step.baseBlock,
          revisedBlocks: [step.revisedBlock, next.block],
          offset: step.revisedBlock.text.length,
          separator,
        });
        index++;
      }
      continue;
    }

    if (step.type === "revisedOnly" && next.type === "pair") {
      const separator = separatorBetween(
        next.baseBlock.text,
        step.block.text,
        next.revisedBlock.text,
      );
      if (separator !== null && contentBlocksShareContainer(step.block, next.revisedBlock)) {
        plans.set(index, {
          type: "split",
          baseBlock: next.baseBlock,
          revisedBlocks: [step.block, next.revisedBlock],
          offset: step.block.text.length,
          separator,
        });
        index++;
      }
      continue;
    }

    if (step.type === "pair" && next.type === "baseOnly") {
      const separator = separatorBetween(
        step.revisedBlock.text,
        step.baseBlock.text,
        next.block.text,
      );
      if (separator !== null && contentBlocksShareContainer(step.baseBlock, next.block)) {
        plans.set(index, {
          type: "merge",
          baseBlocks: [step.baseBlock, next.block],
          revisedBlock: step.revisedBlock,
          separator,
        });
        index++;
      }
      continue;
    }

    if (step.type === "baseOnly" && next.type === "pair") {
      const separator = separatorBetween(
        next.revisedBlock.text,
        step.block.text,
        next.baseBlock.text,
      );
      if (separator !== null && contentBlocksShareContainer(step.block, next.baseBlock)) {
        plans.set(index, {
          type: "merge",
          baseBlocks: [step.block, next.baseBlock],
          revisedBlock: next.revisedBlock,
          separator,
        });
        index++;
      }
    }
  }
  return plans;
};

type TokenProfile = { count: number; occurrences: ReadonlyMap<string, number> };

const tokenProfile = (text: string): TokenProfile | null => {
  if (text.length > MAX_MOVE_PROFILE_CODE_UNITS) return null;
  const occurrences = new Map<string, number>();
  let count = 0;
  for (const match of text.matchAll(/\S+/gu)) {
    count++;
    if (count > MAX_MOVE_PROFILE_TOKENS) return null;
    const token = match[0];
    occurrences.set(token, (occurrences.get(token) ?? 0) + 1);
  }
  return count >= MOVE_MINIMUM_WORD_COUNT ? { count, occurrences } : null;
};

const tokenSimilarity = (
  base: TokenProfile,
  revised: TokenProfile,
  maximumLookups: number,
): { similarity: number; lookups: number } | null => {
  const [smaller, larger] =
    base.occurrences.size <= revised.occurrences.size
      ? [base.occurrences, revised.occurrences]
      : [revised.occurrences, base.occurrences];
  if (smaller.size > maximumLookups) {
    return null;
  }
  let shared = 0;
  for (const [token, count] of smaller) {
    shared += Math.min(larger.get(token) ?? 0, count);
  }
  return {
    similarity: (2 * shared) / (base.count + revised.count),
    lookups: smaller.size,
  };
};

type MovePair<Block extends FolioContentBlock> = {
  baseBlock: Block;
  revisedBlock: Block;
};

type MoveCandidate<Block extends FolioContentBlock> = {
  block: Block;
  moveScope: Extract<FolioContentAlignmentStep<Block>, { type: "baseOnly" }>["moveScope"];
  profile: TokenProfile;
  order: number;
};

const detectFolioContentMoves = <Block extends FolioContentBlock>({
  steps,
  consumedStepIndexes,
  engine,
}: {
  steps: readonly FolioContentAlignmentStep<Block>[];
  consumedStepIndexes: ReadonlySet<number>;
  engine: Pick<ContentComparisonEngine, "scoreMoveSimilarity">;
}): readonly MovePair<Block>[] => {
  const stableBaseById = new Map<string, Pick<MoveCandidate<Block>, "block" | "moveScope">>();
  const exactCandidatesByBucket = new Map<number, Map<string, MoveCandidate<Block>[]>>();
  const similarityCandidatesByBucket = new Map<
    number,
    Map<number, Map<string, MoveCandidate<Block>>>
  >();
  let candidateOrder = 0;
  for (const [index, step] of steps.entries()) {
    if (consumedStepIndexes.has(index) || step.type !== "baseOnly") continue;
    if (folioContentIdentityPairDisposition(step.block.identity, step.block.identity) === "anchor") {
      stableBaseById.set(step.block.identity.id, {
        block: step.block,
        moveScope: step.moveScope,
      });
    }
    const profile = tokenProfile(step.block.text);
    if (!profile) continue;
    const candidate = {
      block: step.block,
      moveScope: step.moveScope,
      profile,
      order: candidateOrder++,
    };
    let candidatesByText = exactCandidatesByBucket.get(step.moveScope.bucket);
    if (!candidatesByText) {
      candidatesByText = new Map();
      exactCandidatesByBucket.set(step.moveScope.bucket, candidatesByText);
    }
    const queue = candidatesByText.get(step.block.text);
    if (!queue) {
      candidatesByText.set(step.block.text, [candidate]);
    } else if (queue.length < MAX_MOVE_CANDIDATES_PER_TEXT) {
      queue.push(candidate);
    }
    let candidatesByGap = similarityCandidatesByBucket.get(step.moveScope.bucket);
    if (!candidatesByGap) {
      candidatesByGap = new Map();
      similarityCandidatesByBucket.set(step.moveScope.bucket, candidatesByGap);
    }
    let gapCandidates = candidatesByGap.get(step.moveScope.gap);
    if (!gapCandidates) {
      gapCandidates = new Map();
      candidatesByGap.set(step.moveScope.gap, gapCandidates);
    }
    gapCandidates.set(step.block.identity.id, candidate);
  }

  const taken = new Set<string>();
  const takenRevised = new Set<string>();
  const moves: MovePair<Block>[] = [];
  const removeSimilarityCandidate = (candidate: {
    block: Block;
    moveScope: MoveCandidate<Block>["moveScope"];
  }): void => {
    const candidatesByGap = similarityCandidatesByBucket.get(candidate.moveScope.bucket);
    const gapCandidates = candidatesByGap?.get(candidate.moveScope.gap);
    gapCandidates?.delete(candidate.block.identity.id);
    if (gapCandidates?.size === 0) {
      candidatesByGap?.delete(candidate.moveScope.gap);
    }
    if (candidatesByGap?.size === 0) {
      similarityCandidatesByBucket.delete(candidate.moveScope.bucket);
    }
  };

  // Stable identity is stronger than either text heuristic. Claim every such
  // counterpart before walking revised blocks in order, so an earlier
  // positional candidate cannot steal its source through equal or similar text.
  for (const [index, step] of steps.entries()) {
    if (consumedStepIndexes.has(index) || step.type !== "revisedOnly") continue;
    const byId = stableBaseById.get(step.block.identity.id);
    const candidate =
      byId &&
      folioContentIdentityPairDisposition(byId.block.identity, step.block.identity) === "anchor"
        ? byId
        : undefined;
    if (!candidate || taken.has(candidate.block.identity.id)) continue;
    taken.add(candidate.block.identity.id);
    takenRevised.add(step.block.identity.id);
    removeSimilarityCandidate(candidate);
    if (candidate.moveScope.bucket === step.moveScope.bucket) {
      moves.push({ baseBlock: candidate.block, revisedBlock: step.block });
    }
  }

  // Exact text wins over edited similarity across the whole stream for the
  // same reason: a merely similar earlier candidate must not consume the only
  // exact source of a later one.
  for (const [index, step] of steps.entries()) {
    if (
      consumedStepIndexes.has(index) ||
      step.type !== "revisedOnly" ||
      takenRevised.has(step.block.identity.id)
    ) {
      continue;
    }
    const exactQueue = exactCandidatesByBucket.get(step.moveScope.bucket)?.get(step.block.text);
    const exact = exactQueue?.find(
      (candidate) =>
        !taken.has(candidate.block.identity.id) &&
        candidate.moveScope.gap !== step.moveScope.gap &&
        contentBlocksShareContainer(candidate.block, step.block) &&
        folioContentIdentityPairDisposition(candidate.block.identity, step.block.identity) !==
          "forbid",
    );
    if (exact) {
      taken.add(exact.block.identity.id);
      takenRevised.add(step.block.identity.id);
      removeSimilarityCandidate(exact);
      moves.push({ baseBlock: exact.block, revisedBlock: step.block });
    }
  }

  for (const [index, step] of steps.entries()) {
    if (
      consumedStepIndexes.has(index) ||
      step.type !== "revisedOnly" ||
      takenRevised.has(step.block.identity.id)
    ) {
      continue;
    }
    const revisedProfile = tokenProfile(step.block.text);
    if (!revisedProfile) continue;
    let best: { candidate: MoveCandidate<Block>; similarity: number } | null = null;
    const candidatesByGap = similarityCandidatesByBucket.get(step.moveScope.bucket);
    candidateGroups: for (const [gap, candidates] of candidatesByGap ?? []) {
      if (gap === step.moveScope.gap) continue;
      for (const candidate of candidates.values()) {
        if (
          !contentBlocksShareContainer(candidate.block, step.block) ||
          folioContentIdentityPairDisposition(candidate.block.identity, step.block.identity) ===
          "forbid"
        ) {
          continue;
        }
        const scored = engine.scoreMoveSimilarity(candidate.profile, revisedProfile);
        if (scored.status === "exhausted") break candidateGroups;
        if (scored.status === "unscored") continue;
        const { similarity } = scored;
        if (
          similarity >= MOVE_SIMILARITY_THRESHOLD &&
          (best === null ||
            similarity > best.similarity ||
            (similarity === best.similarity && candidate.order < best.candidate.order))
        ) {
          best = { candidate, similarity };
        }
      }
    }
    if (best) {
      taken.add(best.candidate.block.identity.id);
      takenRevised.add(step.block.identity.id);
      removeSimilarityCandidate(best.candidate);
      moves.push({ baseBlock: best.candidate.block, revisedBlock: step.block });
    }
  }
  return moves;
};

type Relation = {
  id: number;
  baseBlocks: readonly FolioContentBlock[];
  revisedBlocks: readonly FolioContentBlock[];
  event: FolioContentComparisonEvent;
};

const ownedBlockGroup = (
  blocks: readonly FolioContentBlock[],
): FolioContentBlockGroup => {
  const first = blocks.at(0);
  if (!first) return panic("A structural alignment step must own at least one block");
  const group: [FolioContentBlock, ...FolioContentBlock[]] = [first, ...blocks.slice(1)];
  return Object.freeze(group);
};

type FolioContentStructuralAlignmentStep<Block extends FolioContentBlock> = Extract<
  FolioContentAlignmentStep<Block>,
  { readonly type: "baseTable" | "revisedTable" | "baseRow" | "revisedRow" | "baseColumn" | "revisedColumn" }
>;

const structuralChangeForStep = <Block extends FolioContentBlock>(
  step: FolioContentStructuralAlignmentStep<Block>,
): FolioContentStructuralChange => {
  switch (step.type) {
    case "baseTable":
      return Object.freeze({
        type: "table-delete",
        tableIndex: step.location.tableIndex,
        blocks: ownedBlockGroup(step.blocks),
      });
    case "revisedTable":
      return Object.freeze({
        type: "table-insert",
        tableIndex: step.location.tableIndex,
        blocks: ownedBlockGroup(step.blocks),
      });
    case "baseRow":
      return Object.freeze({
        type: "table-row-delete",
        tableIndex: step.location.tableIndex,
        rowIndex: step.location.rowIndex,
        blocks: ownedBlockGroup(step.blocks),
      });
    case "revisedRow":
      return Object.freeze({
        type: "table-row-insert",
        tableIndex: step.location.tableIndex,
        rowIndex: step.location.rowIndex,
        blocks: ownedBlockGroup(step.blocks),
      });
    case "baseColumn":
      return Object.freeze({
        type: "table-column-delete",
        tableIndex: step.location.tableIndex,
        columnIndex: step.columnIndex,
        blocks: ownedBlockGroup(step.blocks),
      });
    case "revisedColumn":
      return Object.freeze({
        type: "table-column-insert",
        tableIndex: step.location.tableIndex,
        columnIndex: step.columnIndex,
        blocks: ownedBlockGroup(step.blocks),
        anchor: Object.freeze({ ...step.anchor }),
      });
    default: {
      const unreachable: never = step;
      return panic("Unhandled structural alignment step", { step: unreachable });
    }
  }
};

const contentBlockRange = (
  block: FolioContentBlock,
  startOffset: number,
  endOffset: number,
): FolioContentBlockRange => {
  if (
    !Number.isSafeInteger(startOffset) ||
    !Number.isSafeInteger(endOffset) ||
    startOffset < 0 ||
    endOffset < startOffset ||
    endOffset > block.text.length
  ) {
    return panic("A comparison relation received an invalid block range", {
      blockId: block.identity.id,
      startOffset,
      endOffset,
      textLength: block.text.length,
    });
  }
  return Object.freeze({
    block,
    startOffset,
    endOffset,
  });
};

const relationFormattingChange = ({
  base,
  revised,
  segments,
  maxRanges,
}: {
  base: FolioContentBlockRange;
  revised: FolioContentBlockRange;
  segments: readonly FolioContentTextSegment[];
  maxRanges: number;
}): FolioContentFormattingChange | null | "limit" => {
  const paragraph = changedFolioContentParagraphFormatting(base.block, revised.block);
  const ranges = pairedInlineFormattingSegments({
    baseBlock: base.block,
    revisedBlock: revised.block,
    equalRanges: segments
      .filter(({ type, text }) => type === "equal" && text.length > 0)
      .map(({ baseStart, baseEnd, revisedStart, revisedEnd }) => ({
        baseStart,
        baseEnd,
        revisedStart,
        revisedEnd,
      })),
    maxSegments: maxRanges,
  });
  if (ranges === null) return "limit";
  return paragraph.authored.length > 0 || paragraph.effective.length > 0 || ranges.length > 0
    ? Object.freeze({
        paragraph: Object.freeze({
          authored: paragraph.authored,
          effective: paragraph.effective,
        }),
        ranges: Object.freeze(
          ranges.map((range) =>
            Object.freeze({
              ...range,
              formatting: Object.freeze({
                authored: range.formatting.authored,
                effective: range.formatting.effective,
              }),
            }),
          ),
        ),
      })
    : null;
};

type PairRelationResult<Relation extends FolioContentPairRelation = FolioContentPairRelation> = {
  relation: Relation;
  formattingRanges: number;
};

type CreatePairRelationOptions = {
  baseBlock: FolioContentBlock;
  revisedBlock: FolioContentBlock;
  baseStart: number;
  baseEnd: number;
  revisedStart: number;
  revisedEnd: number;
  diffText: ContentComparisonEngine["diffText"];
  maxFormattingRanges: number;
};

function createPairRelation(
  options: CreatePairRelationOptions & { relationType: "whole" },
): PairRelationResult<FolioContentWholePairRelation> | "limit";
function createPairRelation(
  options: CreatePairRelationOptions & { relationType: "range" },
): PairRelationResult<FolioContentRangePairRelation> | "limit";
function createPairRelation(
  options: CreatePairRelationOptions & { relationType: "separator" },
): PairRelationResult<FolioContentSeparatorRelation> | "limit";
function createPairRelation({
  baseBlock,
  revisedBlock,
  baseStart,
  baseEnd,
  revisedStart,
  revisedEnd,
  relationType,
  diffText,
  maxFormattingRanges,
}: CreatePairRelationOptions & {
  relationType: FolioContentPairRelation["relationType"];
}): PairRelationResult | "limit" {
  const base = contentBlockRange(baseBlock, baseStart, baseEnd);
  const revised = contentBlockRange(revisedBlock, revisedStart, revisedEnd);
  const baseText = base.block.text.slice(base.startOffset, base.endOffset);
  const revisedText = revised.block.text.slice(revised.startOffset, revised.endOffset);
  const segments = Object.freeze(
    withTextOffsets(diffText(baseText, revisedText), base.startOffset, revised.startOffset),
  );
  let baseOffset = base.startOffset;
  let revisedOffset = revised.startOffset;
  for (const segment of segments) {
    if (
      segment.baseStart !== baseOffset ||
      segment.revisedStart !== revisedOffset ||
      (segment.type !== "ins" &&
        !base.block.text.startsWith(segment.text, segment.baseStart)) ||
      (segment.type !== "del" &&
        !revised.block.text.startsWith(segment.text, segment.revisedStart))
    ) {
      return panic("A word diff did not reconstruct its owned comparison ranges");
    }
    baseOffset = segment.baseEnd;
    revisedOffset = segment.revisedEnd;
  }
  if (baseOffset !== base.endOffset || revisedOffset !== revised.endOffset) {
    return panic("A word diff did not consume its owned comparison ranges");
  }
  if (relationType === "separator") {
    return {
      relation: Object.freeze({
        relationType,
        base,
        revised,
        segments,
        blockChanges: Object.freeze([] satisfies []),
        formatting: null,
      }),
      formattingRanges: 0,
    };
  }
  const formatting = relationFormattingChange({
    base,
    revised,
    segments,
    maxRanges: maxFormattingRanges,
  });
  if (formatting === "limit") return "limit";
  const blockChanges = changedBlockProperties(baseBlock, revisedBlock);
  const formattingRanges = formatting?.ranges.length ?? 0;
  if (relationType === "whole") {
    return {
      relation: Object.freeze({
        relationType,
        base,
        revised,
        segments,
        blockChanges,
        formatting,
      }),
      formattingRanges,
    };
  }
  return {
    relation: Object.freeze({
      relationType,
      base,
      revised,
      segments,
      blockChanges,
      formatting,
    }),
    formattingRanges,
  };
}

type CompareAlignedContentOptions = {
  captured: CapturedContentComparison;
  steps: readonly FolioContentAlignmentStep<FolioContentBlock>[];
  engine: ContentComparisonEngine;
  maximumChanges: number;
  maximumFormattingRanges: number;
  maximumStructuralMembers: number;
  maximumEvents: number;
};

const compareAlignedFolioContent = ({
  captured,
  steps,
  engine,
  maximumChanges,
  maximumFormattingRanges,
  maximumStructuralMembers,
  maximumEvents,
}: CompareAlignedContentOptions): Result<
  ContentComparisonExecution,
  FolioContentComparisonLimitError
> => {
  const { base, revised } = captured;
  const baseBlocks = base.blocks;
  const revisedBlocks = revised.blocks;
  const paragraphPlans = detectFolioContentParagraphMarkPlans(steps);
  const consumed = new Set([...paragraphPlans.keys()].map((index) => index + 1));
  const detectedMoves = detectFolioContentMoves({
    steps,
    consumedStepIndexes: consumed,
    engine,
  });
  const diffText = (baseText: string, revisedText: string): WordDiffSegment[] =>
    engine.diffText(baseText, revisedText);
  const relations: Relation[] = [];
  const baseRelation = new Map<string, Relation>();
  const revisedRelation = new Map<string, Relation>();
  let nextRelationId = 0;
  let remainingFormattingRanges = maximumFormattingRanges;
  let formattingRangeCount = 0;
  let changeCount = 0;
  let structuralMemberCount = 0;
  let eventCount = 0;

  const addRelation = (
    event: FolioContentComparisonEvent,
    relationBaseBlocks: readonly FolioContentBlock[],
    relationRevisedBlocks: readonly FolioContentBlock[],
    semanticCharge: "change" | "occurrence" = "change",
  ): boolean => {
    eventCount++;
    if (eventCount > maximumEvents) return false;
    const changeWeight =
      semanticCharge === "occurrence" || event.type === "unchanged" ? 0 : 1;
    changeCount += changeWeight;
    if (changeCount > maximumChanges) {
      return false;
    }
    if (event.type === "split" || event.type === "merge") {
      Object.freeze(event.relations);
    }
    const ownedEvent = Object.freeze(event);
    const relation = {
      id: nextRelationId++,
      baseBlocks: relationBaseBlocks,
      revisedBlocks: relationRevisedBlocks,
      event: ownedEvent,
    };
    relations.push(relation);
    for (const block of relationBaseBlocks) {
      if (baseRelation.has(block.identity.id)) {
        panic("Content alignment assigned one base block more than once", {
          blockId: block.identity.id,
        });
      }
      baseRelation.set(block.identity.id, relation);
    }
    for (const block of relationRevisedBlocks) {
      if (revisedRelation.has(block.identity.id)) {
        panic("Content alignment assigned one revised block more than once", {
          blockId: block.identity.id,
        });
      }
      revisedRelation.set(block.identity.id, relation);
    }
    return true;
  };

  const changeLimitExceeded = (): Result<never, FolioContentComparisonLimitError> =>
    Result.err(
      limitExceeded({
        input: "result",
        limit: "changes",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.changes,
        actual: FOLIO_CONTENT_COMPARISON_LIMITS.changes - maximumChanges + changeCount,
      }),
    );

  const formattingRangeLimitExceeded = (): Result<never, FolioContentComparisonLimitError> =>
    Result.err(
      limitExceeded({
        input: "result",
        limit: "formattingRanges",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.formattingRanges,
        actual: FOLIO_CONTENT_COMPARISON_LIMITS.formattingRanges + 1,
      }),
    );

  const structuralMemberLimitExceeded = (
    actual: number,
  ): Result<never, FolioContentComparisonLimitError> =>
    Result.err(
      limitExceeded({
        input: "result",
        limit: "structuralMembers",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.structuralMembers,
        actual:
          FOLIO_CONTENT_COMPARISON_LIMITS.structuralMembers - maximumStructuralMembers + actual,
      }),
    );

  const eventLimitExceeded = (): Result<never, FolioContentComparisonLimitError> =>
    Result.err(
      limitExceeded({
        input: "result",
        limit: "events",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.events,
        actual: FOLIO_CONTENT_COMPARISON_LIMITS.events - maximumEvents + eventCount,
      }),
    );

  const relationLimitExceeded = (): Result<never, FolioContentComparisonLimitError> =>
    eventCount > maximumEvents ? eventLimitExceeded() : changeLimitExceeded();

  const claimPairRelation = <Relation extends FolioContentPairRelation>(
    result: PairRelationResult<Relation> | "limit",
  ): Relation | "limit" => {
    if (result === "limit") return "limit";
    remainingFormattingRanges -= result.formattingRanges;
    formattingRangeCount += result.formattingRanges;
    return result.relation;
  };

  const moves: FolioContentMove[] = [];
  for (const [index, detected] of detectedMoves.entries()) {
    const relation = claimPairRelation(createPairRelation({
      baseBlock: detected.baseBlock,
      revisedBlock: detected.revisedBlock,
      baseStart: 0,
      baseEnd: detected.baseBlock.text.length,
      revisedStart: 0,
      revisedEnd: detected.revisedBlock.text.length,
      relationType: "whole",
      diffText,
      maxFormattingRanges: remainingFormattingRanges,
    }));
    if (relation === "limit") return formattingRangeLimitExceeded();
    moves.push(Object.freeze({ id: index + 1, relation }));
  }
  const moveByBaseId = new Map(
    moves.map((move) => [move.relation.base.block.identity.id, move] as const),
  );
  const moveByRevisedId = new Map(
    moves.map((move) => [move.relation.revised.block.identity.id, move] as const),
  );

  for (const [stepIndex, step] of steps.entries()) {
    if (consumed.has(stepIndex)) continue;
    const paragraphPlan = paragraphPlans.get(stepIndex);
    if (paragraphPlan?.type === "split") {
      const baseBlock = paragraphPlan.baseBlock;
      const [firstRevised, secondRevised] = paragraphPlan.revisedBlocks;
      const secondBaseStart = baseBlock.text.length - secondRevised.text.length;
      const first = claimPairRelation(createPairRelation({
        baseBlock,
        revisedBlock: firstRevised,
        baseStart: 0,
        baseEnd: firstRevised.text.length,
        revisedStart: 0,
        revisedEnd: firstRevised.text.length,
        relationType: "range",
        diffText,
        maxFormattingRanges: remainingFormattingRanges,
      }));
      if (first === "limit") return formattingRangeLimitExceeded();
      const second = claimPairRelation(createPairRelation({
        baseBlock,
        revisedBlock: secondRevised,
        baseStart: secondBaseStart,
        baseEnd: baseBlock.text.length,
        revisedStart: 0,
        revisedEnd: secondRevised.text.length,
        relationType: "range",
        diffText,
        maxFormattingRanges: remainingFormattingRanges,
      }));
      if (second === "limit") return formattingRangeLimitExceeded();
      const separator = claimPairRelation(createPairRelation({
        baseBlock,
        revisedBlock: firstRevised,
        baseStart: firstRevised.text.length,
        baseEnd: secondBaseStart,
        revisedStart: firstRevised.text.length,
        revisedEnd: firstRevised.text.length,
        relationType: "separator",
        diffText,
        maxFormattingRanges: remainingFormattingRanges,
      }));
      if (separator === "limit") {
        return panic("A text-only separator relation exceeded a formatting limit");
      }
      const event = {
        type: "split",
        relations: [first, second],
        separator,
      } as const;
      if (!addRelation(event, [baseBlock], paragraphPlan.revisedBlocks)) {
        return relationLimitExceeded();
      }
      continue;
    }
    if (paragraphPlan?.type === "merge") {
      const [firstBase, secondBase] = paragraphPlan.baseBlocks;
      const revisedBlock = paragraphPlan.revisedBlock;
      const secondRevisedStart = revisedBlock.text.length - secondBase.text.length;
      const first = claimPairRelation(createPairRelation({
        baseBlock: firstBase,
        revisedBlock,
        baseStart: 0,
        baseEnd: firstBase.text.length,
        revisedStart: 0,
        revisedEnd: firstBase.text.length,
        relationType: "range",
        diffText,
        maxFormattingRanges: remainingFormattingRanges,
      }));
      if (first === "limit") return formattingRangeLimitExceeded();
      const second = claimPairRelation(createPairRelation({
        baseBlock: secondBase,
        revisedBlock,
        baseStart: 0,
        baseEnd: secondBase.text.length,
        revisedStart: secondRevisedStart,
        revisedEnd: revisedBlock.text.length,
        relationType: "range",
        diffText,
        maxFormattingRanges: remainingFormattingRanges,
      }));
      if (second === "limit") return formattingRangeLimitExceeded();
      const separator = claimPairRelation(createPairRelation({
        baseBlock: firstBase,
        revisedBlock,
        baseStart: firstBase.text.length,
        baseEnd: firstBase.text.length,
        revisedStart: firstBase.text.length,
        revisedEnd: secondRevisedStart,
        relationType: "separator",
        diffText,
        maxFormattingRanges: remainingFormattingRanges,
      }));
      if (separator === "limit") {
        return panic("A text-only separator relation exceeded a formatting limit");
      }
      const event = {
        type: "merge",
        relations: [first, second],
        separator,
      } as const;
      if (!addRelation(event, paragraphPlan.baseBlocks, [revisedBlock])) {
        return relationLimitExceeded();
      }
      continue;
    }

    if (step.type === "pair") {
      const base = step.baseBlock;
      const revised = step.revisedBlock;
      const relation = claimPairRelation(createPairRelation({
        baseBlock: base,
        revisedBlock: revised,
        baseStart: 0,
        baseEnd: base.text.length,
        revisedStart: 0,
        revisedEnd: revised.text.length,
        relationType: "whole",
        diffText,
        maxFormattingRanges: remainingFormattingRanges,
      }));
      if (relation === "limit") {
        return formattingRangeLimitExceeded();
      }
      let event: FolioContentComparisonEvent;
      if (
        relation.segments.some(({ type }) => type !== "equal") ||
        relation.blockChanges.length > 0
      ) {
        event = {
          type: "modified",
          relation,
        };
      } else if (relation.formatting) {
        event = { type: "formatting", relation };
      } else {
        event = { type: "unchanged", relation };
      }
      if (!addRelation(event, [base], [revised])) {
        return relationLimitExceeded();
      }
      continue;
    }

    if (step.type === "baseOnly") {
      const move = moveByBaseId.get(step.block.identity.id);
      const event: FolioContentComparisonEvent = move
        ? { type: "movedFrom", move }
        : { type: "deleted", block: step.block };
      if (!addRelation(event, [step.block], [], move ? "occurrence" : "change")) {
        return relationLimitExceeded();
      }
      continue;
    }
    if (step.type === "revisedOnly") {
      const move = moveByRevisedId.get(step.block.identity.id);
      const event: FolioContentComparisonEvent = move
        ? { type: "movedTo", move }
        : { type: "inserted", block: step.block };
      if (!addRelation(event, [], [step.block])) {
        return relationLimitExceeded();
      }
      continue;
    }

    if (step.type === "tableReplacement") {
      if (changeCount + 1 > maximumChanges) {
        changeCount++;
        return changeLimitExceeded();
      }
      if (eventCount + 1 > maximumEvents) {
        eventCount++;
        return eventLimitExceeded();
      }
      const refinement = compareAlignedFolioContent({
        captured: {
          base: Object.freeze({ blocks: Object.freeze([...step.baseBlocks]) }),
          revised: Object.freeze({ blocks: Object.freeze([...step.revisedBlocks]) }),
        },
        steps: step.refinementSteps,
        engine,
        maximumChanges: maximumChanges - changeCount - 1,
        maximumFormattingRanges: remainingFormattingRanges,
        maximumStructuralMembers: maximumStructuralMembers - structuralMemberCount,
        maximumEvents: maximumEvents - eventCount - 1,
      });
      if (refinement.isErr()) return refinement;
      changeCount += refinement.value.changes;
      formattingRangeCount += refinement.value.formattingRanges;
      remainingFormattingRanges -= refinement.value.formattingRanges;
      structuralMemberCount += refinement.value.structuralMembers;
      eventCount += refinement.value.events;
      const baseGroup = ownedBlockGroup(step.baseBlocks);
      const revisedGroup = ownedBlockGroup(step.revisedBlocks);
      const replacement = Object.freeze({
        baseBlocks: baseGroup,
        revisedBlocks: revisedGroup,
        baseTableIndex: step.baseLocation.tableIndex,
        revisedTableIndex: step.revisedLocation.tableIndex,
        refinement: refinement.value.comparison,
      });
      if (
        !addRelation(
          { type: "tableReplacement", replacement },
          baseGroup,
          revisedGroup,
        )
      ) {
        return relationLimitExceeded();
      }
      continue;
    }

    const structural = structuralChangeForStep(step);
    if (changeCount + 1 > maximumChanges) {
      changeCount++;
      return changeLimitExceeded();
    }
    if (structuralMemberCount + structural.blocks.length > maximumStructuralMembers) {
      return structuralMemberLimitExceeded(structuralMemberCount + structural.blocks.length);
    }
    changeCount++;
    structuralMemberCount += structural.blocks.length;
    if (
      structural.type === "table-delete" ||
      structural.type === "table-row-delete" ||
      structural.type === "table-column-delete"
    ) {
      for (const [memberIndex, block] of structural.blocks.entries()) {
        if (
          !addRelation(
            { type: "structural", change: structural, memberIndex },
            [block],
            [],
            "occurrence",
          )
        ) {
          return relationLimitExceeded();
        }
      }
    } else {
      for (const [memberIndex, block] of structural.blocks.entries()) {
        if (
          !addRelation(
            { type: "structural", change: structural, memberIndex },
            [],
            [block],
            "occurrence",
          )
        ) {
          return relationLimitExceeded();
        }
      }
    }
  }

  const ordered: FolioContentComparisonEvent[] = [];
  const emitted = new Set<number>();
  let baseIndex = 0;
  let revisedIndex = 0;
  while (baseIndex < baseBlocks.length || revisedIndex < revisedBlocks.length) {
    const base = baseBlocks[baseIndex];
    const revised = revisedBlocks[revisedIndex];
    const fromBase = base ? baseRelation.get(base.identity.id) : undefined;
    const fromRevised = revised ? revisedRelation.get(revised.identity.id) : undefined;
    if (fromBase && emitted.has(fromBase.id)) {
      baseIndex++;
      continue;
    }
    if (fromRevised && emitted.has(fromRevised.id)) {
      revisedIndex++;
      continue;
    }
    let relation: Relation | undefined;
    if (fromBase && fromRevised && fromBase.id === fromRevised.id) {
      relation = fromBase;
    } else if (fromRevised?.baseBlocks.length === 0) {
      relation = fromRevised;
    } else if (fromBase?.revisedBlocks.length === 0) {
      relation = fromBase;
    }
    if (!relation) {
      return panic("Content alignment did not produce one monotone projection", {
        baseBlockId: base?.identity.id,
        revisedBlockId: revised?.identity.id,
      });
    }
    emitted.add(relation.id);
    ordered.push(relation.event);
    baseIndex += relation.baseBlocks.length;
    revisedIndex += relation.revisedBlocks.length;
  }
  if (emitted.size !== relations.length) {
    return panic("Content alignment left relations outside the ordered projection");
  }

  const movedFrom = new Set<FolioContentMove>();
  const movedTo = new Set<FolioContentMove>();
  const structuralMembers = new Map<FolioContentStructuralChange, number[]>();
  for (const event of ordered) {
    switch (event.type) {
      case "movedFrom":
        if (movedFrom.has(event.move)) {
          return panic("A move has more than one source event", { moveId: event.move.id });
        }
        movedFrom.add(event.move);
        break;
      case "movedTo":
        if (movedTo.has(event.move)) {
          return panic("A move has more than one destination event", { moveId: event.move.id });
        }
        movedTo.add(event.move);
        break;
      case "deleted":
      case "inserted":
      case "unchanged":
      case "modified":
      case "formatting":
      case "split":
      case "merge":
        break;
      case "tableReplacement":
        break;
      case "structural": {
        const members = structuralMembers.get(event.change);
        if (members) {
          members.push(event.memberIndex);
        } else {
          structuralMembers.set(event.change, [event.memberIndex]);
        }
        break;
      }
      default: {
        const unreachable: never = event;
        return panic("Unhandled comparison event cross-reference", { event: unreachable });
      }
    }
  }
  if (
    movedFrom.size !== moves.length ||
    movedTo.size !== moves.length ||
    moves.some((move) => !movedFrom.has(move) || !movedTo.has(move))
  ) {
    return panic("Move stream positions do not share one canonical move record");
  }
  for (const [change, members] of structuralMembers) {
    if (
      members.length !== change.blocks.length ||
      members.some((memberIndex, index) => memberIndex !== index)
    ) {
      return panic("Structural stream positions do not exhaust one canonical member set", {
        type: change.type,
      });
    }
  }

  return Result.ok({
    comparison: completeContentComparison(Object.freeze(ordered)),
    changes: changeCount,
    formattingRanges: formattingRangeCount,
    structuralMembers: structuralMemberCount,
    events: eventCount,
  });
};

type ExecuteCapturedContentComparisonOptions = {
  captured: CapturedContentComparison;
  engine: ContentComparisonEngine;
  maximumChanges: number;
  maximumFormattingRanges: number;
  maximumStructuralMembers: number;
  maximumEvents: number;
};

const executeCapturedContentComparison = ({
  captured,
  engine,
  maximumChanges,
  maximumFormattingRanges,
  maximumStructuralMembers,
  maximumEvents,
}: ExecuteCapturedContentComparisonOptions): Result<
  ContentComparisonExecution,
  FolioContentComparisonLimitError
> => {
  const { base, revised } = captured;
  const steps = engine.alignContentStructure({
    baseBlocks: base.blocks,
    revisedBlocks: revised.blocks,
  });
  return compareAlignedFolioContent({
    captured,
    steps,
    engine,
    maximumChanges,
    maximumFormattingRanges,
    maximumStructuralMembers,
    maximumEvents,
  });
};

/**
 * Compare two representation-neutral ordered content snapshots.
 *
 * Folio captures only the fields declared by the public input types. Additional
 * caller metadata is ignored rather than enumerated, bounding the properties
 * Folio requests. As with every synchronous JavaScript API, caller-defined
 * proxy traps remain caller-executed code.
 */
export const compareContent = (
  options: CompareContentOptions,
): Result<FolioContentComparison, FolioContentComparisonError> => {
  const capturedOptions = captureKnownDataRecord(
    options,
    Object.values(COMPARE_CONTENT_OPTION_FIELD_DESCRIPTORS).map(({ field }) => field),
    "options",
    { side: "options" },
  );
  if (capturedOptions.isErr()) return Result.err(capturedOptions.error);
  const base = capturedOptions.value.get("base");
  const revised = capturedOptions.value.get("revised");
  const granularity = capturedOptions.value.get("granularity");
  if (
    granularity !== undefined &&
    !WORD_DIFF_GRANULARITIES.some((candidate) => candidate === granularity)
  ) {
    return Result.err(
      invalidInput("options", "granularity", "Comparison granularity must be word or character."),
    );
  }
  const workSession = createContentComparisonWorkSession({ ...(granularity && { granularity }) });
  const operation = workSession.captureComparison({ base, revised });
  if (operation.isErr()) {
    return operation.error instanceof FolioContentComparisonSessionError
      ? panic("A fresh content comparison session rejected its first capture", {
          cause: operation.error,
        })
      : Result.err(operation.error);
  }
  const compared = operation.value.compare();
  if (compared.isErr() && compared.error instanceof FolioContentComparisonSessionError) {
    return panic("A fresh content comparison operation failed its first consumption", {
      cause: compared.error,
    });
  }
  return compared;
};
