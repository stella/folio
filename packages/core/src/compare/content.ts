/**
 * Pure comparison of two ordered, representation-neutral content snapshots.
 *
 * Every event carries exact base and revised block tuples. Concatenating the
 * base tuples reconstructs the base snapshot; concatenating the revised tuples
 * reconstructs the revised snapshot. This remains true for moves, paragraph
 * splits/merges, and table-column changes whose cells are not contiguous.
 */

import { panic, Result, TaggedError } from "better-result";

import {
  createWordDiffSession,
  WORD_DIFF_GRANULARITIES,
  type WordDiffGranularity,
  type WordDiffSegment,
} from "../ai-edits/word-diff";
import { inlineFormattingSegments } from "./formatting";
import {
  alignFolioContentStructure,
  contentBlocksShareContainer,
  createFolioContentAlignmentWorkSession,
  type FolioContentAlignmentStep,
  type FolioContentAlignmentWorkSession,
} from "./content-alignment";
import type {
  FolioContentBlock,
  FolioContentIdStability,
  FolioContentInlineFormattingPatch,
  FolioContentParagraphSpacing,
  FolioContentSnapshot,
} from "./content-types";

/** Maximum blocks accepted on either side of one comparison. */
export const MAX_FOLIO_CONTENT_BLOCKS = 100_000;

/** Maximum non-unchanged events returned by one comparison. */
export const MAX_FOLIO_CONTENT_CHANGES = 10_000;

/** Words a one-sided block needs before it may be classified as a move. */
const MOVE_MINIMUM_WORD_COUNT = 3;

/** Same-text move candidates retained for one repeated value. */
const MAX_MOVE_CANDIDATES_PER_TEXT = 64;

/** Pairwise comparisons allowed for edited-move discovery. */
const MAX_MOVE_SIMILARITY_COMPARISONS = 20_000;

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
  limit: "base-blocks" | "revised-blocks" | "changes";
  maximum: number;
  actual: number;
}> {}

/** Errors returned by {@link compareContent}. */
export type FolioContentComparisonError =
  | InvalidFolioContentComparisonError
  | FolioContentComparisonLimitError;

/** One Folio word-diff segment with JavaScript-slice-compatible offsets. */
export type FolioContentTextSegment = {
  type: "equal" | "del" | "ins";
  text: string;
  baseStart: number;
  baseEnd: number;
  revisedStart: number;
  revisedEnd: number;
};

/** Target-side paragraph properties that differ from the base block. */
export type FolioContentParagraphFormattingPatch = {
  styleId?: string | null;
  listLevel?: number | null;
  alignment?: FolioContentBlock["directAlignment"] | null;
  spacing?: FolioContentParagraphSpacing | null;
};

/** Presentation differences for one text-aligned block pair. */
export type FolioContentFormatRange = {
  startOffset: number;
  endOffset: number;
  formatting: FolioContentInlineFormattingPatch;
};

/** Presentation differences for one text-aligned block pair. */
export type FolioContentFormattingChange = {
  paragraph?: FolioContentParagraphFormattingPatch;
  ranges: readonly FolioContentFormatRange[];
};

/** Non-presentation block fields that changed on a paired block. */
export type FolioContentBlockProperty = "kind" | "headingLevel" | "displayLabel";

type PairedEvent<Block extends FolioContentBlock> = {
  baseBlocks: readonly [Block];
  revisedBlocks: readonly [Block];
};

type BaseOnlyEvent<Block extends FolioContentBlock> = {
  baseBlocks: readonly [Block];
  revisedBlocks: readonly [];
};

type RevisedOnlyEvent<Block extends FolioContentBlock> = {
  baseBlocks: readonly [];
  revisedBlocks: readonly [Block];
};

/**
 * One item in the complete comparison stream. Tuple cardinality is fixed by
 * the discriminator, making both document projections mechanically exact.
 */
export type FolioContentComparisonEvent<Block extends FolioContentBlock = FolioContentBlock> =
  | ({ type: "unchanged" } & PairedEvent<Block>)
  | ({
      type: "modified";
      segments: readonly FolioContentTextSegment[];
      changedProperties: readonly FolioContentBlockProperty[];
      formatting?: FolioContentFormattingChange;
    } & PairedEvent<Block>)
  | ({ type: "formatting"; formatting: FolioContentFormattingChange } & PairedEvent<Block>)
  | ({ type: "inserted"; structuralChangeId?: number } & RevisedOnlyEvent<Block>)
  | ({ type: "deleted"; structuralChangeId?: number } & BaseOnlyEvent<Block>)
  | ({ type: "movedFrom"; moveId: number } & BaseOnlyEvent<Block>)
  | ({
      type: "movedTo";
      moveId: number;
      baseBlockId: string;
      segments?: readonly FolioContentTextSegment[];
      changedProperties?: readonly FolioContentBlockProperty[];
      formatting?: FolioContentFormattingChange;
    } & RevisedOnlyEvent<Block>)
  | {
      type: "split";
      baseBlocks: readonly [Block];
      revisedBlocks: readonly [Block, Block];
      offset: number;
      separator: string;
    }
  | {
      type: "merge";
      baseBlocks: readonly [Block, Block];
      revisedBlocks: readonly [Block];
      separator: string;
    };

type BaseStructuralChange = {
  id: number;
  baseBlockIds: readonly string[];
};

type RevisedStructuralChange = {
  id: number;
  revisedBlockIds: readonly string[];
};

/** A grouped table operation referenced by its block-granular stream events. */
export type FolioContentStructuralChange =
  | ({ type: "table-delete"; tableIndex: number } & BaseStructuralChange)
  | ({ type: "table-insert"; tableIndex: number } & RevisedStructuralChange)
  | ({ type: "table-row-delete"; tableIndex: number; rowIndex: number } & BaseStructuralChange)
  | ({ type: "table-row-insert"; tableIndex: number; rowIndex: number } & RevisedStructuralChange)
  | ({ type: "table-column-delete"; tableIndex: number; columnIndex: number } & BaseStructuralChange)
  | ({
      type: "table-column-insert";
      tableIndex: number;
      columnIndex: number;
    } & RevisedStructuralChange);

/** Result of one representation-neutral story comparison. */
export type FolioContentComparison<Block extends FolioContentBlock = FolioContentBlock> = {
  events: readonly FolioContentComparisonEvent<Block>[];
  structuralChanges: readonly FolioContentStructuralChange[];
};

/** Inputs to {@link compareContent}. */
export type CompareContentOptions<Kind extends string = string> = {
  base: FolioContentSnapshot<Kind>;
  revised: FolioContentSnapshot<Kind>;
  /** Token size for modified-block segments; defaults to `"word"`. */
  granularity?: WordDiffGranularity;
};

export type FolioContentComparisonWorkSession = {
  alignment: FolioContentAlignmentWorkSession;
  remainingMoveComparisons: number;
  diffText: ReturnType<typeof createWordDiffSession>["diff"];
};

export const createContentComparisonWorkSession = (
  granularity?: WordDiffGranularity,
): FolioContentComparisonWorkSession => ({
  alignment: createFolioContentAlignmentWorkSession(),
  remainingMoveComparisons: MAX_MOVE_SIMILARITY_COMPARISONS,
  diffText: createWordDiffSession({ ...(granularity && { granularity }) }).diff,
});

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

const isFiniteInteger = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value) && Number.isInteger(value);

const isFiniteNumber = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const PARAGRAPH_ALIGNMENTS = new Set([
  "left",
  "center",
  "right",
  "both",
  "distribute",
  "mediumKashida",
  "highKashida",
  "lowKashida",
  "thaiDistribute",
]);

const validateRunFormatting = (
  block: FolioContentBlock,
  side: "base" | "revised",
  blockIndex: number,
): InvalidFolioContentComparisonError | null => {
  for (const [runIndex, run] of (block.previewRuns ?? []).entries()) {
    if (run.directFormatting !== undefined && !isRecord(run.directFormatting)) {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].directFormatting`,
        "Direct inline formatting must be an object.",
        blockIndex,
      );
    }
    for (const property of ["bold", "italic", "underline", "strike"] as const) {
      if (run[property] !== undefined && typeof run[property] !== "boolean") {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].${property}`,
          "Inline boolean formatting must be boolean.",
          blockIndex,
        );
      }
      const direct = run.directFormatting?.[property];
      if (direct !== undefined && typeof direct !== "boolean") {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].directFormatting.${property}`,
          "Direct inline boolean formatting must be boolean.",
          blockIndex,
        );
      }
    }
    for (const property of ["fontFamily", "color"] as const) {
      const effective = run[property];
      if (effective !== undefined && typeof effective !== "string") {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].${property}`,
          "Inline string formatting must be a string.",
          blockIndex,
        );
      }
      const direct = run.directFormatting?.[property];
      if (direct !== undefined && direct !== null && typeof direct !== "string") {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].directFormatting.${property}`,
          "Direct inline string formatting must be a string or null.",
          blockIndex,
        );
      }
    }
    if (
      run.fontSizePt !== undefined &&
      (!isFiniteNumber(run.fontSizePt) || run.fontSizePt < 0)
    ) {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].fontSizePt`,
        "Inline font size must be a non-negative finite number.",
        blockIndex,
      );
    }
    const directSize = run.directFormatting?.fontSizePt;
    if (
      directSize !== undefined &&
      directSize !== null &&
      (!isFiniteNumber(directSize) || directSize < 0)
    ) {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].directFormatting.fontSizePt`,
        "Direct inline font size must be a non-negative finite number or null.",
        blockIndex,
      );
    }
  }
  return null;
};

const validateParagraphFormatting = (
  block: FolioContentBlock,
  side: "base" | "revised",
  blockIndex: number,
): InvalidFolioContentComparisonError | null => {
  for (const property of ["styleId", "displayLabel"] as const) {
    if (block[property] !== undefined && typeof block[property] !== "string") {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].${property}`,
        "Optional block labels and style ids must be strings.",
        blockIndex,
      );
    }
  }
  if (
    block.headingLevel !== undefined &&
    (!isFiniteInteger(block.headingLevel) || block.headingLevel < 1)
  ) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].headingLevel`,
      "Heading levels must be positive integers.",
      blockIndex,
    );
  }
  if (block.listLevel !== undefined && (!isFiniteInteger(block.listLevel) || block.listLevel < 0)) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].listLevel`,
      "List levels must be non-negative integers.",
      blockIndex,
    );
  }
  if (
    block.directAlignment !== undefined &&
    !PARAGRAPH_ALIGNMENTS.has(block.directAlignment)
  ) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].directAlignment`,
      "Paragraph alignment is not recognized.",
      blockIndex,
    );
  }
  const spacing = block.directSpacing;
  if (spacing === undefined) return null;
  if (!isRecord(spacing)) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].directSpacing`,
      "Direct paragraph spacing must be an object.",
      blockIndex,
    );
  }
  for (const property of ["spaceBefore", "spaceAfter", "lineSpacing"] as const) {
    if (spacing[property] !== undefined && !isFiniteNumber(spacing[property])) {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].directSpacing.${property}`,
        "Paragraph spacing values must be finite numbers.",
        blockIndex,
      );
    }
  }
  for (const property of ["beforeAutospacing", "afterAutospacing"] as const) {
    if (spacing[property] !== undefined && typeof spacing[property] !== "boolean") {
      return invalidInput(
        side,
        `blocks[${String(blockIndex)}].directSpacing.${property}`,
        "Paragraph auto-spacing values must be boolean.",
        blockIndex,
      );
    }
  }
  if (
    spacing.lineSpacingRule !== undefined &&
    spacing.lineSpacingRule !== "auto" &&
    spacing.lineSpacingRule !== "exact" &&
    spacing.lineSpacingRule !== "atLeast"
  ) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].directSpacing.lineSpacingRule`,
      "Paragraph line-spacing rule is not recognized.",
      blockIndex,
    );
  }
  return null;
};

const validateTableLocation = (
  table: unknown,
  side: "base" | "revised",
  blockIndex: number,
): InvalidFolioContentComparisonError | null => {
  if (!isRecord(table)) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].table`,
      "A table location must be an object.",
      blockIndex,
    );
  }
  for (const field of [
    "outerTableIndex",
    "tableIndex",
    "rowIndex",
    "cellIndex",
    "gridColumnIndex",
    "paragraphIndex",
  ] as const) {
    if (!isFiniteInteger(table[field]) || table[field] < 0) {
      return invalidInput(side, `blocks[${String(blockIndex)}].table.${field}`, "Table indexes must be non-negative integers.", blockIndex);
    }
  }
  for (const field of ["columnSpan", "rowSpan"] as const) {
    if (!isFiniteInteger(table[field]) || table[field] < 1) {
      return invalidInput(side, `blocks[${String(blockIndex)}].table.${field}`, "Table spans must be positive integers.", blockIndex);
    }
  }
  return null;
};

const validateSnapshot = <Kind extends string>(
  snapshot: FolioContentSnapshot<Kind>,
  side: "base" | "revised",
): InvalidFolioContentComparisonError | FolioContentComparisonLimitError | null => {
  if (!snapshot || !Array.isArray(snapshot.blocks)) {
    return invalidInput(side, "blocks", "A content snapshot must contain an ordered blocks array.");
  }
  if (snapshot.blocks.length > MAX_FOLIO_CONTENT_BLOCKS) {
    return new FolioContentComparisonLimitError({
      message: `The ${side} snapshot contains more blocks than one comparison accepts.`,
      limit: `${side}-blocks`,
      maximum: MAX_FOLIO_CONTENT_BLOCKS,
      actual: snapshot.blocks.length,
    });
  }

  const ids = new Set<string>();
  const lastCoordinateByTable = new Map<string, readonly [number, number, number]>();
  let lastOuterTableIndex = -1;
  for (const [blockIndex, block] of snapshot.blocks.entries()) {
    if (!block || typeof block !== "object") {
      return invalidInput(side, `blocks[${String(blockIndex)}]`, "Every content block must be an object.", blockIndex);
    }
    if (typeof block.id !== "string" || block.id.length === 0) {
      return invalidInput(side, `blocks[${String(blockIndex)}].id`, "Every content block needs a non-empty id.", blockIndex);
    }
    if (ids.has(block.id)) {
      return invalidInput(side, `blocks[${String(blockIndex)}].id`, "Content block ids must be unique within a snapshot.", blockIndex);
    }
    ids.add(block.id);
    if (typeof block.kind !== "string" || block.kind.length === 0) {
      return invalidInput(side, `blocks[${String(blockIndex)}].kind`, "Every content block needs a non-empty kind.", blockIndex);
    }
    if (typeof block.text !== "string") {
      return invalidInput(side, `blocks[${String(blockIndex)}].text`, "Content block text must be a string.", blockIndex);
    }
    if (block.idStability !== undefined && block.idStability !== "stable" && block.idStability !== "positional") {
      return invalidInput(side, `blocks[${String(blockIndex)}].idStability`, "Block id stability must be stable or positional.", blockIndex);
    }
    if (block.previewRuns !== undefined) {
      if (!Array.isArray(block.previewRuns) || block.previewRuns.some((run) => !run || typeof run.text !== "string")) {
        return invalidInput(side, `blocks[${String(blockIndex)}].previewRuns`, "Preview runs must be an array of text runs.", blockIndex);
      }
      if (block.previewRuns.map(({ text }) => text).join("") !== block.text) {
        return invalidInput(side, `blocks[${String(blockIndex)}].previewRuns`, "Preview-run text must reconstruct the block text exactly.", blockIndex);
      }
    }
    const paragraphError = validateParagraphFormatting(block, side, blockIndex);
    if (paragraphError) {
      return paragraphError;
    }
    const runError = validateRunFormatting(block, side, blockIndex);
    if (runError) {
      return runError;
    }
    if (block.containerPath !== undefined) {
      if (
        !Array.isArray(block.containerPath) ||
        block.containerPath.some(
          (entry) =>
            !entry ||
            typeof entry.kind !== "string" ||
            entry.kind.length === 0 ||
            typeof entry.id !== "string" ||
            entry.id.length === 0,
        )
      ) {
        return invalidInput(side, `blocks[${String(blockIndex)}].containerPath`, "Container paths require non-empty kind and id values.", blockIndex);
      }
    }
    if (block.table !== undefined) {
      const error = validateTableLocation(block.table, side, blockIndex);
      if (error) {
        return error;
      }
      const table = block.table;
      if (table.outerTableIndex < lastOuterTableIndex) {
        return invalidInput(
          side,
          `blocks[${String(blockIndex)}].table`,
          "Table blocks must follow document order.",
          blockIndex,
        );
      }
      lastOuterTableIndex = table.outerTableIndex;
      const tableKey = `${String(table.outerTableIndex)}:${String(table.tableIndex)}`;
      const coordinate = [
        table.rowIndex,
        table.cellIndex,
        table.paragraphIndex,
      ] as const;
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
    }
  }
  return null;
};

const withTextOffsets = (segments: readonly WordDiffSegment[]): FolioContentTextSegment[] => {
  const positioned: FolioContentTextSegment[] = [];
  let baseOffset = 0;
  let revisedOffset = 0;
  for (const segment of segments) {
    const baseLength = segment.type === "ins" ? 0 : segment.text.length;
    const revisedLength = segment.type === "del" ? 0 : segment.text.length;
    positioned.push({
      ...segment,
      baseStart: baseOffset,
      baseEnd: baseOffset + baseLength,
      revisedStart: revisedOffset,
      revisedEnd: revisedOffset + revisedLength,
    });
    baseOffset += baseLength;
    revisedOffset += revisedLength;
  }
  return positioned;
};

const paragraphSpacingEqual = (
  base: FolioContentParagraphSpacing | undefined,
  revised: FolioContentParagraphSpacing | undefined,
): boolean =>
  base?.spaceBefore === revised?.spaceBefore &&
  base?.spaceAfter === revised?.spaceAfter &&
  base?.lineSpacing === revised?.lineSpacing &&
  base?.lineSpacingRule === revised?.lineSpacingRule &&
  base?.beforeAutospacing === revised?.beforeAutospacing &&
  base?.afterAutospacing === revised?.afterAutospacing;

export const changedFolioContentParagraphFormatting = (
  base: FolioContentBlock,
  revised: FolioContentBlock,
): FolioContentParagraphFormattingPatch | null => {
  const patch: FolioContentParagraphFormattingPatch = {};
  if ((base.styleId ?? null) !== (revised.styleId ?? null)) {
    patch.styleId = revised.styleId ?? null;
  }
  if (base.listLevel !== revised.listLevel) {
    patch.listLevel = revised.listLevel ?? null;
  }
  if (base.directAlignment !== revised.directAlignment) {
    patch.alignment = revised.directAlignment ?? null;
  }
  if (!paragraphSpacingEqual(base.directSpacing, revised.directSpacing)) {
    patch.spacing = revised.directSpacing ?? null;
  }
  return Object.keys(patch).length > 0 ? patch : null;
};

const changedBlockProperties = (
  base: FolioContentBlock,
  revised: FolioContentBlock,
): FolioContentBlockProperty[] => {
  const changed: FolioContentBlockProperty[] = [];
  if (base.kind !== revised.kind) changed.push("kind");
  if (base.headingLevel !== revised.headingLevel) changed.push("headingLevel");
  if (base.displayLabel !== revised.displayLabel) changed.push("displayLabel");
  return changed;
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
  for (const [index, step] of steps.entries()) {
    const next = steps[index + 1];
    if (step.type !== "pair" || next === undefined) continue;
    if (next.type === "revisedOnly") {
      const separator = separatorBetween(step.baseBlock.text, step.revisedBlock.text, next.block.text);
      if (separator !== null && contentBlocksShareContainer(step.revisedBlock, next.block)) {
        plans.set(index, {
          type: "split",
          baseBlock: step.baseBlock,
          revisedBlocks: [step.revisedBlock, next.block],
          offset: step.revisedBlock.text.length,
          separator,
        });
      }
      continue;
    }
    if (next.type !== "baseOnly") continue;
    const separator = separatorBetween(step.revisedBlock.text, step.baseBlock.text, next.block.text);
    if (separator !== null && contentBlocksShareContainer(step.baseBlock, next.block)) {
      plans.set(index, {
        type: "merge",
        baseBlocks: [step.baseBlock, next.block],
        revisedBlock: step.revisedBlock,
        separator,
      });
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

const tokenSimilarity = (base: TokenProfile, revised: TokenProfile): number => {
  let shared = 0;
  for (const [token, revisedCount] of revised.occurrences) {
    shared += Math.min(base.occurrences.get(token) ?? 0, revisedCount);
  }
  return (2 * shared) / (base.count + revised.count);
};

type MovePair<Block extends FolioContentBlock> = {
  baseBlock: Block;
  revisedBlock: Block;
};

const contentBlocksCanMoveTogether = (
  baseBlock: FolioContentBlock,
  revisedBlock: FolioContentBlock,
): boolean => {
  if (baseBlock.table === undefined && revisedBlock.table === undefined) {
    return true;
  }
  return contentBlocksShareContainer(baseBlock, revisedBlock);
};

export const detectFolioContentMoves = <Block extends FolioContentBlock>({
  steps,
  consumedStepIndexes,
  workSession,
  idStability = (block): FolioContentIdStability => block.idStability ?? "stable",
}: {
  steps: readonly FolioContentAlignmentStep<Block>[];
  consumedStepIndexes: ReadonlySet<number>;
  workSession: FolioContentComparisonWorkSession;
  idStability?: (block: Block) => FolioContentIdStability;
}): readonly MovePair<Block>[] => {
  const baseOnly: Block[] = [];
  const stableBaseById = new Map<string, Block>();
  const profiles = new Map<string, TokenProfile>();
  const candidatesByText = new Map<string, Block[]>();
  for (const [index, step] of steps.entries()) {
    if (consumedStepIndexes.has(index) || step.type !== "baseOnly") continue;
    if (idStability(step.block) === "stable") {
      stableBaseById.set(step.block.id, step.block);
    }
    const profile = tokenProfile(step.block.text);
    if (!profile) continue;
    baseOnly.push(step.block);
    profiles.set(step.block.id, profile);
    const queue = candidatesByText.get(step.block.text);
    if (!queue) {
      candidatesByText.set(step.block.text, [step.block]);
    } else if (queue.length < MAX_MOVE_CANDIDATES_PER_TEXT) {
      queue.push(step.block);
    }
  }

  const taken = new Set<string>();
  const takenRevised = new Set<string>();
  const moves: MovePair<Block>[] = [];

  // Stable identity is stronger than either text heuristic. Claim every such
  // counterpart before walking revised blocks in order, so an earlier
  // positional candidate cannot steal its source through equal or similar text.
  for (const [index, step] of steps.entries()) {
    if (consumedStepIndexes.has(index) || step.type !== "revisedOnly") continue;
    const stable =
      idStability(step.block) === "stable" ? stableBaseById.get(step.block.id) : undefined;
    if (!stable || taken.has(stable.id)) continue;
    taken.add(stable.id);
    takenRevised.add(step.block.id);
    if (contentBlocksCanMoveTogether(stable, step.block)) {
      moves.push({ baseBlock: stable, revisedBlock: step.block });
    }
  }

  // Exact text wins over edited similarity across the whole stream for the
  // same reason: a merely similar earlier candidate must not consume the only
  // exact source of a later one.
  for (const [index, step] of steps.entries()) {
    if (
      consumedStepIndexes.has(index) ||
      step.type !== "revisedOnly" ||
      takenRevised.has(step.block.id)
    ) {
      continue;
    }
    const exactQueue = candidatesByText.get(step.block.text);
    const exact = exactQueue?.find(
      (candidate) =>
        !taken.has(candidate.id) && contentBlocksCanMoveTogether(candidate, step.block),
    );
    if (exact) {
      taken.add(exact.id);
      takenRevised.add(step.block.id);
      moves.push({ baseBlock: exact, revisedBlock: step.block });
    }
  }

  for (const [index, step] of steps.entries()) {
    if (
      consumedStepIndexes.has(index) ||
      step.type !== "revisedOnly" ||
      takenRevised.has(step.block.id)
    ) {
      continue;
    }
    const revisedProfile = tokenProfile(step.block.text);
    if (!revisedProfile) continue;
    let best: { block: Block; similarity: number } | null = null;
    for (const candidate of baseOnly) {
      if (workSession.remainingMoveComparisons <= 0) break;
      if (taken.has(candidate.id)) continue;
      workSession.remainingMoveComparisons--;
      if (!contentBlocksCanMoveTogether(candidate, step.block)) continue;
      const baseProfile = profiles.get(candidate.id);
      if (!baseProfile) return panic("An eligible move candidate has no token profile");
      const similarity = tokenSimilarity(baseProfile, revisedProfile);
      if (similarity >= MOVE_SIMILARITY_THRESHOLD && (best === null || similarity > best.similarity)) {
        best = { block: candidate, similarity };
      }
    }
    if (best) {
      taken.add(best.block.id);
      takenRevised.add(step.block.id);
      moves.push({ baseBlock: best.block, revisedBlock: step.block });
    }
  }
  return moves;
};

type Relation<Block extends FolioContentBlock> = {
  id: number;
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  event: FolioContentComparisonEvent<Block>;
};

const structuralChangeForStep = <Block extends FolioContentBlock>(
  step: FolioContentAlignmentStep<Block>,
  id: number,
): FolioContentStructuralChange | null => {
  switch (step.type) {
    case "baseTable":
      return { id, type: "table-delete", tableIndex: step.location.tableIndex, baseBlockIds: step.blocks.map(({ id: blockId }) => blockId) };
    case "revisedTable":
      return { id, type: "table-insert", tableIndex: step.location.tableIndex, revisedBlockIds: step.blocks.map(({ id: blockId }) => blockId) };
    case "baseRow":
      return { id, type: "table-row-delete", tableIndex: step.location.tableIndex, rowIndex: step.location.rowIndex, baseBlockIds: step.blocks.map(({ id: blockId }) => blockId) };
    case "revisedRow":
      return { id, type: "table-row-insert", tableIndex: step.location.tableIndex, rowIndex: step.location.rowIndex, revisedBlockIds: step.blocks.map(({ id: blockId }) => blockId) };
    case "baseColumn":
      return { id, type: "table-column-delete", tableIndex: step.location.tableIndex, columnIndex: step.columnIndex, baseBlockIds: step.blocks.map(({ id: blockId }) => blockId) };
    case "revisedColumn":
      return { id, type: "table-column-insert", tableIndex: step.location.tableIndex, columnIndex: step.columnIndex, revisedBlockIds: step.blocks.map(({ id: blockId }) => blockId) };
    default:
      return null;
  }
};

const formattingChange = <Block extends FolioContentBlock>(
  base: Block,
  revised: Block,
  maxRanges: number,
): FolioContentFormattingChange | null | "limit" => {
  const paragraph = changedFolioContentParagraphFormatting(base, revised);
  const ranges =
    base.text === revised.text
      ? inlineFormattingSegments({ baseBlock: base, targetBlock: revised, maxSegments: maxRanges })
      : [];
  if (ranges === null) return "limit";
  return paragraph || ranges.length > 0 ? { ...(paragraph && { paragraph }), ranges } : null;
};

type CompareAlignedContentOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  steps: readonly FolioContentAlignmentStep<Block>[];
  workSession: FolioContentComparisonWorkSession;
  maxChanges: number;
  idStability?: (block: Block) => FolioContentIdStability;
};

export const compareAlignedFolioContent = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  steps,
  workSession,
  maxChanges,
  idStability,
}: CompareAlignedContentOptions<Block>): Result<
  FolioContentComparison<Block>,
  FolioContentComparisonLimitError
> => {
  const paragraphPlans = detectFolioContentParagraphMarkPlans(steps);
  const consumed = new Set([...paragraphPlans.keys()].map((index) => index + 1));
  const moves = detectFolioContentMoves({
    steps,
    consumedStepIndexes: consumed,
    workSession,
    idStability,
  });
  const moveByBaseId = new Map(moves.map((move, index) => [move.baseBlock.id, { ...move, moveId: index + 1 }] as const));
  const moveByRevisedId = new Map(moves.map((move, index) => [move.revisedBlock.id, { ...move, moveId: index + 1 }] as const));
  const { diffText } = workSession;
  const relations: Relation<Block>[] = [];
  const baseRelation = new Map<string, Relation<Block>>();
  const revisedRelation = new Map<string, Relation<Block>>();
  const structuralChanges: FolioContentStructuralChange[] = [];
  let nextRelationId = 0;
  let nextStructuralId = 0;
  let remainingFormattingRanges = maxChanges;

  const addRelation = (
    event: FolioContentComparisonEvent<Block>,
    relationBaseBlocks: readonly Block[],
    relationRevisedBlocks: readonly Block[],
  ): void => {
    const relation = { id: nextRelationId++, baseBlocks: relationBaseBlocks, revisedBlocks: relationRevisedBlocks, event };
    relations.push(relation);
    for (const block of relationBaseBlocks) {
      if (baseRelation.has(block.id)) {
        panic("Content alignment assigned one base block more than once", { blockId: block.id });
      }
      baseRelation.set(block.id, relation);
    }
    for (const block of relationRevisedBlocks) {
      if (revisedRelation.has(block.id)) {
        panic("Content alignment assigned one revised block more than once", { blockId: block.id });
      }
      revisedRelation.set(block.id, relation);
    }
  };

  for (const [stepIndex, step] of steps.entries()) {
    if (consumed.has(stepIndex)) continue;
    const paragraphPlan = paragraphPlans.get(stepIndex);
    if (paragraphPlan?.type === "split") {
      const event = {
        type: "split",
        baseBlocks: [paragraphPlan.baseBlock],
        revisedBlocks: paragraphPlan.revisedBlocks,
        offset: paragraphPlan.offset,
        separator: paragraphPlan.separator,
      } as const;
      addRelation(event, event.baseBlocks, event.revisedBlocks);
      continue;
    }
    if (paragraphPlan?.type === "merge") {
      const event = {
        type: "merge",
        baseBlocks: paragraphPlan.baseBlocks,
        revisedBlocks: [paragraphPlan.revisedBlock],
        separator: paragraphPlan.separator,
      } as const;
      addRelation(event, event.baseBlocks, event.revisedBlocks);
      continue;
    }

    if (step.type === "pair") {
      const base = step.baseBlock;
      const revised = step.revisedBlock;
      const properties = changedBlockProperties(base, revised);
      const formatting = formattingChange(base, revised, remainingFormattingRanges);
      if (formatting === "limit") {
        return Result.err(new FolioContentComparisonLimitError({ message: "The comparison contains more formatting ranges than it returns.", limit: "changes", maximum: maxChanges, actual: maxChanges + 1 }));
      }
      remainingFormattingRanges -= formatting?.ranges.length ?? 0;
      let event: FolioContentComparisonEvent<Block>;
      if (base.text !== revised.text || properties.length > 0) {
        event = {
          type: "modified",
          baseBlocks: [base],
          revisedBlocks: [revised],
          segments: withTextOffsets(diffText(base.text, revised.text)),
          changedProperties: properties,
          ...(formatting && { formatting }),
        };
      } else if (formatting) {
        event = { type: "formatting", baseBlocks: [base], revisedBlocks: [revised], formatting };
      } else {
        event = { type: "unchanged", baseBlocks: [base], revisedBlocks: [revised] };
      }
      addRelation(event, [base], [revised]);
      continue;
    }

    if (step.type === "baseOnly") {
      const move = moveByBaseId.get(step.block.id);
      const event: FolioContentComparisonEvent<Block> = move
        ? { type: "movedFrom", baseBlocks: [step.block], revisedBlocks: [], moveId: move.moveId }
        : { type: "deleted", baseBlocks: [step.block], revisedBlocks: [] };
      addRelation(event, [step.block], []);
      continue;
    }
    if (step.type === "revisedOnly") {
      const move = moveByRevisedId.get(step.block.id);
      const moveFormatting = move
        ? formattingChange(move.baseBlock, step.block, remainingFormattingRanges)
        : null;
      if (moveFormatting === "limit") {
        return Result.err(
          new FolioContentComparisonLimitError({
            message: "The comparison contains more formatting ranges than it returns.",
            limit: "changes",
            maximum: maxChanges,
            actual: maxChanges + 1,
          }),
        );
      }
      remainingFormattingRanges -= moveFormatting?.ranges.length ?? 0;
      const moveProperties = move
        ? changedBlockProperties(move.baseBlock, step.block)
        : [];
      const event: FolioContentComparisonEvent<Block> = move
        ? {
            type: "movedTo",
            baseBlocks: [],
            revisedBlocks: [step.block],
            moveId: move.moveId,
            baseBlockId: move.baseBlock.id,
            ...(move.baseBlock.text !== step.block.text && { segments: withTextOffsets(diffText(move.baseBlock.text, step.block.text)) }),
            ...(moveProperties.length > 0 && { changedProperties: moveProperties }),
            ...(moveFormatting && { formatting: moveFormatting }),
          }
        : { type: "inserted", baseBlocks: [], revisedBlocks: [step.block] };
      addRelation(event, [], [step.block]);
      continue;
    }

    const structural = structuralChangeForStep(step, ++nextStructuralId);
    if (!structural) return panic("Unhandled content alignment step", { step });
    structuralChanges.push(structural);
    if ("baseBlockIds" in structural) {
      for (const block of step.blocks) {
        const event = { type: "deleted", baseBlocks: [block], revisedBlocks: [], structuralChangeId: structural.id } as const;
        addRelation(event, [block], []);
      }
    } else {
      for (const block of step.blocks) {
        const event = { type: "inserted", baseBlocks: [], revisedBlocks: [block], structuralChangeId: structural.id } as const;
        addRelation(event, [], [block]);
      }
    }
  }

  const ordered: FolioContentComparisonEvent<Block>[] = [];
  const emitted = new Set<number>();
  let baseIndex = 0;
  let revisedIndex = 0;
  while (baseIndex < baseBlocks.length || revisedIndex < revisedBlocks.length) {
    const base = baseBlocks[baseIndex];
    const revised = revisedBlocks[revisedIndex];
    const fromBase = base ? baseRelation.get(base.id) : undefined;
    const fromRevised = revised ? revisedRelation.get(revised.id) : undefined;
    if (fromBase && emitted.has(fromBase.id)) {
      baseIndex++;
      continue;
    }
    if (fromRevised && emitted.has(fromRevised.id)) {
      revisedIndex++;
      continue;
    }
    const relation =
      fromBase && fromRevised && fromBase.id === fromRevised.id
        ? fromBase
        : fromBase?.revisedBlocks.length === 0
          ? fromBase
          : fromRevised?.baseBlocks.length === 0
            ? fromRevised
            : undefined;
    if (!relation) {
      return panic("Content alignment did not produce one monotone projection", {
        baseBlockId: base?.id,
        revisedBlockId: revised?.id,
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

  const changeCount = ordered.reduce((count, event) => count + (event.type === "unchanged" ? 0 : 1), 0);
  if (changeCount > maxChanges) {
    return Result.err(new FolioContentComparisonLimitError({
      message: "The comparison contains more changes than one result returns.",
      limit: "changes",
      maximum: maxChanges,
      actual: changeCount,
    }));
  }
  return Result.ok({ events: ordered, structuralChanges });
};

/** Compare two representation-neutral ordered content snapshots. */
export const compareContent = <Kind extends string = string>(
  options: CompareContentOptions<Kind>,
): Result<
  FolioContentComparison<FolioContentBlock<Kind>>,
  FolioContentComparisonError
> => {
  if (!isRecord(options)) {
    return Result.err(
      invalidInput("options", "options", "Comparison options must be an object."),
    );
  }
  const { base, revised, granularity } = options;
  if (
    granularity !== undefined &&
    !WORD_DIFF_GRANULARITIES.some((candidate) => candidate === granularity)
  ) {
    return Result.err(
      invalidInput(
        "options",
        "granularity",
        "Comparison granularity must be word or character.",
      ),
    );
  }
  const baseError = validateSnapshot(base, "base");
  if (baseError) return Result.err(baseError);
  const revisedError = validateSnapshot(revised, "revised");
  if (revisedError) return Result.err(revisedError);
  const workSession = createContentComparisonWorkSession(granularity);
  const steps = alignFolioContentStructure({
    baseBlocks: base.blocks,
    revisedBlocks: revised.blocks,
    workSession: workSession.alignment,
  });
  return compareAlignedFolioContent({
    baseBlocks: base.blocks,
    revisedBlocks: revised.blocks,
    steps,
    workSession,
    maxChanges: MAX_FOLIO_CONTENT_CHANGES,
  });
};
