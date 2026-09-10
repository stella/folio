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

/** Hard resource ceilings for one representation-neutral comparison. */
export const FOLIO_CONTENT_COMPARISON_LIMITS = Object.freeze({
  blocksPerSnapshot: 100_000,
  changes: 10_000,
  blockCodeUnits: 1_048_576,
  textCodeUnitsPerSnapshot: 8_000_000,
  previewRunsPerBlock: 65_536,
  previewRunsPerSnapshot: 1_000_000,
  containerDepth: 64,
  containerEntriesPerSnapshot: 1_000_000,
  attributeCodeUnits: 16_384,
  attributeCodeUnitsPerSnapshot: 8_000_000,
} as const);

export type FolioContentComparisonLimit = keyof typeof FOLIO_CONTENT_COMPARISON_LIMITS;

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
  input: "base" | "revised" | "result";
  limit: FolioContentComparisonLimit;
  maximum: number;
  actual: number;
  blockIndex?: number;
  field?: string;
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
export type CompareContentOptions<
  Block extends FolioContentBlock = FolioContentBlock,
> = {
  base: FolioContentSnapshot<Block>;
  revised: FolioContentSnapshot<Block>;
  /** Token size for modified-block segments; defaults to `"word"`. */
  granularity?: WordDiffGranularity;
};

export type FolioContentComparisonWorkSession = {
  alignment: FolioContentAlignmentWorkSession;
  remainingMoveComparisons: number;
  remainingMoveTokenLookups: number;
  diffText: ReturnType<typeof createWordDiffSession>["diff"];
};

export const createContentComparisonWorkSession = (
  granularity?: WordDiffGranularity,
): FolioContentComparisonWorkSession => ({
  alignment: createFolioContentAlignmentWorkSession(),
  remainingMoveComparisons: MAX_MOVE_SIMILARITY_COMPARISONS,
  remainingMoveTokenLookups: MAX_MOVE_SIMILARITY_TOKEN_LOOKUPS,
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

const limitExceeded = ({
  input,
  limit,
  maximum,
  actual,
  blockIndex,
  field,
}: {
  input: "base" | "revised" | "result";
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

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

// Runtime validation must not replace an already typed generic block with the
// narrower Record<string, unknown> view produced by a type predicate.
const hasRecordShape = (value: unknown): boolean => isRecord(value);

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
  const gridColumnIndex = table["gridColumnIndex"];
  const columnSpan = table["columnSpan"];
  const right =
    typeof gridColumnIndex === "number" && typeof columnSpan === "number"
      ? gridColumnIndex + columnSpan
      : Number.NaN;
  if (!Number.isSafeInteger(right)) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].table.columnSpan`,
      "A table cell's ending grid column must be a safe integer.",
      blockIndex,
    );
  }
  const rowIndex = table["rowIndex"];
  const rowSpan = table["rowSpan"];
  const bottom =
    typeof rowIndex === "number" && typeof rowSpan === "number"
      ? rowIndex + rowSpan
      : Number.NaN;
  if (!Number.isSafeInteger(bottom)) {
    return invalidInput(
      side,
      `blocks[${String(blockIndex)}].table.rowSpan`,
      "A table cell's ending row must be a safe integer.",
      blockIndex,
    );
  }
  return null;
};

type SnapshotResourceUsage = {
  textCodeUnits: number;
  previewRuns: number;
  containerEntries: number;
  attributeCodeUnits: number;
};

type TableCellRectangle = {
  blockIndex: number;
  left: number;
  right: number;
  top: number;
  bottom: number;
};

const overlappingTableCellBlockIndex = (
  cells: readonly TableCellRectangle[],
): number | null => {
  if (cells.length < 2) return null;
  const columns = [
    ...new Set(cells.flatMap(({ left, right }) => [left, right])),
  ].toSorted((left, right) => left - right);
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
      (pending[node] ?? 0) +
      Math.max(maximum[node * 2] ?? 0, maximum[node * 2 + 1] ?? 0);
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
  if (
    usage.attributeCodeUnits >
    FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnitsPerSnapshot
  ) {
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

const validateSnapshot = <Block extends FolioContentBlock>(
  snapshot: FolioContentSnapshot<Block>,
  side: "base" | "revised",
): InvalidFolioContentComparisonError | FolioContentComparisonLimitError | null => {
  if (!snapshot || !Array.isArray(snapshot.blocks)) {
    return invalidInput(side, "blocks", "A content snapshot must contain an ordered blocks array.");
  }
  if (snapshot.blocks.length > FOLIO_CONTENT_COMPARISON_LIMITS.blocksPerSnapshot) {
    return limitExceeded({
      input: side,
      limit: "blocksPerSnapshot",
      maximum: FOLIO_CONTENT_COMPARISON_LIMITS.blocksPerSnapshot,
      actual: snapshot.blocks.length,
    });
  }

  const usage: SnapshotResourceUsage = {
    textCodeUnits: 0,
    previewRuns: 0,
    containerEntries: 0,
    attributeCodeUnits: 0,
  };
  const ids = new Set<string>();
  const lastCoordinateByTable = new Map<string, readonly [number, number, number]>();
  const geometryByCell = new Map<string, readonly [number, number, number]>();
  const cellsByTable = new Map<string, TableCellRectangle[]>();
  const outerTableByTableIndex = new Map<number, number>();
  let lastOuterTableIndex = -1;
  let activeOuterTableIndex: number | null = null;
  for (const [blockIndex, block] of snapshot.blocks.entries()) {
    if (!hasRecordShape(block)) {
      return invalidInput(side, `blocks[${String(blockIndex)}]`, "Every content block must be an object.", blockIndex);
    }
    if (typeof block.id !== "string" || block.id.length === 0) {
      return invalidInput(side, `blocks[${String(blockIndex)}].id`, "Every content block needs a non-empty id.", blockIndex);
    }
    const idLimit = chargeAttributeString({
      value: block.id,
      side,
      blockIndex,
      field: `blocks[${String(blockIndex)}].id`,
      usage,
    });
    if (idLimit) return idLimit;
    if (ids.has(block.id)) {
      return invalidInput(side, `blocks[${String(blockIndex)}].id`, "Content block ids must be unique within a snapshot.", blockIndex);
    }
    ids.add(block.id);
    if (typeof block.kind !== "string" || block.kind.length === 0) {
      return invalidInput(side, `blocks[${String(blockIndex)}].kind`, "Every content block needs a non-empty kind.", blockIndex);
    }
    const kindLimit = chargeAttributeString({
      value: block.kind,
      side,
      blockIndex,
      field: `blocks[${String(blockIndex)}].kind`,
      usage,
    });
    if (kindLimit) return kindLimit;
    if (typeof block.text !== "string") {
      return invalidInput(side, `blocks[${String(blockIndex)}].text`, "Content block text must be a string.", blockIndex);
    }
    if (block.text.length > FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits) {
      return limitExceeded({
        input: side,
        limit: "blockCodeUnits",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits,
        actual: block.text.length,
        blockIndex,
        field: `blocks[${String(blockIndex)}].text`,
      });
    }
    usage.textCodeUnits += block.text.length;
    if (usage.textCodeUnits > FOLIO_CONTENT_COMPARISON_LIMITS.textCodeUnitsPerSnapshot) {
      return limitExceeded({
        input: side,
        limit: "textCodeUnitsPerSnapshot",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.textCodeUnitsPerSnapshot,
        actual: usage.textCodeUnits,
        blockIndex,
        field: `blocks[${String(blockIndex)}].text`,
      });
    }
    if (block.idStability !== undefined && block.idStability !== "stable" && block.idStability !== "positional") {
      return invalidInput(side, `blocks[${String(blockIndex)}].idStability`, "Block id stability must be stable or positional.", blockIndex);
    }
    if (block.previewRuns !== undefined) {
      if (!Array.isArray(block.previewRuns)) {
        return invalidInput(side, `blocks[${String(blockIndex)}].previewRuns`, "Preview runs must be an array of text runs.", blockIndex);
      }
      if (
        block.previewRuns.length >
        FOLIO_CONTENT_COMPARISON_LIMITS.previewRunsPerBlock
      ) {
        return limitExceeded({
          input: side,
          limit: "previewRunsPerBlock",
          maximum: FOLIO_CONTENT_COMPARISON_LIMITS.previewRunsPerBlock,
          actual: block.previewRuns.length,
          blockIndex,
          field: `blocks[${String(blockIndex)}].previewRuns`,
        });
      }
      usage.previewRuns += block.previewRuns.length;
      if (
        usage.previewRuns >
        FOLIO_CONTENT_COMPARISON_LIMITS.previewRunsPerSnapshot
      ) {
        return limitExceeded({
          input: side,
          limit: "previewRunsPerSnapshot",
          maximum: FOLIO_CONTENT_COMPARISON_LIMITS.previewRunsPerSnapshot,
          actual: usage.previewRuns,
          blockIndex,
          field: `blocks[${String(blockIndex)}].previewRuns`,
        });
      }
      let runTextOffset = 0;
      for (const [runIndex, run] of block.previewRuns.entries()) {
        if (!hasRecordShape(run) || typeof run.text !== "string") {
          return invalidInput(side, `blocks[${String(blockIndex)}].previewRuns`, "Preview runs must be an array of text runs.", blockIndex);
        }
        if (!block.text.startsWith(run.text, runTextOffset)) {
          return invalidInput(side, `blocks[${String(blockIndex)}].previewRuns`, "Preview-run text must reconstruct the block text exactly.", blockIndex);
        }
        runTextOffset += run.text.length;
        for (const property of ["fontFamily", "color"] as const) {
          for (const [suffix, value] of [
            [property, run[property]],
            [`directFormatting.${property}`, run.directFormatting?.[property]],
          ] as const) {
            const formattingLimit = chargeAttributeString({
              value: typeof value === "string" ? value : undefined,
              side,
              blockIndex,
              field: `blocks[${String(blockIndex)}].previewRuns[${String(runIndex)}].${suffix}`,
              usage,
            });
            if (formattingLimit) return formattingLimit;
          }
        }
      }
      if (runTextOffset !== block.text.length) {
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
    for (const property of ["styleId", "displayLabel"] as const) {
      const propertyLimit = chargeAttributeString({
        value: block[property],
        side,
        blockIndex,
        field: `blocks[${String(blockIndex)}].${property}`,
        usage,
      });
      if (propertyLimit) return propertyLimit;
    }
    if (block.containerPath !== undefined) {
      if (!Array.isArray(block.containerPath)) {
        return invalidInput(side, `blocks[${String(blockIndex)}].containerPath`, "Container paths require non-empty kind and id values.", blockIndex);
      }
      if (block.containerPath.length > FOLIO_CONTENT_COMPARISON_LIMITS.containerDepth) {
        return limitExceeded({
          input: side,
          limit: "containerDepth",
          maximum: FOLIO_CONTENT_COMPARISON_LIMITS.containerDepth,
          actual: block.containerPath.length,
          blockIndex,
          field: `blocks[${String(blockIndex)}].containerPath`,
        });
      }
      usage.containerEntries += block.containerPath.length;
      if (
        usage.containerEntries >
        FOLIO_CONTENT_COMPARISON_LIMITS.containerEntriesPerSnapshot
      ) {
        return limitExceeded({
          input: side,
          limit: "containerEntriesPerSnapshot",
          maximum: FOLIO_CONTENT_COMPARISON_LIMITS.containerEntriesPerSnapshot,
          actual: usage.containerEntries,
          blockIndex,
          field: `blocks[${String(blockIndex)}].containerPath`,
        });
      }
      for (const [pathIndex, entry] of block.containerPath.entries()) {
        if (
          !hasRecordShape(entry) ||
          typeof entry.kind !== "string" ||
          entry.kind.length === 0 ||
          typeof entry.id !== "string" ||
          entry.id.length === 0
        ) {
          return invalidInput(side, `blocks[${String(blockIndex)}].containerPath`, "Container paths require non-empty kind and id values.", blockIndex);
        }
        for (const property of ["kind", "id"] as const) {
          const pathLimit = chargeAttributeString({
            value: entry[property],
            side,
            blockIndex,
            field: `blocks[${String(blockIndex)}].containerPath[${String(pathIndex)}].${property}`,
            usage,
          });
          if (pathLimit) return pathLimit;
        }
      }
    }
    if (block.table !== undefined) {
      const error = validateTableLocation(block.table, side, blockIndex);
      if (error) {
        return error;
      }
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
      if (
        knownOuterTableIndex !== undefined &&
        knownOuterTableIndex !== table.outerTableIndex
      ) {
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
      const cellKey = `${tableKey}:${String(table.rowIndex)}:${String(table.cellIndex)}`;
      const geometry = [
        table.gridColumnIndex,
        table.columnSpan,
        table.rowSpan,
      ] as const;
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
    } else {
      activeOuterTableIndex = null;
    }
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
  moveScope: Extract<
    FolioContentAlignmentStep<Block>,
    { type: "baseOnly" }
  >["moveScope"];
  profile: TokenProfile;
  order: number;
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
  const stableBaseById = new Map<
    string,
    Pick<MoveCandidate<Block>, "block" | "moveScope">
  >();
  const exactCandidatesByBucket = new Map<number, Map<string, MoveCandidate<Block>[]>>();
  const similarityCandidatesByBucket = new Map<
    number,
    Map<number, Map<string, MoveCandidate<Block>>>
  >();
  let candidateOrder = 0;
  for (const [index, step] of steps.entries()) {
    if (consumedStepIndexes.has(index) || step.type !== "baseOnly") continue;
    if (idStability(step.block) === "stable") {
      stableBaseById.set(step.block.id, {
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
    gapCandidates.set(step.block.id, candidate);
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
    gapCandidates?.delete(candidate.block.id);
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
    const candidate =
      idStability(step.block) === "stable" ? stableBaseById.get(step.block.id) : undefined;
    if (!candidate || taken.has(candidate.block.id)) continue;
    taken.add(candidate.block.id);
    takenRevised.add(step.block.id);
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
      takenRevised.has(step.block.id)
    ) {
      continue;
    }
    const exactQueue = exactCandidatesByBucket
      .get(step.moveScope.bucket)
      ?.get(step.block.text);
    const exact = exactQueue?.find(
      (candidate) =>
        !taken.has(candidate.block.id) && candidate.moveScope.gap !== step.moveScope.gap,
    );
    if (exact) {
      taken.add(exact.block.id);
      takenRevised.add(step.block.id);
      removeSimilarityCandidate(exact);
      moves.push({ baseBlock: exact.block, revisedBlock: step.block });
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
    let best: { candidate: MoveCandidate<Block>; similarity: number } | null = null;
    const candidatesByGap = similarityCandidatesByBucket.get(step.moveScope.bucket);
    candidateGroups: for (const [gap, candidates] of candidatesByGap ?? []) {
      if (gap === step.moveScope.gap) continue;
      for (const candidate of candidates.values()) {
        if (workSession.remainingMoveComparisons <= 0) break candidateGroups;
        workSession.remainingMoveComparisons--;
        const scored = tokenSimilarity(
          candidate.profile,
          revisedProfile,
          workSession.remainingMoveTokenLookups,
        );
        if (!scored) continue;
        workSession.remainingMoveTokenLookups -= scored.lookups;
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
      taken.add(best.candidate.block.id);
      takenRevised.add(step.block.id);
      removeSimilarityCandidate(best.candidate);
      moves.push({ baseBlock: best.candidate.block, revisedBlock: step.block });
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
    ...(idStability && { idStability }),
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
  let changeCount = 0;

  const addRelation = (
    event: FolioContentComparisonEvent<Block>,
    relationBaseBlocks: readonly Block[],
    relationRevisedBlocks: readonly Block[],
  ): boolean => {
    if (event.type !== "unchanged") {
      changeCount++;
      if (changeCount > maxChanges) return false;
    }
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
    return true;
  };

  const changeLimitExceeded = (): Result<never, FolioContentComparisonLimitError> =>
    Result.err(
      limitExceeded({
        input: "result",
        limit: "changes",
        maximum: maxChanges,
        actual: maxChanges + 1,
      }),
    );

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
      if (!addRelation(event, event.baseBlocks, event.revisedBlocks)) {
        return changeLimitExceeded();
      }
      continue;
    }
    if (paragraphPlan?.type === "merge") {
      const event = {
        type: "merge",
        baseBlocks: paragraphPlan.baseBlocks,
        revisedBlocks: [paragraphPlan.revisedBlock],
        separator: paragraphPlan.separator,
      } as const;
      if (!addRelation(event, event.baseBlocks, event.revisedBlocks)) {
        return changeLimitExceeded();
      }
      continue;
    }

    if (step.type === "pair") {
      const base = step.baseBlock;
      const revised = step.revisedBlock;
      const properties = changedBlockProperties(base, revised);
      const formatting = formattingChange(base, revised, remainingFormattingRanges);
      if (formatting === "limit") {
        return Result.err(
          limitExceeded({
            input: "result",
            limit: "changes",
            maximum: maxChanges,
            actual: maxChanges + 1,
          }),
        );
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
      if (!addRelation(event, [base], [revised])) {
        return changeLimitExceeded();
      }
      continue;
    }

    if (step.type === "baseOnly") {
      const move = moveByBaseId.get(step.block.id);
      const event: FolioContentComparisonEvent<Block> = move
        ? { type: "movedFrom", baseBlocks: [step.block], revisedBlocks: [], moveId: move.moveId }
        : { type: "deleted", baseBlocks: [step.block], revisedBlocks: [] };
      if (!addRelation(event, [step.block], [])) {
        return changeLimitExceeded();
      }
      continue;
    }
    if (step.type === "revisedOnly") {
      const move = moveByRevisedId.get(step.block.id);
      const moveFormatting = move
        ? formattingChange(move.baseBlock, step.block, remainingFormattingRanges)
        : null;
      if (moveFormatting === "limit") {
        return Result.err(
          limitExceeded({
            input: "result",
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
      if (!addRelation(event, [], [step.block])) {
        return changeLimitExceeded();
      }
      continue;
    }

    const structural = structuralChangeForStep(step, ++nextStructuralId);
    if (!structural) return panic("Unhandled content alignment step", { step });
    structuralChanges.push(structural);
    if ("baseBlockIds" in structural) {
      for (const block of step.blocks) {
        const event = { type: "deleted", baseBlocks: [block], revisedBlocks: [], structuralChangeId: structural.id } as const;
        if (!addRelation(event, [block], [])) {
          return changeLimitExceeded();
        }
      }
    } else {
      for (const block of step.blocks) {
        const event = { type: "inserted", baseBlocks: [], revisedBlocks: [block], structuralChangeId: structural.id } as const;
        if (!addRelation(event, [], [block])) {
          return changeLimitExceeded();
        }
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
    let relation: Relation<Block> | undefined;
    if (fromBase && fromRevised && fromBase.id === fromRevised.id) {
      relation = fromBase;
    } else if (fromBase?.revisedBlocks.length === 0) {
      relation = fromBase;
    } else if (fromRevised?.baseBlocks.length === 0) {
      relation = fromRevised;
    }
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

  return Result.ok({ events: ordered, structuralChanges });
};

/** Compare two representation-neutral ordered content snapshots. */
export const compareContent = <Block extends FolioContentBlock = FolioContentBlock>(
  options: CompareContentOptions<Block>,
): Result<
  FolioContentComparison<Block>,
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
    maxChanges: FOLIO_CONTENT_COMPARISON_LIMITS.changes,
  });
};
