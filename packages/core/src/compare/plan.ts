/**
 * Turn two block snapshots of one story into the change list a caller reads
 * and the {@link FolioAIEditOperation}s that reproduce the target when every
 * generated tracked change is accepted.
 *
 * Pure: no parsing, no serialization, no clock. Every alignment runs through
 * {@link alignFolioBlocks}, shared with the version-diff and redline paths.
 *
 * ## Aligning by container, not by block
 *
 * A single alignment over every paragraph in the story cannot see structure.
 * It pairs on text and document order, so it will put a paragraph inside a
 * table cell opposite one outside it, or a cell of one row opposite a cell of
 * the next — and rewriting either pair in place leaves the target's text in
 * the wrong container. A whole row that was added or removed likewise arrives
 * as a scatter of unmatched paragraphs, which no block operation can turn back
 * into a row: a block insertion anchored in a table lands beside the table,
 * never as a new row in it.
 *
 * So the story is aligned in three nested passes, each over things that can
 * actually stand in for one another:
 *
 * 1. **Segments.** Maximal runs of body text and of one table each, so a table
 *    is only ever compared against a table.
 * 2. **Rows**, within a paired table segment, so a whole-row change stays whole
 *    and becomes `insertTableRow` / `deleteTableRow`.
 * 3. **Cells**, within a paired row, matched by physical cell index.
 *
 * ## Move detection
 *
 * A base-only and a target-only block with identical text and at least
 * {@link MOVE_MINIMUM_WORD_COUNT} words are one relocation. The operations stay
 * a delete plus an insert — the tracked-change grammar Word round-trips has no
 * durable "moved from here" mark on this path — but the change list reports a
 * single `move` so the relocation is not read as unrelated churn. The word
 * floor keeps boilerplate one-liners from pairing as spurious moves.
 */

import { panic } from "better-result";

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import { createFolioAITextRangeHandle, trailingBodyBlockId } from "../ai-edits/snapshot";
import type {
  FolioAIBlock,
  FolioAIBlockParagraphProperties,
  FolioAIBlockTableLocation,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "../ai-edits/types";
import { alignFolioBlocks } from "../version-comparison";
import { inlineFormattingSegments } from "./formatting";
import type { CompareChange, CompareChangeLocation } from "./types";

/** Words a relocated block needs before the move pass will pair it. */
const MOVE_MINIMUM_WORD_COUNT = 3;

/**
 * Cap on same-text base-only blocks the move pass keeps per text. Without it a
 * document repeating one paragraph thousands of times would make the pass
 * quadratic on attacker-controlled input.
 */
const MAX_MOVE_CANDIDATES_PER_TEXT = 64;

/**
 * How much of a relocated paragraph must survive the relocation for it still
 * to read as one. Below it the two paragraphs are a deletion and an unrelated
 * insertion, and calling them a move would tell the reader the wrong story
 * about where the text came from.
 */
const MOVE_SIMILARITY_THRESHOLD = 0.8;

/**
 * Total pair comparisons the similarity pass may make in one story. Exact text
 * matches are found by lookup; only what is left pays this, and it is capped
 * so two documents of unmatched paragraphs cannot make the pass quadratic.
 */
const MAX_MOVE_SIMILARITY_COMPARISONS = 20_000;

/**
 * Dice coefficient over word tokens: twice the shared tokens over the two
 * token counts. Multiset rather than set, so a paragraph repeating a word does
 * not match one that says it once.
 */
const tokenSimilarity = (left: string, right: string): number => {
  const leftTokens = left.split(/\s+/u).filter((token) => token.length > 0);
  const rightTokens = right.split(/\s+/u).filter((token) => token.length > 0);
  if (leftTokens.length === 0 || rightTokens.length === 0) {
    return 0;
  }
  const remaining = new Map<string, number>();
  for (const token of leftTokens) {
    remaining.set(token, (remaining.get(token) ?? 0) + 1);
  }
  let shared = 0;
  for (const token of rightTokens) {
    const count = remaining.get(token) ?? 0;
    if (count > 0) {
      remaining.set(token, count - 1);
      shared += 1;
    }
  }
  return (2 * shared) / (leftTokens.length + rightTokens.length);
};

const wordCountReaches = (text: string, minimum: number): boolean => {
  let count = 0;
  for (const word of text.split(/\s+/u)) {
    if (word.length > 0 && ++count >= minimum) {
      return true;
    }
  }
  return false;
};

/**
 * One step of the interpreted alignment. `baseRow` / `targetRow` are whole-row
 * changes; the rest are block-level.
 */
type CompareStep =
  | { type: "pair"; baseBlock: FolioAIBlock; targetBlock: FolioAIBlock }
  | { type: "baseOnly"; block: FolioAIBlock }
  | { type: "targetOnly"; block: FolioAIBlock }
  | { type: "baseRow"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | { type: "targetRow"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | { type: "baseTable"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | { type: "targetTable"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation };

/**
 * A maximal run of blocks that share a container: body text, or one table.
 * Blocks arrive in document order and a table's blocks are contiguous, so a
 * single forward scan recovers the document's structure.
 */
type DocumentSegment =
  | { kind: "body"; blocks: FolioAIBlock[] }
  | { kind: "table"; blocks: FolioAIBlock[] };

/**
 * The document's segments, normalized to strictly alternate body, table, body,
 * ... and to begin and end with a body segment — inserting empty body segments
 * where the document has none.
 *
 * Without the padding the two sides can hold different segment shapes (a
 * document that opens with a table against one that opens with a heading), and
 * the alignment's positional fallback would then put a body segment opposite a
 * table. That pair is unusable, so the whole table would be reissued as a
 * deletion plus an insertion. Padding keeps like opposite like.
 */
const splitSegments = (blocks: readonly FolioAIBlock[]): DocumentSegment[] => {
  const segments: DocumentSegment[] = [];
  let currentTableIndex: number | null = null;
  for (const block of blocks) {
    // The OUTERMOST table, so a table inside a cell stays part of its parent's
    // segment instead of splitting it in three.
    const tableIndex = block.table?.outerTableIndex ?? null;
    const current = segments.at(-1);
    if (current !== undefined && currentTableIndex === tableIndex) {
      current.blocks.push(block);
      continue;
    }
    if (block.table && segments.at(-1)?.kind !== "body") {
      segments.push({ kind: "body", blocks: [] });
    }
    segments.push({ kind: block.table ? "table" : "body", blocks: [block] });
    currentTableIndex = tableIndex;
  }
  if (segments.length === 0 || segments.at(-1)?.kind === "table") {
    segments.push({ kind: "body", blocks: [] });
  }
  return segments;
};

/**
 * Blocks of one table segment grouped into rows, in document order.
 *
 * Keyed by table AND row, not by row alone: a segment holds the whole
 * outermost table, so a nested table's first row would otherwise merge with
 * its parent's first row and the alignment would compare one against the
 * other.
 */
const groupRows = (blocks: readonly FolioAIBlock[]): FolioAIBlock[][] => {
  const rows = new Map<string, FolioAIBlock[]>();
  for (const block of blocks) {
    if (!block.table) {
      continue;
    }
    const key = `${String(block.table.tableIndex)}:${String(block.table.rowIndex)}`;
    const row = rows.get(key);
    if (row) {
      row.push(block);
    } else {
      rows.set(key, [block]);
    }
  }
  return [...rows.values()];
};

const rowText = (row: readonly FolioAIBlock[]): string => row.map(({ text }) => text).join(" ");

const rowLocation = (row: readonly FolioAIBlock[]): FolioAIBlockTableLocation | null =>
  row.at(0)?.table ?? null;

/** Zip one paired row's cells by physical cell index, then by paragraph order. */
const alignRowCells = (
  baseRow: readonly FolioAIBlock[],
  targetRow: readonly FolioAIBlock[],
): CompareStep[] => {
  const byCell = (row: readonly FolioAIBlock[]): Map<number, FolioAIBlock[]> => {
    const cells = new Map<number, FolioAIBlock[]>();
    for (const block of row) {
      const cellIndex = block.table?.cellIndex ?? 0;
      const blocks = cells.get(cellIndex);
      if (blocks) {
        blocks.push(block);
      } else {
        cells.set(cellIndex, [block]);
      }
    }
    return cells;
  };

  const baseCells = byCell(baseRow);
  const targetCells = byCell(targetRow);
  const cellIndexes = [...new Set([...baseCells.keys(), ...targetCells.keys()])].toSorted(
    (left, right) => left - right,
  );

  const steps: CompareStep[] = [];
  for (const cellIndex of cellIndexes) {
    const baseBlocks = baseCells.get(cellIndex) ?? [];
    const targetBlocks = targetCells.get(cellIndex) ?? [];
    const paired = Math.min(baseBlocks.length, targetBlocks.length);
    for (let index = 0; index < paired; index++) {
      const baseBlock = baseBlocks[index];
      const targetBlock = targetBlocks[index];
      if (baseBlock && targetBlock) {
        steps.push({ type: "pair", baseBlock, targetBlock });
      }
    }
    for (const block of baseBlocks.slice(paired)) {
      steps.push({ type: "baseOnly", block });
    }
    for (const block of targetBlocks.slice(paired)) {
      steps.push({ type: "targetOnly", block });
    }
  }
  return steps;
};

/**
 * Align one table's rows, then its cells.
 *
 * Running the block alignment straight over a table's paragraphs cannot see
 * rows: it happily pairs a cell of one row with a cell of another, and a row
 * that was wholly added or removed shows up as a scatter of unmatched
 * paragraphs. Aligning rows first — each row standing in as one proxy block of
 * its joined cell text — keeps a whole-row change whole, and confines every
 * other difference to a cell that really corresponds.
 */
/**
 * One segment's blocks split per table, in first-appearance order. A segment
 * is a whole outermost table, so it holds the parent's blocks and every nested
 * table's; a row alignment that mixed them would compare the parent's first
 * row against a nested table's.
 */
const groupTables = (blocks: readonly FolioAIBlock[]): FolioAIBlock[][] => {
  const tables = new Map<number, FolioAIBlock[]>();
  for (const block of blocks) {
    if (!block.table) {
      continue;
    }
    const existing = tables.get(block.table.tableIndex);
    if (existing) {
      existing.push(block);
    } else {
      tables.set(block.table.tableIndex, [block]);
    }
  }
  return [...tables.values()];
};

/**
 * Align a paired table segment: each table in it against the table at the same
 * place on the other side, then that table's rows, then its cells. Tables are
 * paired by order within the segment because the segment IS one table plus
 * whatever nests inside it — the nth nested table of one answers to the nth of
 * the other.
 */
const buildTableSegmentSteps = (
  baseBlocks: readonly FolioAIBlock[],
  targetBlocks: readonly FolioAIBlock[],
): CompareStep[] => {
  const baseTables = groupTables(baseBlocks);
  const targetTables = groupTables(targetBlocks);
  const steps: CompareStep[] = [];
  const paired = Math.min(baseTables.length, targetTables.length);
  for (let index = 0; index < paired; index++) {
    steps.push(...buildTableSteps(baseTables[index] ?? [], targetTables[index] ?? []));
  }
  for (const blocks of baseTables.slice(paired)) {
    const location = blocks.at(0)?.table;
    if (location) {
      steps.push({ type: "baseTable", blocks, location });
    }
  }
  for (const blocks of targetTables.slice(paired)) {
    const location = blocks.at(0)?.table;
    if (location) {
      steps.push({ type: "targetTable", blocks, location });
    }
  }
  return steps;
};

/**
 * How alike two rows must be for one to be read as the other, edited. Below
 * it the pair is a deleted row and an inserted one.
 */
const ROW_PAIR_SIMILARITY = 0.5;

/**
 * How alike two rows are: their text, halved when their shapes differ.
 *
 * A row's shape is its physical cell count. Two rows with the same words in a
 * different number of cells are not the same row edited, and pairing them
 * would report every cell as changed rather than the row as replaced.
 */
const rowSimilarity = (
  base: readonly FolioAIBlock[] | undefined,
  target: readonly FolioAIBlock[] | undefined,
): number => {
  if (!base || !target) {
    return 0;
  }
  const text = tokenSimilarity(rowText(base), rowText(target));
  return rowCellTexts(base).length === rowCellTexts(target).length ? text : text / 2;
};

/**
 * Align one table's rows.
 *
 * Exact text first, then SIMILARITY rather than position. Positional fallback
 * is what made a deleted row plus a few cell edits report as a change in every
 * row of the table: each row was paired with the one below it, so every cell
 * differed. Pairing on similarity, and stepping one row on the side whose next
 * row matches better, keeps the deletion where it happened.
 */
const alignTableRows = (
  baseRows: readonly FolioAIBlock[][],
  targetRows: readonly FolioAIBlock[][],
): CompareStep[] => {
  const steps: CompareStep[] = [];
  const pushRow = (row: readonly FolioAIBlock[], side: "base" | "target"): void => {
    const location = rowLocation(row);
    if (location) {
      steps.push({ type: side === "base" ? "baseRow" : "targetRow", blocks: row, location });
    }
  };

  let baseCursor = 0;
  let targetCursor = 0;
  while (baseCursor < baseRows.length && targetCursor < targetRows.length) {
    const baseRow = baseRows[baseCursor];
    const targetRow = targetRows[targetCursor];
    if (!baseRow || !targetRow) {
      break;
    }
    const here = rowSimilarity(baseRow, targetRow);
    if (here < 1) {
      const baseAhead = rowSimilarity(baseRows[baseCursor + 1], targetRow);
      const targetAhead = rowSimilarity(baseRow, targetRows[targetCursor + 1]);
      if (baseAhead >= ROW_PAIR_SIMILARITY && baseAhead > here && baseAhead >= targetAhead) {
        pushRow(baseRow, "base");
        baseCursor += 1;
        continue;
      }
      if (targetAhead >= ROW_PAIR_SIMILARITY && targetAhead > here) {
        pushRow(targetRow, "target");
        targetCursor += 1;
        continue;
      }
    }
    steps.push(...alignRowCells(baseRow, targetRow));
    baseCursor += 1;
    targetCursor += 1;
  }
  for (const row of baseRows.slice(baseCursor)) {
    pushRow(row, "base");
  }
  for (const row of targetRows.slice(targetCursor)) {
    pushRow(row, "target");
  }
  return steps;
};

const buildTableSteps = (
  baseBlocks: readonly FolioAIBlock[],
  targetBlocks: readonly FolioAIBlock[],
): CompareStep[] => alignTableRows(groupRows(baseBlocks), groupRows(targetBlocks));

const buildBodySteps = (
  baseBlocks: readonly FolioAIBlock[],
  targetBlocks: readonly FolioAIBlock[],
): CompareStep[] =>
  alignFolioBlocks(baseBlocks, targetBlocks).map((event) => {
    switch (event.type) {
      case "pair":
        return { type: "pair", baseBlock: event.baseBlock, targetBlock: event.revisedBlock };
      case "baseOnly":
        return { type: "baseOnly", block: event.block };
      case "revisedOnly":
        return { type: "targetOnly", block: event.block };
      default: {
        const unreachable: never = event;
        return panic("Unhandled block alignment event", { event: unreachable });
      }
    }
  });

/**
 * Every block of an unpaired segment, as one-sided steps. A whole table stays
 * whole: reissuing it row by row would need a table to put the rows in, and
 * the point of an unpaired table segment is that there is none.
 */
const unpairedSegmentSteps = (segment: DocumentSegment, side: "base" | "target"): CompareStep[] => {
  if (segment.kind !== "table") {
    return segment.blocks.map((block) =>
      side === "base" ? { type: "baseOnly", block } : { type: "targetOnly", block },
    );
  }
  const location = segment.blocks.at(0)?.table;
  if (!location) {
    return [];
  }
  return [
    { type: side === "base" ? "baseTable" : "targetTable", blocks: segment.blocks, location },
  ];
};

type BuildStepsOptions = {
  baseBlocks: readonly FolioAIBlock[];
  targetBlocks: readonly FolioAIBlock[];
};

/**
 * Align the two stories segment by segment, in order.
 *
 * Segments are paired by position rather than by text. Both sides have been
 * normalized to the same alternating shape, so position already carries the
 * meaning: the nth table of one document answers to the nth table of the
 * other, and the body text between two tables answers to the body text between
 * the same two tables. Matching segments on text instead lets a paragraph that
 * moved across a table steal the table's own pairing, which reissues the whole
 * table as a deletion and an insertion.
 *
 * When the two documents hold different numbers of tables the shapes diverge,
 * and the surplus segments are reported one-sided. The operation builder turns
 * those segments into `insertTable` or `deleteTable` operations.
 */
const segmentText = (segment: DocumentSegment): string =>
  segment.blocks.map(({ text }) => text).join(" ");

/**
 * How alike two table segments must be before a lookahead match may steal the
 * pairing from the table in front of it. Two tables drawn from one document's
 * vocabulary score alike by chance, so a lookahead has to clear this bar as
 * well as beat what it displaces.
 */
const TABLE_PAIR_SIMILARITY = 0.5;

/**
 * Pair the two segment sequences, marking the segments only one side has.
 *
 * Pairing by index cannot see a table added or removed: the nth table of one
 * document is put opposite the nth of the other, so deleting the first table
 * shifts every later one and the comparison rewrites each table's contents
 * into the next table along. Segments strictly alternate body, table, body,
 * ..., so a one-segment lookahead on each side is enough to tell "this table
 * changed a lot" from "this table is gone": the next table on the other side
 * matching better is what says the current one is unpaired.
 */
const alignSegments = (
  baseSegments: readonly DocumentSegment[],
  targetSegments: readonly DocumentSegment[],
): { baseSegment: DocumentSegment | null; targetSegment: DocumentSegment | null }[] => {
  const paired: { baseSegment: DocumentSegment | null; targetSegment: DocumentSegment | null }[] =
    [];
  const similarity = (
    left: DocumentSegment | undefined,
    right: DocumentSegment | undefined,
  ): number =>
    left === undefined || right === undefined || left.kind !== right.kind
      ? 0
      : tokenSimilarity(segmentText(left), segmentText(right));

  let baseCursor = 0;
  let targetCursor = 0;
  while (baseCursor < baseSegments.length && targetCursor < targetSegments.length) {
    const baseSegment = baseSegments[baseCursor];
    const targetSegment = targetSegments[targetCursor];
    if (!baseSegment || !targetSegment) {
      break;
    }
    // Body segments always answer to body segments: they are the text between
    // two tables, and the block alignment inside them handles the rest.
    if (baseSegment.kind !== targetSegment.kind) {
      paired.push({ baseSegment, targetSegment: null });
      baseCursor += 1;
      continue;
    }
    if (baseSegment.kind === "body") {
      paired.push({ baseSegment, targetSegment });
      baseCursor += 1;
      targetCursor += 1;
      continue;
    }
    // The segment two along is the next one of the same kind: the sequence
    // alternates body, table, body. A lookahead only wins when it is both a
    // real match and a better one — two tables drawn from the same vocabulary
    // score alike by chance, and "better than nothing" is not evidence that
    // this table is gone.
    const here = similarity(baseSegment, targetSegment);
    const baseAhead = similarity(baseSegments[baseCursor + 2], targetSegment);
    const targetAhead = similarity(baseSegment, targetSegments[targetCursor + 2]);
    if (baseAhead >= TABLE_PAIR_SIMILARITY && baseAhead > here && baseAhead >= targetAhead) {
      paired.push({ baseSegment, targetSegment: null });
      baseCursor += 1;
      continue;
    }
    if (targetAhead >= TABLE_PAIR_SIMILARITY && targetAhead > here) {
      paired.push({ baseSegment: null, targetSegment });
      targetCursor += 1;
      continue;
    }
    paired.push({ baseSegment, targetSegment });
    baseCursor += 1;
    targetCursor += 1;
  }
  for (const segment of baseSegments.slice(baseCursor)) {
    paired.push({ baseSegment: segment, targetSegment: null });
  }
  for (const segment of targetSegments.slice(targetCursor)) {
    paired.push({ baseSegment: null, targetSegment: segment });
  }
  return paired;
};

const buildSteps = ({ baseBlocks, targetBlocks }: BuildStepsOptions): CompareStep[] => {
  const steps: CompareStep[] = [];
  for (const { baseSegment, targetSegment } of alignSegments(
    splitSegments(baseBlocks),
    splitSegments(targetBlocks),
  )) {
    if (baseSegment && targetSegment) {
      steps.push(
        ...(baseSegment.kind === "table"
          ? buildTableSegmentSteps(baseSegment.blocks, targetSegment.blocks)
          : buildBodySteps(baseSegment.blocks, targetSegment.blocks)),
      );
      continue;
    }
    if (baseSegment) {
      steps.push(...unpairedSegmentSteps(baseSegment, "base"));
      continue;
    }
    if (targetSegment) {
      steps.push(...unpairedSegmentSteps(targetSegment, "target"));
    }
  }
  return steps;
};

/**
 * Two blocks are in the same container when a paragraph mark between them
 * exists at all: two body paragraphs, or two paragraphs of one table cell. A
 * mark cannot span a cell boundary, so a split or a merge across one is not a
 * paragraph-mark edit however similar the text looks.
 */
const shareAContainer = (left: FolioAIBlock, right: FolioAIBlock): boolean => {
  if (!left.table || !right.table) {
    return left.table === undefined && right.table === undefined;
  }
  return (
    left.table.tableIndex === right.table.tableIndex &&
    left.table.rowIndex === right.table.rowIndex &&
    left.table.cellIndex === right.table.cellIndex
  );
};

/**
 * The text between `head` and `tail` inside `whole`, when `whole` is exactly
 * the two joined by whitespace (or by nothing). `null` when it is not: any
 * other difference is a rewrite, not a moved paragraph mark.
 */
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

/** A split or a merge, and the step it consumed alongside the paired one. */
type ParagraphMarkPlan =
  | {
      type: "split";
      baseBlock: FolioAIBlock;
      targetBlocks: readonly [FolioAIBlock, FolioAIBlock];
      offset: number;
      separator: string;
    }
  | {
      type: "merge";
      baseBlocks: readonly [FolioAIBlock, FolioAIBlock];
      targetBlock: FolioAIBlock;
      separator: string;
    };

/**
 * Where the alignment produced a rewrite plus an insertion or a deletion that
 * is really one paragraph mark moving, by step index of the PAIR step. The
 * step after it is consumed with it.
 *
 * The alignment cannot see this: it pairs the base paragraph with the target
 * half that still matches it and leaves the other half unpaired, which is a
 * correct alignment and a misleading redline.
 */
const detectParagraphMarkEdits = (
  steps: readonly CompareStep[],
): ReadonlyMap<number, ParagraphMarkPlan> => {
  const plans = new Map<number, ParagraphMarkPlan>();
  for (const [index, step] of steps.entries()) {
    const next = steps[index + 1];
    if (step.type !== "pair" || next === undefined) {
      continue;
    }
    if (next.type === "targetOnly") {
      const separator = separatorBetween(
        step.baseBlock.text,
        step.targetBlock.text,
        next.block.text,
      );
      if (separator !== null && shareAContainer(step.targetBlock, next.block)) {
        plans.set(index, {
          type: "split",
          baseBlock: step.baseBlock,
          targetBlocks: [step.targetBlock, next.block],
          offset: step.targetBlock.text.length,
          separator,
        });
      }
      continue;
    }
    if (next.type !== "baseOnly") {
      continue;
    }
    const separator = separatorBetween(step.targetBlock.text, step.baseBlock.text, next.block.text);
    if (separator !== null && shareAContainer(step.baseBlock, next.block)) {
      plans.set(index, {
        type: "merge",
        baseBlocks: [step.baseBlock, next.block],
        targetBlock: step.targetBlock,
        separator,
      });
    }
  }
  return plans;
};

/** Base block id -> target block id for every relocation the move pass found. */
const detectMoves = (
  steps: readonly CompareStep[],
  consumed: ReadonlySet<number>,
): ReadonlyMap<string, string> => {
  const candidatesByText = new Map<string, string[]>();
  for (const [index, step] of steps.entries()) {
    if (
      consumed.has(index) ||
      step.type !== "baseOnly" ||
      !wordCountReaches(step.block.text, MOVE_MINIMUM_WORD_COUNT)
    ) {
      continue;
    }
    const queue = candidatesByText.get(step.block.text);
    if (!queue) {
      candidatesByText.set(step.block.text, [step.block.id]);
      continue;
    }
    if (queue.length < MAX_MOVE_CANDIDATES_PER_TEXT) {
      queue.push(step.block.id);
    }
  }

  const unmatched: FolioAIBlock[] = [];
  for (const [index, step] of steps.entries()) {
    if (
      consumed.has(index) ||
      step.type !== "baseOnly" ||
      !wordCountReaches(step.block.text, MOVE_MINIMUM_WORD_COUNT)
    ) {
      continue;
    }
    unmatched.push(step.block);
  }

  const movesByBaseBlockId = new Map<string, string>();
  const takenBaseBlockIds = new Set<string>();
  let comparisonBudget = MAX_MOVE_SIMILARITY_COMPARISONS;
  for (const [index, step] of steps.entries()) {
    if (consumed.has(index) || step.type !== "targetOnly") {
      continue;
    }
    const exact = candidatesByText.get(step.block.text)?.shift();
    if (exact !== undefined) {
      takenBaseBlockIds.add(exact);
      movesByBaseBlockId.set(exact, step.block.id);
      continue;
    }
    // A relocated paragraph is often edited on the way. Exact text is the
    // fast path; everything else pays a bounded similarity pass, because a
    // document repeating one paragraph thousands of times would otherwise
    // make this quadratic on attacker-controlled input.
    if (!wordCountReaches(step.block.text, MOVE_MINIMUM_WORD_COUNT)) {
      continue;
    }
    let best: { block: FolioAIBlock; similarity: number } | null = null;
    for (const candidate of unmatched) {
      if (comparisonBudget <= 0) {
        break;
      }
      if (takenBaseBlockIds.has(candidate.id)) {
        continue;
      }
      comparisonBudget -= 1;
      const similarity = tokenSimilarity(candidate.text, step.block.text);
      if (
        similarity >= MOVE_SIMILARITY_THRESHOLD &&
        (best === null || similarity > best.similarity)
      ) {
        best = { block: candidate, similarity };
      }
    }
    if (best) {
      takenBaseBlockIds.add(best.block.id);
      movesByBaseBlockId.set(best.block.id, step.block.id);
    }
  }
  return movesByBaseBlockId;
};

/**
 * For each step, the id of the next base block at or after it — the anchor a
 * target-only insertion is placed before. `null` once no base block follows.
 */
const nextBaseBlockIdByStep = (steps: readonly CompareStep[]): (string | null)[] => {
  const anchors = Array.from<string | null>({ length: steps.length });
  let next: string | null = null;
  for (let index = steps.length - 1; index >= 0; index--) {
    anchors[index] = next;
    const step = steps[index];
    if (step?.type === "pair") {
      next = step.baseBlock.id;
    } else if (step?.type === "baseOnly") {
      next = step.block.id;
    } else if (step?.type === "baseRow") {
      next = step.blocks[0]?.id ?? next;
    }
  }
  return anchors;
};

/** A base block sitting inside a table, and whether it precedes or follows the step. */
type RowAnchor = { blockId: string; position: "after" | "before" };

const baseTableBlockOf = (step: CompareStep): FolioAIBlock | null => {
  switch (step.type) {
    case "pair":
      return step.baseBlock.table ? step.baseBlock : null;
    case "baseOnly":
      return step.block.table ? step.block : null;
    case "baseRow":
    case "baseTable":
      return step.blocks[0] ?? null;
    case "targetOnly":
    case "targetRow":
    case "targetTable":
      return null;
    default: {
      const unreachable: never = step;
      return panic("Unhandled compare step", { step: unreachable });
    }
  }
};

/**
 * The base-document row a new row is inserted next to: the nearest base block
 * inside a table, preferring the one before the insertion so a run of new rows
 * keeps its order.
 */
const findRowAnchor = (steps: readonly CompareStep[], stepIndex: number): RowAnchor | null => {
  for (let index = stepIndex - 1; index >= 0; index--) {
    const step = steps[index];
    const block = step ? baseTableBlockOf(step) : null;
    if (block) {
      return { blockId: block.id, position: "after" };
    }
  }
  for (let index = stepIndex + 1; index < steps.length; index++) {
    const step = steps[index];
    const block = step ? baseTableBlockOf(step) : null;
    if (block) {
      return { blockId: block.id, position: "before" };
    }
  }
  return null;
};

/**
 * A row's text per physical cell, indexed BY cell so an empty cell keeps its
 * slot. An empty cell carries no block at all, so packing only the cells that
 * have text would shift every later cell one column left.
 */
/** One table's cell texts, row by row, for a whole-table change. */
/**
 * One table's cell texts, row by row, padded to the widest row.
 *
 * A row's own width is its highest occupied cell index, so a table with
 * merged cells or a short last row produces a ragged grid — and a ragged grid
 * is not a table any consumer can lay out. Padding states the grid the table
 * actually occupies; the empty strings are the cells a `w:gridSpan` covers.
 */
const tableCellTexts = (blocks: readonly FolioAIBlock[]): string[][] => {
  const rows = groupRows(blocks).map((row) => rowCellTexts(row));
  let width = 0;
  for (const row of rows) {
    width = Math.max(width, row.length);
  }
  for (const row of rows) {
    while (row.length < width) {
      row.push("");
    }
  }
  return rows;
};

const rowCellTexts = (blocks: readonly FolioAIBlock[]): string[] => {
  const byCell: (string | undefined)[] = [];
  for (const block of blocks) {
    const cellIndex = block.table?.cellIndex ?? 0;
    const existing = byCell[cellIndex];
    byCell[cellIndex] = existing === undefined ? block.text : `${existing}\n${block.text}`;
  }
  return Array.from(byCell, (text) => text ?? "");
};

/**
 * The paragraph properties that differ, or `null` when they agree. Only the
 * ones a block projection can see and an operation can set: a list level and
 * a paragraph style, the two edits that move no words and are invisible in a
 * text diff.
 */
const changedParagraphProperties = (
  baseBlock: FolioAIBlock,
  targetBlock: FolioAIBlock,
): FolioAIBlockParagraphProperties | null => {
  const properties: FolioAIBlockParagraphProperties = {};
  if ((baseBlock.styleId ?? null) !== (targetBlock.styleId ?? null)) {
    properties.styleId = targetBlock.styleId ?? null;
  }
  // `null` when the target's paragraph carries no numbering at all: a list
  // item that stopped being one moves no words, and reading only the target's
  // level left the difference unreported and the round trip unsatisfiable.
  if (baseBlock.listLevel !== targetBlock.listLevel) {
    properties.listLevel = targetBlock.listLevel ?? null;
  }
  return Object.keys(properties).length > 0 ? properties : null;
};

const locationOf = (story: FolioDocumentStoryHandle, block: FolioAIBlock): CompareChangeLocation =>
  block.table ? { story, cell: block.table } : { story };

export type CompareStoryPlan = {
  changes: CompareChange[];
  operations: FolioAIEditOperation[];
};

export type PlanStoryCompareOptions = {
  story: FolioDocumentStoryHandle;
  baseSnapshot: FolioAIEditSnapshot;
  targetBlocks: readonly FolioAIBlock[];
  /** Cap on generated operations; the caller turns `null` into its own error. */
  maxOperations: number;
};

/**
 * Plan one story's comparison, or `null` when it needs more operations than
 * `maxOperations`.
 */
export const planStoryCompare = ({
  story,
  baseSnapshot,
  targetBlocks,
  maxOperations,
}: PlanStoryCompareOptions): CompareStoryPlan | null => {
  const steps = buildSteps({ baseBlocks: baseSnapshot.blocks, targetBlocks });
  const paragraphMarkPlans = detectParagraphMarkEdits(steps);
  // The step after each paragraph-mark plan is part of it, so neither the move
  // pass nor the main loop may claim it again.
  const consumedSteps = new Set([...paragraphMarkPlans.keys()].map((index) => index + 1));
  const movesByBaseBlockId = detectMoves(steps, consumedSteps);
  const moveSourceByTargetBlockId = new Map<string, string>();
  for (const [baseBlockId, targetBlockId] of movesByBaseBlockId) {
    moveSourceByTargetBlockId.set(targetBlockId, baseBlockId);
  }

  const anchorIds = nextBaseBlockIdByStep(steps);
  const changes: CompareChange[] = [];
  const operations: FolioAIEditOperation[] = [];
  /**
   * The anchor everything past the base document's content hangs from: its
   * last BODY-LEVEL paragraph, which the format guarantees exists because a
   * table may not be the last child of a body. Anchoring to the last block
   * instead put the anchor inside a table whenever the story ended with one,
   * and an insertion anchored there escapes to the table's boundary, where no
   * paragraph mark can express the break it added.
   *
   * The applier orders insertions that resolve to one position by their order
   * in this array, so the tail is emitted where its step sits rather than
   * collected and appended — a table and a paragraph both added after the last
   * base block otherwise come out in operation order, which is not target
   * order.
   */
  const tailAnchorId = trailingBodyBlockId(baseSnapshot);

  let operationSequence = 0;

  const nextOperationId = (): string => `compare-${++operationSequence}`;

  /**
   * The relocation this block belongs to, named so the applier can link the
   * deletion at the source with the insertion at the destination as
   * `w:moveFrom` and `w:moveTo` instead of writing two unrelated revisions.
   */
  const moveIdOf = (baseBlockId: string): string => `move-${baseBlockId}`;

  const pushInsertOperation = (block: FolioAIBlock, anchorId: string | null): void => {
    const moveSourceId = moveSourceByTargetBlockId.get(block.id);
    // Both always explicit, `null` included: an inserted paragraph that says
    // nothing about its style or its list level takes the anchor's, and the
    // anchor is whichever block happened to follow it. A new paragraph beside
    // a list item is not a list item.
    const shared = {
      text: block.text,
      ...(moveSourceId !== undefined && { moveId: moveIdOf(moveSourceId) }),
      styleId: block.styleId ?? null,
      listLevel: block.listLevel ?? null,
    };
    if (anchorId !== null) {
      operations.push({
        id: nextOperationId(),
        type: "insertBeforeBlock",
        blockId: anchorId,
        ...shared,
      });
      return;
    }
    if (tailAnchorId === null) {
      // An empty base with no anchor paragraph cannot receive tracked
      // insertions at all.
      return;
    }
    operations.push({
      id: nextOperationId(),
      type: "insertAfterBlock",
      blockId: tailAnchorId,
      ...shared,
    });
  };

  for (const [stepIndex, step] of steps.entries()) {
    if (consumedSteps.has(stepIndex)) {
      continue;
    }
    const paragraphMarkPlan = paragraphMarkPlans.get(stepIndex);
    if (paragraphMarkPlan?.type === "split") {
      const { baseBlock, targetBlocks: splitInto, offset, separator } = paragraphMarkPlan;
      changes.push({
        kind: "split",
        location: locationOf(story, baseBlock),
        baseBlockId: baseBlock.id,
        targetBlockIds: splitInto.map(({ id }) => id),
        text: baseBlock.text,
      });
      operations.push({
        id: nextOperationId(),
        type: "splitBlock",
        blockId: baseBlock.id,
        offset,
        ...(separator.length > 0 && { separator }),
      });
      continue;
    }
    if (paragraphMarkPlan?.type === "merge") {
      const { baseBlocks, targetBlock, separator } = paragraphMarkPlan;
      changes.push({
        kind: "merge",
        location: locationOf(story, baseBlocks[0]),
        baseBlockIds: baseBlocks.map(({ id }) => id),
        targetBlockId: targetBlock.id,
        text: targetBlock.text,
      });
      operations.push({
        id: nextOperationId(),
        type: "mergeBlockWithNext",
        blockId: baseBlocks[0].id,
        ...(separator.length > 0 && { separator }),
      });
      continue;
    }
    switch (step.type) {
      case "pair": {
        const { baseBlock, targetBlock } = step;
        const properties = changedParagraphProperties(baseBlock, targetBlock);
        if (properties) {
          changes.push({
            kind: "paragraph-format",
            location: locationOf(story, baseBlock),
            baseBlockId: baseBlock.id,
            targetBlockId: targetBlock.id,
            properties,
          });
          operations.push({
            id: nextOperationId(),
            type: "setBlockParagraphProperties",
            blockId: baseBlock.id,
            properties,
          });
        }
        if (baseBlock.text !== targetBlock.text) {
          changes.push({
            kind: "replace",
            location: locationOf(story, baseBlock),
            baseBlockId: baseBlock.id,
            targetBlockId: targetBlock.id,
            before: baseBlock.text,
            after: targetBlock.text,
          });
          operations.push({
            id: nextOperationId(),
            type: "replaceBlock",
            blockId: baseBlock.id,
            text: targetBlock.text,
          });
          break;
        }
        const segments = inlineFormattingSegments({
          baseBlock,
          targetBlock,
          maxSegments: maxOperations,
        });
        if (segments === null) {
          return null;
        }
        if (segments.length === 0) {
          break;
        }
        changes.push({
          kind: "format",
          location: locationOf(story, baseBlock),
          baseBlockId: baseBlock.id,
          targetBlockId: targetBlock.id,
          text: baseBlock.text,
          ranges: segments,
        });
        for (const { startOffset, endOffset, formatting } of segments) {
          const range = createFolioAITextRangeHandle({
            blockId: baseBlock.id,
            text: baseBlock.text,
            startOffset,
            endOffset,
          });
          if (!range) {
            panic("An aligned formatting range could not be represented");
          }
          operations.push({ id: nextOperationId(), type: "formatRange", range, formatting });
        }
        break;
      }
      case "baseOnly": {
        const targetBlockId = movesByBaseBlockId.get(step.block.id);
        if (targetBlockId === undefined) {
          changes.push({
            kind: "delete",
            location: locationOf(story, step.block),
            baseBlockId: step.block.id,
            before: step.block.text,
          });
        }
        operations.push({
          id: nextOperationId(),
          type: "deleteBlock",
          blockId: step.block.id,
          ...(targetBlockId !== undefined && { moveId: moveIdOf(step.block.id) }),
        });
        break;
      }
      case "targetOnly": {
        const baseBlockId = moveSourceByTargetBlockId.get(step.block.id);
        changes.push(
          baseBlockId === undefined
            ? {
                kind: "insert",
                location: locationOf(story, step.block),
                targetBlockId: step.block.id,
                after: step.block.text,
              }
            : {
                kind: "move",
                location: locationOf(story, step.block),
                baseBlockId,
                targetBlockId: step.block.id,
                text: step.block.text,
              },
        );
        pushInsertOperation(step.block, anchorIds[stepIndex] ?? null);
        break;
      }
      case "baseRow": {
        const anchorBlockId = step.blocks[0]?.id;
        if (anchorBlockId === undefined) {
          panic("A collapsed table row carried no blocks");
        }
        changes.push({
          kind: "table-row-delete",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rowIndex: step.location.rowIndex,
          cells: rowCellTexts(step.blocks),
          baseBlockIds: step.blocks.map(({ id }) => id),
        });
        operations.push({ id: nextOperationId(), type: "deleteTableRow", blockId: anchorBlockId });
        break;
      }
      case "baseTable": {
        const anchorBlockId = step.blocks[0]?.id;
        if (anchorBlockId === undefined) {
          panic("An unpaired table segment carried no blocks");
        }
        changes.push({
          kind: "table-delete",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rows: tableCellTexts(step.blocks),
          baseBlockIds: step.blocks.map(({ id }) => id),
        });
        operations.push({ id: nextOperationId(), type: "deleteTable", blockId: anchorBlockId });
        break;
      }
      case "targetTable": {
        const rows = tableCellTexts(step.blocks);
        changes.push({
          kind: "table-insert",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rows,
          targetBlockIds: step.blocks.map(({ id }) => id),
        });
        const before = anchorIds[stepIndex] ?? null;
        if (before !== null) {
          operations.push({
            id: nextOperationId(),
            type: "insertTable",
            blockId: before,
            position: "before",
            rows,
          });
          break;
        }
        if (tailAnchorId !== null) {
          operations.push({
            id: nextOperationId(),
            type: "insertTable",
            blockId: tailAnchorId,
            position: "after",
            rows,
          });
        }
        break;
      }
      case "targetRow": {
        const anchor = findRowAnchor(steps, stepIndex);
        const cells = rowCellTexts(step.blocks);
        changes.push({
          kind: "table-row-insert",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rowIndex: step.location.rowIndex,
          cells,
          targetBlockIds: step.blocks.map(({ id }) => id),
        });
        if (anchor === null) {
          // The base story has no table to grow, so the row's text can only
          // land as ordinary paragraphs.
          for (const block of step.blocks) {
            pushInsertOperation(block, anchorIds[stepIndex] ?? null);
          }
          break;
        }
        operations.push({
          id: nextOperationId(),
          type: "insertTableRow",
          blockId: anchor.blockId,
          position: anchor.position,
          cellTexts: cells,
        });
        break;
      }
      default: {
        const unreachable: never = step;
        panic("Unhandled compare step", { step: unreachable });
      }
    }
    if (operations.length > maxOperations) {
      return null;
    }
  }

  return operations.length > maxOperations ? null : { changes, operations };
};
