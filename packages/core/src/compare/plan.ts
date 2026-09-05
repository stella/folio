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
import { createFolioAITextRangeHandle } from "../ai-edits/snapshot";
import type {
  FolioAIBlock,
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
  | { type: "targetRow"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation };

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
    const tableIndex = block.table?.tableIndex ?? null;
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
 * A stand-in block used only to run {@link alignFolioBlocks} over table rows,
 * which are not paragraphs. The id is always the `seq-NNNN` shape, which the
 * alignment treats as unstable, so it pairs rows on text and position alone
 * instead of mistaking a synthetic id for identity.
 */
const proxyBlock = (index: number, text: string): FolioAIBlock => ({
  id: `seq-${String(index + 1).padStart(4, "0")}`,
  kind: "paragraph",
  text,
});

/** Blocks of one table grouped into rows, in row order. */
const groupRows = (blocks: readonly FolioAIBlock[]): FolioAIBlock[][] => {
  const rows = new Map<number, FolioAIBlock[]>();
  for (const block of blocks) {
    if (!block.table) {
      continue;
    }
    const row = rows.get(block.table.rowIndex);
    if (row) {
      row.push(block);
    } else {
      rows.set(block.table.rowIndex, [block]);
    }
  }
  return [...rows.keys()].toSorted((left, right) => left - right).map((key) => rows.get(key) ?? []);
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
const buildTableSteps = (
  baseBlocks: readonly FolioAIBlock[],
  targetBlocks: readonly FolioAIBlock[],
): CompareStep[] => {
  const baseRows = groupRows(baseBlocks);
  const targetRows = groupRows(targetBlocks);
  const steps: CompareStep[] = [];

  let baseCursor = 0;
  let targetCursor = 0;
  for (const event of alignFolioBlocks(
    baseRows.map((row, index) => proxyBlock(index, rowText(row))),
    targetRows.map((row, index) => proxyBlock(index, rowText(row))),
  )) {
    switch (event.type) {
      case "pair": {
        const baseRow = baseRows[baseCursor++];
        const targetRow = targetRows[targetCursor++];
        if (baseRow && targetRow) {
          steps.push(...alignRowCells(baseRow, targetRow));
        }
        break;
      }
      case "baseOnly": {
        const row = baseRows[baseCursor++];
        const location = row ? rowLocation(row) : null;
        if (row && location) {
          steps.push({ type: "baseRow", blocks: row, location });
        }
        break;
      }
      case "revisedOnly": {
        const row = targetRows[targetCursor++];
        const location = row ? rowLocation(row) : null;
        if (row && location) {
          steps.push({ type: "targetRow", blocks: row, location });
        }
        break;
      }
      default: {
        const unreachable: never = event;
        panic("Unhandled block alignment event", { event: unreachable });
      }
    }
  }
  return steps;
};

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

/** Every block of an unpaired segment, as one-sided steps. */
const unpairedSegmentSteps = (segment: DocumentSegment, side: "base" | "target"): CompareStep[] => {
  if (segment.kind === "table") {
    return groupRows(segment.blocks).flatMap((row) => {
      const location = rowLocation(row);
      if (!location) {
        return [];
      }
      return [{ type: side === "base" ? "baseRow" : "targetRow", blocks: row, location }];
    });
  }
  return segment.blocks.map((block) =>
    side === "base" ? { type: "baseOnly", block } : { type: "targetOnly", block },
  );
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
 * and the surplus segments are reported one-sided. The operation vocabulary
 * cannot create or destroy a table, so `compareDocx`'s round-trip self-check
 * refuses those comparisons rather than returning a package that silently
 * drops one.
 */
const buildSteps = ({ baseBlocks, targetBlocks }: BuildStepsOptions): CompareStep[] => {
  const baseSegments = splitSegments(baseBlocks);
  const targetSegments = splitSegments(targetBlocks);
  const pairedCount = Math.min(baseSegments.length, targetSegments.length);
  const steps: CompareStep[] = [];

  for (let index = 0; index < pairedCount; index++) {
    const baseSegment = baseSegments[index];
    const targetSegment = targetSegments[index];
    if (!baseSegment || !targetSegment) {
      continue;
    }
    if (baseSegment.kind !== targetSegment.kind) {
      steps.push(
        ...unpairedSegmentSteps(baseSegment, "base"),
        ...unpairedSegmentSteps(targetSegment, "target"),
      );
      continue;
    }
    steps.push(
      ...(baseSegment.kind === "table"
        ? buildTableSteps(baseSegment.blocks, targetSegment.blocks)
        : buildBodySteps(baseSegment.blocks, targetSegment.blocks)),
    );
  }

  for (const segment of baseSegments.slice(pairedCount)) {
    steps.push(...unpairedSegmentSteps(segment, "base"));
  }
  for (const segment of targetSegments.slice(pairedCount)) {
    steps.push(...unpairedSegmentSteps(segment, "target"));
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
      return step.blocks[0] ?? null;
    case "targetOnly":
    case "targetRow":
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
const rowCellTexts = (blocks: readonly FolioAIBlock[]): string[] => {
  const byCell: (string | undefined)[] = [];
  for (const block of blocks) {
    const cellIndex = block.table?.cellIndex ?? 0;
    const existing = byCell[cellIndex];
    byCell[cellIndex] = existing === undefined ? block.text : `${existing}\n${block.text}`;
  }
  return Array.from(byCell, (text) => text ?? "");
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
  const trailingInserts: FolioAIBlock[] = [];
  const lastBaseBlockId = baseSnapshot.blocks.at(-1)?.id ?? null;
  let operationSequence = 0;

  const nextOperationId = (): string => `compare-${++operationSequence}`;

  /**
   * The relocation this block belongs to, named so the applier can link the
   * deletion at the source with the insertion at the destination as
   * `w:moveFrom` and `w:moveTo` instead of writing two unrelated revisions.
   */
  const moveIdOf = (baseBlockId: string): string => `move-${baseBlockId}`;

  const pushInsertOperation = (block: FolioAIBlock, anchorId: string | null): void => {
    if (anchorId === null) {
      trailingInserts.push(block);
      return;
    }
    const moveSourceId = moveSourceByTargetBlockId.get(block.id);
    operations.push({
      id: nextOperationId(),
      type: "insertBeforeBlock",
      blockId: anchorId,
      text: block.text,
      ...(moveSourceId !== undefined && { moveId: moveIdOf(moveSourceId) }),
      ...(block.styleId !== undefined && { styleId: block.styleId }),
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

  // An empty base document has no block to insert after, only the hidden
  // anchor paragraph. Its first addition therefore replaces that paragraph
  // instead of following it, so the result does not open with a stray blank.
  const emptyBaseAnchorId =
    lastBaseBlockId === null ? baseSnapshot.emptyDocumentAnchorId : undefined;
  const anchorId = lastBaseBlockId ?? emptyBaseAnchorId;
  for (const [insertIndex, block] of trailingInserts.entries()) {
    if (anchorId === undefined) {
      // Nothing to anchor to: an empty base with no anchor paragraph cannot
      // receive tracked insertions at all.
      break;
    }
    const styleId = block.styleId === undefined ? {} : { styleId: block.styleId };
    if (insertIndex === 0 && emptyBaseAnchorId !== undefined) {
      operations.push({
        id: nextOperationId(),
        type: "replaceBlock",
        blockId: emptyBaseAnchorId,
        text: block.text,
        ...styleId,
      });
      continue;
    }
    const moveSourceId = moveSourceByTargetBlockId.get(block.id);
    operations.push({
      id: nextOperationId(),
      type: "insertAfterBlock",
      blockId: anchorId,
      text: block.text,
      ...(moveSourceId !== undefined && { moveId: moveIdOf(moveSourceId) }),
      ...styleId,
    });
  }

  return operations.length > maxOperations ? null : { changes, operations };
};
