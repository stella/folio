/**
 * Turn two block snapshots of one story into the change list a caller reads
 * and the {@link FolioAIEditOperation}s that reproduce the target when every
 * generated tracked change is accepted.
 *
 * Pure: no parsing, no serialization, no clock. The alignment itself is
 * {@link alignFolioBlocks}, shared with the version-diff and redline paths, so
 * all three interpret one document walk rather than re-deriving it.
 *
 * ## Structural passes over the alignment
 *
 * 1. **Whole-row collapse.** A run of consecutive base-only (or target-only)
 *    blocks that covers every block of one table row is a row deletion (or
 *    insertion), and is emitted as `deleteTableRow` / `insertTableRow`. Left as
 *    paragraph operations it would strand the target row's text outside the
 *    table, because a block insertion anchored inside a table lands as a
 *    sibling of the table, not as a new row.
 * 2. **Move detection.** A base-only and a target-only block with identical
 *    text and at least {@link MOVE_MINIMUM_WORD_COUNT} words are one relocation.
 *    The operations stay a delete plus an insert (the tracked-change grammar
 *    Word round-trips has no durable "moved from here" mark for this path), but
 *    the change list reports a single `move` so the relocation is not read as
 *    unrelated churn. The word floor keeps boilerplate one-liners from pairing
 *    as spurious moves.
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

const rowKey = ({ tableIndex, rowIndex }: FolioAIBlockTableLocation): string =>
  `${tableIndex}:${rowIndex}`;

/** How many blocks each table row holds, so a partial run is not read as a whole-row change. */
const countBlocksPerRow = (blocks: readonly FolioAIBlock[]): ReadonlyMap<string, number> => {
  const counts = new Map<string, number>();
  for (const { table } of blocks) {
    if (table) {
      const key = rowKey(table);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
  }
  return counts;
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
 * One step of the interpreted alignment. `baseRow` / `targetRow` are the
 * whole-row collapses; the rest mirror {@link alignFolioBlocks}'s events with
 * the revised side renamed to the target side.
 */
type CompareStep =
  | { type: "pair"; baseBlock: FolioAIBlock; targetBlock: FolioAIBlock }
  | { type: "baseOnly"; block: FolioAIBlock }
  | { type: "targetOnly"; block: FolioAIBlock }
  | { type: "baseRow"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | { type: "targetRow"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation };

/**
 * The blocks a same-side run covers, grouped into maximal same-row prefixes.
 * A group whose size matches the row's block count is a whole-row change.
 */
const collapseRun = (
  run: readonly FolioAIBlock[],
  blocksPerRow: ReadonlyMap<string, number>,
  side: "base" | "target",
): CompareStep[] => {
  const steps: CompareStep[] = [];
  let index = 0;
  while (index < run.length) {
    const first = run[index];
    if (!first) {
      break;
    }
    const { table } = first;
    if (!table) {
      steps.push(
        side === "base"
          ? { type: "baseOnly", block: first }
          : { type: "targetOnly", block: first },
      );
      index++;
      continue;
    }
    const key = rowKey(table);
    let end = index;
    while (end < run.length) {
      const candidate = run[end]?.table;
      if (!candidate || rowKey(candidate) !== key) {
        break;
      }
      end++;
    }
    const group = run.slice(index, end);
    if (group.length === blocksPerRow.get(key)) {
      steps.push(
        side === "base"
          ? { type: "baseRow", blocks: group, location: table }
          : { type: "targetRow", blocks: group, location: table },
      );
    } else {
      for (const block of group) {
        steps.push(
          side === "base" ? { type: "baseOnly", block } : { type: "targetOnly", block },
        );
      }
    }
    index = end;
  }
  return steps;
};

type BuildStepsOptions = {
  baseBlocks: readonly FolioAIBlock[];
  targetBlocks: readonly FolioAIBlock[];
};

const buildSteps = ({ baseBlocks, targetBlocks }: BuildStepsOptions): CompareStep[] => {
  const baseBlocksPerRow = countBlocksPerRow(baseBlocks);
  const targetBlocksPerRow = countBlocksPerRow(targetBlocks);
  const steps: CompareStep[] = [];
  let baseRun: FolioAIBlock[] = [];
  let targetRun: FolioAIBlock[] = [];

  const flush = () => {
    if (baseRun.length > 0) {
      steps.push(...collapseRun(baseRun, baseBlocksPerRow, "base"));
      baseRun = [];
    }
    if (targetRun.length > 0) {
      steps.push(...collapseRun(targetRun, targetBlocksPerRow, "target"));
      targetRun = [];
    }
  };

  for (const event of alignFolioBlocks(baseBlocks, targetBlocks)) {
    switch (event.type) {
      case "pair":
        flush();
        steps.push({ type: "pair", baseBlock: event.baseBlock, targetBlock: event.revisedBlock });
        break;
      case "baseOnly":
        baseRun.push(event.block);
        break;
      case "revisedOnly":
        targetRun.push(event.block);
        break;
      default: {
        const unreachable: never = event;
        panic("Unhandled block alignment event", { event: unreachable });
      }
    }
  }
  flush();
  return steps;
};

/** Base block id -> target block id for every relocation the move pass found. */
const detectMoves = (steps: readonly CompareStep[]): ReadonlyMap<string, string> => {
  const candidatesByText = new Map<string, string[]>();
  for (const step of steps) {
    if (step.type !== "baseOnly" || !wordCountReaches(step.block.text, MOVE_MINIMUM_WORD_COUNT)) {
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

  const movesByBaseBlockId = new Map<string, string>();
  for (const step of steps) {
    if (step.type !== "targetOnly") {
      continue;
    }
    const queue = candidatesByText.get(step.block.text);
    const baseBlockId = queue?.shift();
    if (baseBlockId !== undefined) {
      movesByBaseBlockId.set(baseBlockId, step.block.id);
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

/** A row's blocks joined per physical cell, in cell order. */
const rowCellTexts = (blocks: readonly FolioAIBlock[]): string[] => {
  const byCell = new Map<number, string[]>();
  for (const block of blocks) {
    if (!block.table) {
      continue;
    }
    const texts = byCell.get(block.table.cellIndex);
    if (texts) {
      texts.push(block.text);
    } else {
      byCell.set(block.table.cellIndex, [block.text]);
    }
  }
  return [...byCell.keys()]
    .toSorted((left, right) => left - right)
    .map((cellIndex) => (byCell.get(cellIndex) ?? []).join("\n"));
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
  const movesByBaseBlockId = detectMoves(steps);
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

  const pushInsertOperation = (block: FolioAIBlock, anchorId: string | null): void => {
    if (anchorId === null) {
      trailingInserts.push(block);
      return;
    }
    operations.push({
      id: nextOperationId(),
      type: "insertBeforeBlock",
      blockId: anchorId,
      text: block.text,
      ...(block.styleId !== undefined && { styleId: block.styleId }),
    });
  };

  for (const [stepIndex, step] of steps.entries()) {
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
        operations.push({ id: nextOperationId(), type: "deleteBlock", blockId: step.block.id });
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
    operations.push({
      id: nextOperationId(),
      type: "insertAfterBlock",
      blockId: anchorId,
      text: block.text,
      ...styleId,
    });
  }

  return operations.length > maxOperations ? null : { changes, operations };
};
