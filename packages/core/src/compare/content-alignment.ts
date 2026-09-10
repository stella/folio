/**
 * Pure alignment for representation-neutral content snapshots.
 *
 * The block walk is split by structural container before any textual
 * alignment runs. Body blocks are compared with body blocks; tables are
 * aligned table by table, then row by row and cell by cell. All quadratic
 * text alignment in one caller-owned work session shares one hard budget.
 */

import { panic } from "better-result";

import type {
  FolioContentBlock,
  FolioContentIdStability,
  FolioContentTableLocation,
} from "./content-types";
import { alignTableColumns, type TableColumnAlignmentStep } from "./column-alignment";

const MAX_CONTENT_ALIGNMENT_LCS_CELLS = 4_000_000;
const MAX_CONTENT_STRUCTURE_TOKEN_LOOKUPS = 4_000_000;

export type FolioContentAlignmentWorkSession = {
  remainingLcsCells: number;
  remainingStructuralTokenLookups: number;
};

export type CreateFolioContentAlignmentWorkSessionOptions = {
  /** @internal A lower allowance used by focused boundary tests. */
  lcsCells?: number;
  /** @internal A lower allowance used by focused boundary tests. */
  structuralTokenLookups?: number;
};

export const createFolioContentAlignmentWorkSession = (
  options: CreateFolioContentAlignmentWorkSessionOptions = {},
): FolioContentAlignmentWorkSession => {
  const lcsCells = options.lcsCells ?? MAX_CONTENT_ALIGNMENT_LCS_CELLS;
  const structuralTokenLookups =
    options.structuralTokenLookups ?? MAX_CONTENT_STRUCTURE_TOKEN_LOOKUPS;
  if (
    !Number.isSafeInteger(lcsCells) ||
    lcsCells < 0 ||
    lcsCells > MAX_CONTENT_ALIGNMENT_LCS_CELLS
  ) {
    return panic("A content-alignment LCS allowance is outside the supported range", {
      lcsCells,
    });
  }
  if (
    !Number.isSafeInteger(structuralTokenLookups) ||
    structuralTokenLookups < 0 ||
    structuralTokenLookups > MAX_CONTENT_STRUCTURE_TOKEN_LOOKUPS
  ) {
    return panic("A content-alignment structural token allowance is outside the supported range", {
      structuralTokenLookups,
    });
  }
  return {
    remainingLcsCells: lcsCells,
    remainingStructuralTokenLookups: structuralTokenLookups,
  };
};

export const exceedsFolioContentLcsBudget = (baseCount: number, revisedCount: number): boolean =>
  baseCount > 0 &&
  revisedCount > 0 &&
  baseCount > Math.floor(MAX_CONTENT_ALIGNMENT_LCS_CELLS / revisedCount);

const claimFolioContentAlignmentCells = (
  baseCount: number,
  revisedCount: number,
  workSession: FolioContentAlignmentWorkSession,
): boolean => {
  if (baseCount === 0 || revisedCount === 0) {
    return true;
  }
  if (
    exceedsFolioContentLcsBudget(baseCount, revisedCount) ||
    baseCount > Math.floor(workSession.remainingLcsCells / revisedCount)
  ) {
    return false;
  }
  workSession.remainingLcsCells -= baseCount * revisedCount;
  return true;
};

export type FolioContentBlockPair = { baseIndex: number; revisedIndex: number };
type IndexedBlock<Block extends FolioContentBlock> = { block: Block; index: number };

const folioContentIdStability = <Block extends FolioContentBlock>(
  block: Block,
): FolioContentIdStability => block.idStability ?? "stable";

/**
 * Longest strictly increasing subsequence by revised index.
 *
 * The Fenwick tree stores the earliest end index for equal-length candidates,
 * preserving the historical deterministic tie break while reducing the
 * stable-anchor pass from quadratic to O(n log n).
 */
export const longestIncreasingFolioContentPairs = (
  pairs: readonly FolioContentBlockPair[],
): FolioContentBlockPair[] => {
  if (pairs.length === 0) {
    return [];
  }
  const coordinates = [...new Set(pairs.map(({ revisedIndex }) => revisedIndex))].toSorted(
    (left, right) => left - right,
  );
  const rankByCoordinate = new Map(
    coordinates.map((coordinate, index) => [coordinate, index + 1] as const),
  );
  const treeLengths = new Int32Array(coordinates.length + 1);
  const treeIndexes = new Int32Array(coordinates.length + 1).fill(-1);
  const lengths = new Int32Array(pairs.length).fill(1);
  const predecessors = new Int32Array(pairs.length).fill(-1);

  const earlier = (
    leftLength: number,
    leftIndex: number,
    rightLength: number,
    rightIndex: number,
  ): boolean =>
    leftLength > rightLength ||
    (leftLength === rightLength && leftLength > 0 && leftIndex < rightIndex);

  let bestEnd = 0;
  let groupStart = 0;
  while (groupStart < pairs.length) {
    const baseIndex = pairs[groupStart]?.baseIndex;
    let groupEnd = groupStart + 1;
    while (groupEnd < pairs.length && pairs[groupEnd]?.baseIndex === baseIndex) {
      groupEnd += 1;
    }

    // Delay tree updates until the whole base-index group is scored. Otherwise two
    // candidates for one base container can become predecessor and successor.
    for (let pairIndex = groupStart; pairIndex < groupEnd; pairIndex++) {
      const pair = pairs[pairIndex];
      const rank = pair === undefined ? undefined : rankByCoordinate.get(pair.revisedIndex);
      if (rank === undefined) {
        continue;
      }

      let predecessorLength = 0;
      let predecessorIndex = -1;
      for (let cursor = rank - 1; cursor > 0; cursor -= cursor & -cursor) {
        const candidateLength = treeLengths[cursor] ?? 0;
        const candidateIndex = treeIndexes[cursor] ?? -1;
        if (earlier(candidateLength, candidateIndex, predecessorLength, predecessorIndex)) {
          predecessorLength = candidateLength;
          predecessorIndex = candidateIndex;
        }
      }
      lengths[pairIndex] = predecessorLength + 1;
      predecessors[pairIndex] = predecessorIndex;
      if ((lengths[pairIndex] ?? 0) > (lengths[bestEnd] ?? 0)) {
        bestEnd = pairIndex;
      }
    }

    for (let pairIndex = groupStart; pairIndex < groupEnd; pairIndex++) {
      const pair = pairs[pairIndex];
      const rank = pair === undefined ? undefined : rankByCoordinate.get(pair.revisedIndex);
      if (rank === undefined) {
        continue;
      }
      for (let cursor = rank; cursor < treeLengths.length; cursor += cursor & -cursor) {
        const currentLength = treeLengths[cursor] ?? 0;
        const currentIndex = treeIndexes[cursor] ?? -1;
        if (earlier(lengths[pairIndex] ?? 0, pairIndex, currentLength, currentIndex)) {
          treeLengths[cursor] = lengths[pairIndex] ?? 0;
          treeIndexes[cursor] = pairIndex;
        }
      }
    }
    groupStart = groupEnd;
  }

  const ordered: FolioContentBlockPair[] = [];
  for (let cursor = bestEnd; cursor !== -1; cursor = predecessors[cursor] ?? -1) {
    const pair = pairs[cursor];
    if (pair) {
      ordered.push(pair);
    }
  }
  return ordered.toReversed();
};

const pairByStableId = <Block extends FolioContentBlock>(
  base: readonly Block[],
  revised: readonly Block[],
  idStability: (block: Block) => FolioContentIdStability,
): FolioContentBlockPair[] => {
  const revisedIndexById = new Map<string, number>();
  revised.forEach((block, revisedIndex) => {
    if (idStability(block) === "stable") {
      revisedIndexById.set(block.id, revisedIndex);
    }
  });

  const candidates: FolioContentBlockPair[] = [];
  base.forEach((block, baseIndex) => {
    if (idStability(block) !== "stable") {
      return;
    }
    const revisedIndex = revisedIndexById.get(block.id);
    if (revisedIndex !== undefined) {
      candidates.push({ baseIndex, revisedIndex });
    }
  });
  return longestIncreasingFolioContentPairs(candidates);
};

type PairByResidualIdContinuityOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  baseFrom: number;
  baseTo: number;
  revisedFrom: number;
  revisedTo: number;
  idStability: (block: Block) => FolioContentIdStability;
};

/**
 * Pair ids that became stable when one side was serialized and reopened.
 *
 * An adapter can mark a synthesized id positional in the live snapshot that
 * minted it, then stable after that same id is persisted. This evidence is
 * deliberately weaker than a stable id: it is considered only inside a gap
 * already bounded by stable/exact anchors, so it cannot turn a positional id
 * into a move or pull a block across an established correspondence.
 */
const pairByResidualIdContinuity = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  baseFrom,
  baseTo,
  revisedFrom,
  revisedTo,
  idStability,
}: PairByResidualIdContinuityOptions<Block>): FolioContentBlockPair[] => {
  const uniqueIndexesById = ({
    blocks,
    from,
    to,
  }: {
    blocks: readonly Block[];
    from: number;
    to: number;
  }): ReadonlyMap<string, number | null> => {
    const indexes = new Map<string, number | null>();
    for (let index = from; index < to; index++) {
      const id = blocks[index]?.id;
      if (id === undefined) {
        continue;
      }
      indexes.set(id, indexes.has(id) ? null : index);
    }
    return indexes;
  };

  const baseIndexes = uniqueIndexesById({ blocks: baseBlocks, from: baseFrom, to: baseTo });
  const revisedIndexes = uniqueIndexesById({
    blocks: revisedBlocks,
    from: revisedFrom,
    to: revisedTo,
  });
  const candidates: FolioContentBlockPair[] = [];
  for (const [id, baseIndex] of baseIndexes) {
    const revisedIndex = revisedIndexes.get(id);
    if (baseIndex === null || revisedIndex === undefined || revisedIndex === null) {
      continue;
    }
    const baseBlock = baseBlocks[baseIndex];
    const revisedBlock = revisedBlocks[revisedIndex];
    if (!baseBlock || !revisedBlock) {
      continue;
    }
    const baseStability = idStability(baseBlock);
    const revisedStability = idStability(revisedBlock);
    if (baseStability === revisedStability) {
      continue;
    }
    candidates.push({ baseIndex, revisedIndex });
  }
  return longestIncreasingFolioContentPairs(candidates);
};

const pairByExactText = <Block extends FolioContentBlock>(
  base: readonly IndexedBlock<Block>[],
  revised: readonly IndexedBlock<Block>[],
  workSession: FolioContentAlignmentWorkSession,
  idStability: (block: Block) => FolioContentIdStability,
  stableIdMismatch: "pair" | "separate",
): FolioContentBlockPair[] => {
  const baseCount = base.length;
  const revisedCount = revised.length;
  if (baseCount === 0 || revisedCount === 0) {
    return [];
  }
  if (!claimFolioContentAlignmentCells(baseCount, revisedCount, workSession)) {
    return [];
  }

  const textKeys = new Map<string, number>();
  let nextTextKey = 0;
  const internText = (text: string): number => {
    const existing = textKeys.get(text);
    if (existing !== undefined) {
      return existing;
    }
    const key = nextTextKey++;
    textKeys.set(text, key);
    return key;
  };
  const baseTextKeys = base.map(({ block }) => internText(block.text));
  const revisedTextKeys = revised.map(({ block }) => internText(block.text));
  const entriesCanPair = (baseIndex: number, revisedIndex: number): boolean => {
    const baseBlock = base[baseIndex]?.block;
    const revisedBlock = revised[revisedIndex]?.block;
    if (!baseBlock || !revisedBlock || baseTextKeys[baseIndex] !== revisedTextKeys[revisedIndex]) {
      return false;
    }
    return (
      stableIdMismatch === "pair" ||
      !(
        idStability(baseBlock) === "stable" &&
        idStability(revisedBlock) === "stable" &&
        baseBlock.id !== revisedBlock.id
      )
    );
  };
  const stride = revisedCount + 1;
  const lengths = new Int32Array((baseCount + 1) * stride);
  for (let baseIndex = baseCount - 1; baseIndex >= 0; baseIndex--) {
    const rowOffset = baseIndex * stride;
    const nextRowOffset = (baseIndex + 1) * stride;
    for (let revisedIndex = revisedCount - 1; revisedIndex >= 0; revisedIndex--) {
      lengths[rowOffset + revisedIndex] = entriesCanPair(baseIndex, revisedIndex)
        ? (lengths[nextRowOffset + revisedIndex + 1] ?? 0) + 1
        : Math.max(
            lengths[nextRowOffset + revisedIndex] ?? 0,
            lengths[rowOffset + revisedIndex + 1] ?? 0,
          );
    }
  }

  const pairs: FolioContentBlockPair[] = [];
  let baseIndex = 0;
  let revisedIndex = 0;
  while (baseIndex < baseCount && revisedIndex < revisedCount) {
    const baseEntry = base[baseIndex];
    const revisedEntry = revised[revisedIndex];
    if (!baseEntry || !revisedEntry) {
      break;
    }
    if (entriesCanPair(baseIndex, revisedIndex)) {
      pairs.push({ baseIndex: baseEntry.index, revisedIndex: revisedEntry.index });
      baseIndex += 1;
      revisedIndex += 1;
      continue;
    }
    const down = lengths[(baseIndex + 1) * stride + revisedIndex] ?? 0;
    const right = lengths[baseIndex * stride + revisedIndex + 1] ?? 0;
    if (down >= right) {
      baseIndex += 1;
    } else {
      revisedIndex += 1;
    }
  }
  return pairs;
};

export type FolioContentAlignedBlockEvent<Block extends FolioContentBlock = FolioContentBlock> =
  | { type: "pair"; baseBlock: Block; revisedBlock: Block }
  | { type: "baseOnly"; block: Block }
  | { type: "revisedOnly"; block: Block };

export type AlignFolioContentBlocksOptions<Block extends FolioContentBlock> = {
  workSession?: FolioContentAlignmentWorkSession;
  /** Let an already-aligned structural slot outrank differing authored IDs. */
  stableIdMismatch?: "pair" | "separate";
  /** @internal Compatibility hook for adapters whose source model predates `idStability`. */
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
};

export const alignFolioContentBlocks = <Block extends FolioContentBlock>(
  baseBlocks: readonly Block[],
  revisedBlocks: readonly Block[],
  options: AlignFolioContentBlocksOptions<Block> = {},
): FolioContentAlignedBlockEvent<Block>[] => {
  const workSession = options.workSession ?? createFolioContentAlignmentWorkSession();
  const stableIdMismatch = options.stableIdMismatch ?? "separate";
  const idStability = options.idStability ?? folioContentIdStability;
  const stableIdAnchors = pairByStableId(baseBlocks, revisedBlocks, idStability);
  const usedBaseIndexes = new Set(stableIdAnchors.map(({ baseIndex }) => baseIndex));
  const usedRevisedIndexes = new Set(stableIdAnchors.map(({ revisedIndex }) => revisedIndex));
  const baseRemaining = baseBlocks.flatMap((block, index) =>
    usedBaseIndexes.has(index) ? [] : [{ block, index }],
  );
  const revisedRemaining = revisedBlocks.flatMap((block, index) =>
    usedRevisedIndexes.has(index) ? [] : [{ block, index }],
  );
  const exactTextAnchors = pairByExactText(
    baseRemaining,
    revisedRemaining,
    workSession,
    idStability,
    stableIdMismatch,
  );
  const anchors = longestIncreasingFolioContentPairs(
    [...stableIdAnchors, ...exactTextAnchors].toSorted(
      (left, right) => left.baseIndex - right.baseIndex,
    ),
  );
  const events: FolioContentAlignedBlockEvent<Block>[] = [];

  const emitPositionalGap = (
    baseFrom: number,
    baseTo: number,
    revisedFrom: number,
    revisedTo: number,
  ): void => {
    const pairedCount = Math.min(baseTo - baseFrom, revisedTo - revisedFrom);
    for (let offset = 0; offset < pairedCount; offset++) {
      const baseBlock = baseBlocks[baseFrom + offset];
      const revisedBlock = revisedBlocks[revisedFrom + offset];
      if (baseBlock && revisedBlock) {
        const differentStableIdentities =
          idStability(baseBlock) === "stable" &&
          idStability(revisedBlock) === "stable" &&
          baseBlock.id !== revisedBlock.id;
        if (differentStableIdentities && stableIdMismatch === "separate") {
          events.push({ type: "baseOnly", block: baseBlock });
          events.push({ type: "revisedOnly", block: revisedBlock });
        } else {
          events.push({ type: "pair", baseBlock, revisedBlock });
        }
      }
    }
    for (let index = baseFrom + pairedCount; index < baseTo; index++) {
      const block = baseBlocks[index];
      if (block) {
        events.push({ type: "baseOnly", block });
      }
    }
    for (let index = revisedFrom + pairedCount; index < revisedTo; index++) {
      const block = revisedBlocks[index];
      if (block) {
        events.push({ type: "revisedOnly", block });
      }
    }
  };

  const emitGap = (
    baseFrom: number,
    baseTo: number,
    revisedFrom: number,
    revisedTo: number,
  ): void => {
    if (baseFrom === baseTo || revisedFrom === revisedTo) {
      emitPositionalGap(baseFrom, baseTo, revisedFrom, revisedTo);
      return;
    }
    const continuityPairs = pairByResidualIdContinuity({
      baseBlocks,
      revisedBlocks,
      baseFrom,
      baseTo,
      revisedFrom,
      revisedTo,
      idStability,
    });
    let baseCursor = baseFrom;
    let revisedCursor = revisedFrom;
    for (const pair of continuityPairs) {
      emitPositionalGap(baseCursor, pair.baseIndex, revisedCursor, pair.revisedIndex);
      const baseBlock = baseBlocks[pair.baseIndex];
      const revisedBlock = revisedBlocks[pair.revisedIndex];
      if (baseBlock && revisedBlock) {
        events.push({ type: "pair", baseBlock, revisedBlock });
      }
      baseCursor = pair.baseIndex + 1;
      revisedCursor = pair.revisedIndex + 1;
    }
    emitPositionalGap(baseCursor, baseTo, revisedCursor, revisedTo);
  };

  let baseCursor = 0;
  let revisedCursor = 0;
  for (const anchor of anchors) {
    emitGap(baseCursor, anchor.baseIndex, revisedCursor, anchor.revisedIndex);
    const baseBlock = baseBlocks[anchor.baseIndex];
    const revisedBlock = revisedBlocks[anchor.revisedIndex];
    if (baseBlock && revisedBlock) {
      events.push({ type: "pair", baseBlock, revisedBlock });
    }
    baseCursor = anchor.baseIndex + 1;
    revisedCursor = anchor.revisedIndex + 1;
  }
  emitGap(baseCursor, baseBlocks.length, revisedCursor, revisedBlocks.length);
  return events;
};

export type FolioContentAlignmentStep<Block extends FolioContentBlock = FolioContentBlock> =
  | { type: "pair"; baseBlock: Block; revisedBlock: Block }
  | { type: "baseOnly"; block: Block; moveScope: FolioContentMoveScope }
  | { type: "revisedOnly"; block: Block; moveScope: FolioContentMoveScope }
  | { type: "baseRow"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | { type: "revisedRow"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | { type: "baseTable"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | { type: "revisedTable"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | TableColumnAlignmentStep<Block>;

/** @internal Legal move bucket and alignment gap for one unpaired block. */
export type FolioContentMoveScope = {
  readonly bucket: number;
  readonly gap: number;
};

type MoveScopeContext = {
  nextTableCellBucket: number;
  nextGap: number;
};

const BODY_MOVE_BUCKET = 0;

const scopedAlignmentSteps = <Block extends FolioContentBlock>(
  events: readonly FolioContentAlignedBlockEvent<Block>[],
  context: MoveScopeContext,
  bucketForBlock: (block: Block) => number,
): FolioContentAlignmentStep<Block>[] => {
  const steps: FolioContentAlignmentStep<Block>[] = [];
  let gap = context.nextGap++;
  for (const event of events) {
    if (event.type === "pair") {
      steps.push(event);
      gap = context.nextGap++;
      continue;
    }
    steps.push({
      ...event,
      moveScope: { bucket: bucketForBlock(event.block), gap },
    });
  }
  return steps;
};

type DocumentSegment<Block extends FolioContentBlock> =
  | {
      kind: "body";
      blocks: Block[];
      containerPathKey: string | null;
      structuralKey: string;
    }
  | { kind: "table"; blocks: Block[]; containerPathKey: null; structuralKey: string };

const containerPathKeyOf = (block: FolioContentBlock): string | null =>
  block.containerPath === undefined || block.containerPath.length === 0
    ? null
    : JSON.stringify(block.containerPath.map(({ kind, id }) => [kind, id]));

const contentBlocksShareContainerPath = (
  left: FolioContentBlock,
  right: FolioContentBlock,
): boolean => containerPathKeyOf(left) === containerPathKeyOf(right);

const splitSegments = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
): DocumentSegment<Block>[] => {
  const segments: DocumentSegment<Block>[] = [];
  let currentTableIndex: number | null = null;
  for (const block of blocks) {
    const tableIndex = block.table?.outerTableIndex ?? null;
    const containerPathKey = block.table ? null : containerPathKeyOf(block);
    const current = segments.at(-1);
    if (
      current !== undefined &&
      currentTableIndex === tableIndex &&
      current.containerPathKey === containerPathKey
    ) {
      current.blocks.push(block);
      continue;
    }
    segments.push(
      block.table
        ? {
            kind: "table",
            blocks: [block],
            containerPathKey: null,
            structuralKey: JSON.stringify(["table", null]),
          }
        : {
            kind: "body",
            blocks: [block],
            containerPathKey,
            structuralKey: JSON.stringify(["body", containerPathKey]),
          },
    );
    currentTableIndex = tableIndex;
  }
  return segments;
};

export const groupFolioContentTableRows = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
): Block[][] => {
  const rows = new Map<string, Block[]>();
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

const groupTables = <Block extends FolioContentBlock>(blocks: readonly Block[]): Block[][] => {
  const tables = new Map<number, Block[]>();
  for (const block of blocks) {
    const tableIndex = block.table?.tableIndex;
    if (tableIndex === undefined) {
      continue;
    }
    const table = tables.get(tableIndex);
    if (table) {
      table.push(block);
    } else {
      tables.set(tableIndex, [block]);
    }
  }
  return [...tables.values()];
};

const CONTENT_STRUCTURE_PAIR_SIMILARITY = 0.5;
const CONTENT_STRUCTURE_SIMILARITY_SCALE = 1_000;
const MAX_CONTENT_STRUCTURE_PROFILE_BLOCKS = 128;
const MAX_CONTENT_STRUCTURE_PROFILE_TEXT_CODE_UNITS = 16_384;
const MAX_CONTENT_STRUCTURE_PROFILE_TOKENS = 64;
const MAX_CONTENT_STRUCTURE_PROFILE_KIND_CODE_UNITS = 256;
const CONTENT_STRUCTURE_ANCHOR_MINIMUM_WORD_COUNT = 3;

type ContentStructureAtom = string | number | null;

/** Complete row identity, or an explicit refusal when retaining it would exceed the profile cap. */
type ContentContainerIdentityProfile =
  | {
      status: "available";
      blockIds: readonly string[];
      idStabilities: readonly FolioContentIdStability[];
    }
  | { status: "unavailable" };

type ContentStructureProfile = {
  anchorTexts: readonly string[];
  blockIds: readonly string[];
  containerIdentity: ContentContainerIdentityProfile | null;
  exactSignature: string | null;
  physicalCellCount: number | null;
  stableIds: readonly string[];
  tokenCounts: ReadonlyMap<string, number> | null;
  tokenCount: number;
};

type CreateContentStructureProfileOptions<Block extends FolioContentBlock> = {
  blocks: readonly Block[];
  blockStructure: (block: Block) => readonly ContentStructureAtom[];
  idStability: (block: Block) => FolioContentIdStability;
  physicalCellCount?: number;
  retainContainerIdentity?: boolean;
};

const createContentStructureProfile = <Block extends FolioContentBlock>({
  blocks,
  blockStructure,
  idStability,
  physicalCellCount,
  retainContainerIdentity = false,
}: CreateContentStructureProfileOptions<Block>): ContentStructureProfile => {
  let containerIdentity: ContentContainerIdentityProfile | null = null;
  if (retainContainerIdentity) {
    if (blocks.length > MAX_CONTENT_STRUCTURE_PROFILE_BLOCKS) {
      containerIdentity = { status: "unavailable" };
    } else {
      const blockIds: string[] = [];
      const idStabilities: FolioContentIdStability[] = [];
      for (const block of blocks) {
        blockIds.push(block.id);
        idStabilities.push(idStability(block));
      }
      containerIdentity = { status: "available", blockIds, idStabilities };
    }
  }
  const profiledBlockCount = Math.min(blocks.length, MAX_CONTENT_STRUCTURE_PROFILE_BLOCKS);
  const blockIds: string[] = [];
  const stableIds: string[] = [];
  for (let index = 0; index < profiledBlockCount; index++) {
    const block = blocks[index];
    if (block) {
      blockIds.push(block.id);
      if (idStability(block) === "stable") {
        stableIds.push(block.id);
      }
    }
  }
  if (blocks.length > MAX_CONTENT_STRUCTURE_PROFILE_BLOCKS) {
    return {
      anchorTexts: [],
      blockIds,
      containerIdentity,
      exactSignature: null,
      physicalCellCount: physicalCellCount ?? null,
      stableIds,
      tokenCounts: null,
      tokenCount: 0,
    };
  }

  let textCodeUnits = 0;
  for (const block of blocks) {
    if (
      block.kind.length > MAX_CONTENT_STRUCTURE_PROFILE_KIND_CODE_UNITS ||
      block.text.length > MAX_CONTENT_STRUCTURE_PROFILE_TEXT_CODE_UNITS - textCodeUnits
    ) {
      return {
        anchorTexts: [],
        blockIds,
        containerIdentity,
        exactSignature: null,
        physicalCellCount: physicalCellCount ?? null,
        stableIds,
        tokenCounts: null,
        tokenCount: 0,
      };
    }
    textCodeUnits += block.text.length;
  }

  const tokenCounts = new Map<string, number>();
  const anchorTexts = new Set<string>();
  let tokenCount = 0;
  let tokensComplete = true;
  const signature = blocks.map((block) => [...blockStructure(block), block.kind, block.text]);
  for (const block of blocks) {
    let blockWordCount = 0;
    for (const match of block.text.matchAll(/\S+/gu)) {
      const token = match[0];
      blockWordCount += 1;
      tokenCount += 1;
      if (tokenCount > MAX_CONTENT_STRUCTURE_PROFILE_TOKENS) {
        tokensComplete = false;
        break;
      }
      tokenCounts.set(token, (tokenCounts.get(token) ?? 0) + 1);
    }
    if (blockWordCount >= CONTENT_STRUCTURE_ANCHOR_MINIMUM_WORD_COUNT) {
      anchorTexts.add(block.text);
    }
    if (!tokensComplete) {
      break;
    }
  }
  return {
    anchorTexts: [...anchorTexts],
    blockIds,
    containerIdentity,
    exactSignature: JSON.stringify(signature),
    physicalCellCount: physicalCellCount ?? null,
    stableIds,
    tokenCounts: tokensComplete ? tokenCounts : null,
    tokenCount: tokensComplete ? tokenCount : 0,
  };
};

const contentStructureProfileSimilarity = (
  base: ContentStructureProfile,
  revised: ContentStructureProfile,
  workSession: FolioContentAlignmentWorkSession,
): number => {
  if (
    !base.tokenCounts ||
    !revised.tokenCounts ||
    base.tokenCount === 0 ||
    revised.tokenCount === 0
  ) {
    return 0;
  }
  const [tokens, counterparts] =
    base.tokenCounts.size <= revised.tokenCounts.size
      ? [base.tokenCounts, revised.tokenCounts]
      : [revised.tokenCounts, base.tokenCounts];
  if (tokens.size > workSession.remainingStructuralTokenLookups) {
    return 0;
  }
  workSession.remainingStructuralTokenLookups -= tokens.size;
  let shared = 0;
  for (const [token, count] of tokens) {
    shared += Math.min(count, counterparts.get(token) ?? 0);
  }
  return (2 * shared) / (base.tokenCount + revised.tokenCount);
};

const tableStructureProfile = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
  idStability: (block: Block) => FolioContentIdStability,
): ContentStructureProfile => {
  const tableOrdinalByIndex = new Map<number, number>();
  return createContentStructureProfile({
    blocks,
    idStability,
    blockStructure: (block) => {
      const table = block.table ?? panic("A table profile contains a body block");
      let tableOrdinal = tableOrdinalByIndex.get(table.tableIndex);
      if (tableOrdinal === undefined) {
        tableOrdinal = tableOrdinalByIndex.size;
        tableOrdinalByIndex.set(table.tableIndex, tableOrdinal);
      }
      return [
        tableOrdinal,
        table.rowIndex,
        table.cellIndex,
        table.gridColumnIndex,
        table.columnSpan,
        table.rowSpan,
        table.paragraphIndex,
      ];
    },
  });
};

const rowStructureProfile = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
  idStability: (block: Block) => FolioContentIdStability,
): ContentStructureProfile => {
  let physicalCellCount = 0;
  for (const block of blocks) {
    physicalCellCount = Math.max(physicalCellCount, (block.table?.cellIndex ?? -1) + 1);
  }
  return createContentStructureProfile({
    blocks,
    idStability,
    physicalCellCount,
    retainContainerIdentity: true,
    blockStructure: (block) => {
      const table = block.table ?? panic("A row profile contains a body block");
      return [
        table.cellIndex,
        table.gridColumnIndex,
        table.columnSpan,
        table.rowSpan,
        table.paragraphIndex,
      ];
    },
  });
};

type ProfiledContentSequenceItem<Item> = {
  item: Item;
  profile: ContentStructureProfile;
};

type ContentSequenceAlignment<Item> =
  | { type: "pair"; base: Item; revised: Item }
  | { type: "baseOnly"; item: Item }
  | { type: "revisedOnly"; item: Item };

const unpairedContentSequence = <Item>(
  baseItems: readonly Item[],
  revisedItems: readonly Item[],
): ContentSequenceAlignment<Item>[] => [
  ...baseItems.map((item): ContentSequenceAlignment<Item> => ({ type: "baseOnly", item })),
  ...revisedItems.map((item): ContentSequenceAlignment<Item> => ({ type: "revisedOnly", item })),
];

const stableContentSequencePairs = <Item>(
  base: readonly ProfiledContentSequenceItem<Item>[],
  revised: readonly ProfiledContentSequenceItem<Item>[],
): ReadonlySet<number> => {
  const uniqueIndexesById = (
    items: readonly ProfiledContentSequenceItem<Item>[],
  ): ReadonlyMap<string, number | null> => {
    const indexes = new Map<string, number | null>();
    items.forEach(({ profile }, itemIndex) => {
      for (const id of profile.stableIds) {
        const existing = indexes.get(id);
        if (existing === undefined) {
          indexes.set(id, itemIndex);
        } else if (existing !== itemIndex) {
          indexes.set(id, null);
        }
      }
    });
    return indexes;
  };

  const baseIndexes = uniqueIndexesById(base);
  const revisedIndexes = uniqueIndexesById(revised);
  const pairs = new Set<number>();
  for (const [id, baseIndex] of baseIndexes) {
    const revisedIndex = revisedIndexes.get(id);
    if (baseIndex !== null && revisedIndex !== undefined && revisedIndex !== null) {
      pairs.add(baseIndex * revised.length + revisedIndex);
    }
  }
  return pairs;
};

const contentContainerIdentityTransition = (
  base: ContentContainerIdentityProfile | null,
  revised: ContentContainerIdentityProfile | null,
): boolean => {
  if (
    base === null ||
    revised === null ||
    base.status !== "available" ||
    revised.status !== "available" ||
    base.blockIds.length === 0 ||
    base.blockIds.length !== revised.blockIds.length
  ) {
    return false;
  }
  let hasStabilityTransition = false;
  for (let index = 0; index < base.blockIds.length; index++) {
    if (base.blockIds[index] !== revised.blockIds[index]) {
      return false;
    }
    hasStabilityTransition ||= base.idStabilities[index] !== revised.idStabilities[index];
  }
  return hasStabilityTransition;
};

type UniqueExactContentSequencePairsOptions = {
  baseKeys: readonly number[];
  revisedKeys: readonly number[];
};

const uniqueExactContentSequencePairs = ({
  baseKeys,
  revisedKeys,
}: UniqueExactContentSequencePairsOptions): FolioContentBlockPair[] => {
  const uniqueIndexes = (keys: readonly number[]): ReadonlyMap<number, number | null> => {
    const indexes = new Map<number, number | null>();
    keys.forEach((key, index) => {
      if (key === -1) {
        return;
      }
      indexes.set(key, indexes.has(key) ? null : index);
    });
    return indexes;
  };
  const revisedIndexes = uniqueIndexes(revisedKeys);
  const candidates: FolioContentBlockPair[] = [];
  for (const [key, baseIndex] of uniqueIndexes(baseKeys)) {
    const revisedIndex = revisedIndexes.get(key);
    if (baseIndex !== null && revisedIndex !== undefined && revisedIndex !== null) {
      candidates.push({ baseIndex, revisedIndex });
    }
  }
  return candidates;
};

const pairIsCompatibleWithAnchors = (
  pair: FolioContentBlockPair,
  anchors: readonly FolioContentBlockPair[],
): boolean => {
  let lower = 0;
  let upper = anchors.length;
  while (lower < upper) {
    const middle = lower + Math.floor((upper - lower) / 2);
    const anchor = anchors[middle];
    if (anchor !== undefined && anchor.baseIndex < pair.baseIndex) {
      lower = middle + 1;
    } else {
      upper = middle;
    }
  }
  const next = anchors[lower];
  if (next?.baseIndex === pair.baseIndex) {
    return next.revisedIndex === pair.revisedIndex;
  }
  const previous = anchors[lower - 1];
  return (
    (previous === undefined || previous.revisedIndex < pair.revisedIndex) &&
    (next === undefined || pair.revisedIndex < next.revisedIndex)
  );
};

const monotoneContentSequencePairs = (
  candidates: readonly FolioContentBlockPair[],
): FolioContentBlockPair[] =>
  longestIncreasingFolioContentPairs(
    candidates.toSorted(
      (left, right) =>
        left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
    ),
  );

type TrustedContentSequencePairsOptions = {
  exactPairs: readonly FolioContentBlockPair[];
  stablePairs: ReadonlySet<number>;
  revisedLength: number;
  primaryEvidence: "stable" | "exact";
  canPairIndexes: (pair: FolioContentBlockPair) => boolean;
};

const trustedContentSequencePairs = ({
  exactPairs,
  stablePairs,
  revisedLength,
  primaryEvidence,
  canPairIndexes,
}: TrustedContentSequencePairsOptions): FolioContentBlockPair[] => {
  const stableCandidates = [...stablePairs]
    .map((pairIndex) => ({
      baseIndex: Math.floor(pairIndex / revisedLength),
      revisedIndex: pairIndex % revisedLength,
    }))
    .filter(canPairIndexes);
  const allowedExactPairs = exactPairs.filter(canPairIndexes);
  const primaryCandidates = primaryEvidence === "exact" ? allowedExactPairs : stableCandidates;
  const secondaryCandidates = primaryEvidence === "exact" ? stableCandidates : allowedExactPairs;
  const primary = monotoneContentSequencePairs(primaryCandidates);
  const secondary = monotoneContentSequencePairs(
    secondaryCandidates.filter((pair) => pairIsCompatibleWithAnchors(pair, primary)),
  );
  return [...primary, ...secondary].toSorted(
    (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
  );
};

type PersistedContentSequencePairsOptions<Item> = {
  base: readonly ProfiledContentSequenceItem<Item>[];
  revised: readonly ProfiledContentSequenceItem<Item>[];
  anchors: readonly FolioContentBlockPair[];
  canPair: (
    base: ProfiledContentSequenceItem<Item>,
    revised: ProfiledContentSequenceItem<Item>,
  ) => boolean;
};

const persistedContentSequencePairs = <Item>({
  base,
  revised,
  anchors,
  canPair,
}: PersistedContentSequencePairsOptions<Item>): ReadonlySet<number> => {
  if (anchors.length === 0) {
    return new Set();
  }
  const uniqueIndexesByFirstId = (
    items: readonly ProfiledContentSequenceItem<Item>[],
  ): ReadonlyMap<string, number | null> => {
    const indexes = new Map<string, number | null>();
    items.forEach(({ profile }, itemIndex) => {
      const identity = profile.containerIdentity;
      const firstId = identity?.status === "available" ? identity.blockIds.at(0) : undefined;
      if (firstId !== undefined) {
        indexes.set(firstId, indexes.has(firstId) ? null : itemIndex);
      }
    });
    return indexes;
  };
  const baseIndexesByFirstId = uniqueIndexesByFirstId(base);
  const revisedIndexesByFirstId = uniqueIndexesByFirstId(revised);
  const candidates: FolioContentBlockPair[] = [];
  for (const [firstId, baseIndex] of baseIndexesByFirstId) {
    if (baseIndex === null) {
      continue;
    }
    const baseItem = base[baseIndex];
    if (!baseItem) {
      continue;
    }
    const revisedIndex = revisedIndexesByFirstId.get(firstId);
    if (revisedIndex === undefined || revisedIndex === null) {
      continue;
    }
    const revisedItem = revised[revisedIndex];
    const pair = { baseIndex, revisedIndex };
    if (
      revisedItem &&
      canPair(baseItem, revisedItem) &&
      contentContainerIdentityTransition(
        baseItem.profile.containerIdentity,
        revisedItem.profile.containerIdentity,
      ) &&
      pairIsCompatibleWithAnchors(pair, anchors)
    ) {
      candidates.push(pair);
    }
  }
  return new Set(
    monotoneContentSequencePairs(candidates).map(
      ({ baseIndex, revisedIndex }) => baseIndex * revised.length + revisedIndex,
    ),
  );
};

type AlignProfiledContentSequenceOptions<Item> = {
  base: readonly ProfiledContentSequenceItem<Item>[];
  revised: readonly ProfiledContentSequenceItem<Item>[];
  workSession: FolioContentAlignmentWorkSession;
  canPair?: (
    base: ProfiledContentSequenceItem<Item>,
    revised: ProfiledContentSequenceItem<Item>,
  ) => boolean;
  pairSoleStructuralSlot?: boolean;
  primaryEvidence?: "stable" | "exact";
  similarityFactor?:
    | ((
        base: ProfiledContentSequenceItem<Item>,
        revised: ProfiledContentSequenceItem<Item>,
      ) => number)
    | undefined;
};

const ALIGNMENT_DIRECTION = {
  pair: 1,
  baseOnly: 2,
  revisedOnly: 3,
} as const;

const alignProfiledContentSequence = <Item>({
  base,
  revised,
  workSession,
  canPair = () => true,
  pairSoleStructuralSlot = false,
  primaryEvidence = "stable",
  similarityFactor = () => 1,
}: AlignProfiledContentSequenceOptions<Item>): ContentSequenceAlignment<Item>[] => {
  if (base.length === 0 || revised.length === 0) {
    return unpairedContentSequence(
      base.map(({ item }) => item),
      revised.map(({ item }) => item),
    );
  }
  if (
    (base.length !== 1 || revised.length !== 1) &&
    !claimFolioContentAlignmentCells(base.length, revised.length, workSession)
  ) {
    return unpairedContentSequence(
      base.map(({ item }) => item),
      revised.map(({ item }) => item),
    );
  }

  const stablePairs = stableContentSequencePairs(base, revised);
  const provenanceTransitionAtSamePosition = new Uint8Array(Math.min(base.length, revised.length));
  for (let index = 0; index < provenanceTransitionAtSamePosition.length; index++) {
    const baseProfile = base[index]?.profile;
    const revisedProfile = revised[index]?.profile;
    if (!baseProfile || !revisedProfile) {
      continue;
    }
    const hasContainerIdentity =
      baseProfile.containerIdentity !== null || revisedProfile.containerIdentity !== null;
    let hasProvenanceTransition = contentContainerIdentityTransition(
      baseProfile.containerIdentity,
      revisedProfile.containerIdentity,
    );
    if (!hasContainerIdentity) {
      const revisedIds = new Set(revisedProfile.blockIds);
      const baseStableIds = new Set(baseProfile.stableIds);
      const revisedStableIds = new Set(revisedProfile.stableIds);
      // The profile itself is representation-neutral; stableIds records only
      // the stable half of the observed positional-to-stable transition.
      hasProvenanceTransition = baseProfile.blockIds.some(
        (id) => revisedIds.has(id) && baseStableIds.has(id) !== revisedStableIds.has(id),
      );
    }
    if (hasProvenanceTransition) {
      provenanceTransitionAtSamePosition[index] = 1;
    }
  }
  const exactSignatureKeys = new Map<string, number>();
  let nextExactSignatureKey = 0;
  const internExactSignature = (signature: string | null): number => {
    if (signature === null) {
      return -1;
    }
    const existing = exactSignatureKeys.get(signature);
    if (existing !== undefined) {
      return existing;
    }
    const key = nextExactSignatureKey++;
    exactSignatureKeys.set(signature, key);
    return key;
  };
  const baseExactSignatureKeys = base.map(({ profile }) =>
    internExactSignature(profile.exactSignature),
  );
  const revisedExactSignatureKeys = revised.map(({ profile }) =>
    internExactSignature(profile.exactSignature),
  );
  let persistedPairs: ReadonlySet<number> = new Set();
  const usesContainerIdentity =
    base.some(({ profile }) => profile.containerIdentity !== null) ||
    revised.some(({ profile }) => profile.containerIdentity !== null);
  if (usesContainerIdentity) {
    const exactPairs = uniqueExactContentSequencePairs({
      baseKeys: baseExactSignatureKeys,
      revisedKeys: revisedExactSignatureKeys,
    });
    const canPairIndexes = ({
      baseIndex,
      revisedIndex,
    }: FolioContentBlockPair): boolean => {
      const baseItem = base[baseIndex];
      const revisedItem = revised[revisedIndex];
      return baseItem !== undefined && revisedItem !== undefined && canPair(baseItem, revisedItem);
    };
    const trustedPairs = trustedContentSequencePairs({
      exactPairs,
      stablePairs,
      revisedLength: revised.length,
      primaryEvidence,
      canPairIndexes,
    });
    persistedPairs = persistedContentSequencePairs({
      base,
      revised,
      anchors: trustedPairs,
      canPair,
    });
  }
  const maxPairs = Math.min(base.length, revised.length);
  const continuityBonus = 1;
  const secondaryWeight = maxPairs * (CONTENT_STRUCTURE_SIMILARITY_SCALE + continuityBonus) + 1;
  const primaryWeight =
    maxPairs * (secondaryWeight + CONTENT_STRUCTURE_SIMILARITY_SCALE + continuityBonus) + 1;
  const stableWeight = primaryEvidence === "stable" ? primaryWeight : secondaryWeight;
  const exactWeight = primaryEvidence === "exact" ? primaryWeight : secondaryWeight;
  const pairScores = new Float64Array(base.length * revised.length);
  for (let baseIndex = 0; baseIndex < base.length; baseIndex++) {
    const baseItem = base[baseIndex];
    if (!baseItem) {
      continue;
    }
    for (let revisedIndex = 0; revisedIndex < revised.length; revisedIndex++) {
      const revisedItem = revised[revisedIndex];
      if (!revisedItem) {
        continue;
      }
      if (!canPair(baseItem, revisedItem)) {
        continue;
      }
      const pairIndex = baseIndex * revised.length + revisedIndex;
      const exact =
        baseExactSignatureKeys[baseIndex] !== -1 &&
        baseExactSignatureKeys[baseIndex] === revisedExactSignatureKeys[revisedIndex];
      const stable = stablePairs.has(pairIndex);
      const profileSimilarity = exact
        ? 1
        : contentStructureProfileSimilarity(baseItem.profile, revisedItem.profile, workSession);
      const similar = Math.max(
        0,
        Math.min(1, profileSimilarity * similarityFactor(baseItem, revisedItem)),
      );
      const stableAtSamePosition = stable && baseIndex === revisedIndex;
      const persistedAtSamePosition =
        baseIndex === revisedIndex && provenanceTransitionAtSamePosition[baseIndex] === 1;
      const shiftedPersisted = baseIndex !== revisedIndex && persistedPairs.has(pairIndex);
      const persisted = persistedAtSamePosition || shiftedPersisted;
      const soleStructuralSlot =
        pairSoleStructuralSlot && base.length === 1 && revised.length === 1;
      if (
        !exact &&
        !stableAtSamePosition &&
        !persisted &&
        !soleStructuralSlot &&
        similar < CONTENT_STRUCTURE_PAIR_SIMILARITY
      ) {
        continue;
      }
      pairScores[pairIndex] =
        (stable ? stableWeight : 0) +
        (exact ? exactWeight : 0) +
        (shiftedPersisted && !stable && !exact ? secondaryWeight : 0) +
        (persistedAtSamePosition ? continuityBonus : 0) +
        Math.round(similar * CONTENT_STRUCTURE_SIMILARITY_SCALE) +
        (soleStructuralSlot ? 1 : 0);
    }
  }

  if (base.length === 1 && revised.length === 1) {
    const baseItem = base.at(0);
    const revisedItem = revised.at(0);
    if (!baseItem || !revisedItem) {
      return panic("A singleton content sequence has no item");
    }
    if ((pairScores[0] ?? 0) > 0) {
      return [{ type: "pair", base: baseItem.item, revised: revisedItem.item }];
    }
    return unpairedContentSequence(
      base.map(({ item }) => item),
      revised.map(({ item }) => item),
    );
  }

  const directions = new Uint8Array(base.length * revised.length);
  let nextScores = new Float64Array(revised.length + 1);
  let currentScores = new Float64Array(revised.length + 1);
  for (let baseIndex = base.length - 1; baseIndex >= 0; baseIndex--) {
    currentScores[revised.length] = 0;
    for (let revisedIndex = revised.length - 1; revisedIndex >= 0; revisedIndex--) {
      const pairIndex = baseIndex * revised.length + revisedIndex;
      const evidence = pairScores[pairIndex] ?? 0;
      const pair = evidence > 0 ? evidence + (nextScores[revisedIndex + 1] ?? 0) : -1;
      const baseOnly = nextScores[revisedIndex] ?? 0;
      const revisedOnly = currentScores[revisedIndex + 1] ?? 0;
      if (pair >= baseOnly && pair >= revisedOnly) {
        currentScores[revisedIndex] = pair;
        directions[pairIndex] = ALIGNMENT_DIRECTION.pair;
      } else if (baseOnly >= revisedOnly) {
        currentScores[revisedIndex] = baseOnly;
        directions[pairIndex] = ALIGNMENT_DIRECTION.baseOnly;
      } else {
        currentScores[revisedIndex] = revisedOnly;
        directions[pairIndex] = ALIGNMENT_DIRECTION.revisedOnly;
      }
    }
    const completedScores = nextScores;
    nextScores = currentScores;
    currentScores = completedScores;
  }

  const aligned: ContentSequenceAlignment<Item>[] = [];
  let baseIndex = 0;
  let revisedIndex = 0;
  let pairCount = 0;
  let solePairIsShifted = false;
  let hasBaseOnly = false;
  let hasRevisedOnly = false;
  while (baseIndex < base.length && revisedIndex < revised.length) {
    const direction = directions[baseIndex * revised.length + revisedIndex];
    const baseItem = base[baseIndex]?.item;
    const revisedItem = revised[revisedIndex]?.item;
    if (
      direction === ALIGNMENT_DIRECTION.pair &&
      baseItem !== undefined &&
      revisedItem !== undefined
    ) {
      aligned.push({ type: "pair", base: baseItem, revised: revisedItem });
      pairCount += 1;
      solePairIsShifted = baseIndex !== revisedIndex;
      baseIndex += 1;
      revisedIndex += 1;
    } else if (direction === ALIGNMENT_DIRECTION.revisedOnly && revisedItem !== undefined) {
      aligned.push({ type: "revisedOnly", item: revisedItem });
      hasRevisedOnly = true;
      revisedIndex += 1;
    } else if (baseItem !== undefined) {
      aligned.push({ type: "baseOnly", item: baseItem });
      hasBaseOnly = true;
      baseIndex += 1;
    }
  }
  for (const { item } of base.slice(baseIndex)) {
    aligned.push({ type: "baseOnly", item });
    hasBaseOnly = true;
  }
  for (const { item } of revised.slice(revisedIndex)) {
    aligned.push({ type: "revisedOnly", item });
    hasRevisedOnly = true;
  }
  if (pairCount === 1 && solePairIsShifted && hasBaseOnly && hasRevisedOnly) {
    // One content match cannot establish a shifted container mapping when doing so
    // also strands containers on both sides; that shape is equally consistent with
    // content moving between a deletion and an insertion.
    return aligned.flatMap((entry): ContentSequenceAlignment<Item>[] =>
      entry.type === "pair"
        ? [
            { type: "baseOnly", item: entry.base },
            { type: "revisedOnly", item: entry.revised },
          ]
        : [entry],
    );
  }
  return aligned;
};

const rowLocation = <Block extends FolioContentBlock>(
  row: readonly Block[],
): FolioContentTableLocation | null => row.at(0)?.table ?? null;

type AlignRowCellsOptions<Block extends FolioContentBlock> = {
  baseRow: readonly Block[];
  revisedRow: readonly Block[];
  workSession: FolioContentAlignmentWorkSession;
  moveScopeContext: MoveScopeContext;
  stableIdMismatch: "pair" | "separate";
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
  baseColumnKeys?: ReadonlyMap<number, number> | undefined;
  revisedColumnKeys?: ReadonlyMap<number, number> | undefined;
};

const alignRowCells = <Block extends FolioContentBlock>({
  baseRow,
  revisedRow,
  workSession,
  moveScopeContext,
  stableIdMismatch,
  idStability,
  baseColumnKeys,
  revisedColumnKeys,
}: AlignRowCellsOptions<Block>): FolioContentAlignmentStep<Block>[] => {
  const byCell = (
    row: readonly Block[],
    columnKeys: ReadonlyMap<number, number> | undefined,
  ): Map<number, Block[]> => {
    const cells = new Map<number, Block[]>();
    for (const block of row) {
      const table = block.table;
      let cellIndex = table?.cellIndex ?? 0;
      if (table && columnKeys) {
        const alignedColumn = columnKeys.get(table.gridColumnIndex);
        if (alignedColumn === undefined) {
          return panic("A paired table cell has no aligned grid column", {
            gridColumnIndex: table.gridColumnIndex,
          });
        }
        cellIndex = alignedColumn;
      }
      const blocks = cells.get(cellIndex);
      if (blocks) {
        blocks.push(block);
      } else {
        cells.set(cellIndex, [block]);
      }
    }
    return cells;
  };

  const baseCells = byCell(baseRow, baseColumnKeys);
  const revisedCells = byCell(revisedRow, revisedColumnKeys);
  const cellIndexes = [...new Set([...baseCells.keys(), ...revisedCells.keys()])].toSorted(
    (left, right) => left - right,
  );
  const steps: FolioContentAlignmentStep<Block>[] = [];
  for (const cellIndex of cellIndexes) {
    const baseBlocks = baseCells.get(cellIndex) ?? [];
    const revisedBlocks = revisedCells.get(cellIndex) ?? [];
    const aligned = alignFolioContentBlocks(baseBlocks, revisedBlocks, {
      workSession,
      idStability,
      stableIdMismatch,
    }).flatMap((event): FolioContentAlignedBlockEvent<Block>[] => {
      if (
        event.type === "pair" &&
        !contentBlocksShareContainerPath(event.baseBlock, event.revisedBlock)
      ) {
        return [
          { type: "baseOnly", block: event.baseBlock },
          { type: "revisedOnly", block: event.revisedBlock },
        ];
      }
      return [event];
    });
    const bucketByContainerPath = new Map<string | null, number>();
    const bucketForBlock = (block: Block): number => {
      const path = containerPathKeyOf(block);
      const existing = bucketByContainerPath.get(path);
      if (existing !== undefined) {
        return existing;
      }
      const bucket = moveScopeContext.nextTableCellBucket++;
      bucketByContainerPath.set(path, bucket);
      return bucket;
    };
    steps.push(...scopedAlignmentSteps(aligned, moveScopeContext, bucketForBlock));
  }
  return steps;
};

type TableRowAlignment<Block extends FolioContentBlock> =
  | { type: "pair"; baseRow: readonly Block[]; revisedRow: readonly Block[] }
  | { type: "baseOnly"; row: readonly Block[] }
  | { type: "revisedOnly"; row: readonly Block[] };

type PairTableRowsOptions<Block extends FolioContentBlock> = {
  baseRows: readonly Block[][];
  revisedRows: readonly Block[][];
  workSession: FolioContentAlignmentWorkSession;
  idStability: (block: Block) => FolioContentIdStability;
};

const pairTableRows = <Block extends FolioContentBlock>({
  baseRows,
  revisedRows,
  workSession,
  idStability,
}: PairTableRowsOptions<Block>): TableRowAlignment<Block>[] => {
  const profile = (row: readonly Block[]): ProfiledContentSequenceItem<readonly Block[]> => ({
    item: row,
    profile: rowStructureProfile(row, idStability),
  });
  return alignProfiledContentSequence({
    base: baseRows.map(profile),
    revised: revisedRows.map(profile),
    workSession,
    // Once the table itself is paired, its sole row on each side is the same
    // structural slot even when every word in that row changed.
    pairSoleStructuralSlot: true,
    primaryEvidence: "exact",
    similarityFactor: (base, revised) =>
      base.profile.physicalCellCount === revised.profile.physicalCellCount ? 1 : 0.5,
  }).map((alignment): TableRowAlignment<Block> => {
    switch (alignment.type) {
      case "pair":
        return { type: "pair", baseRow: alignment.base, revisedRow: alignment.revised };
      case "baseOnly":
        return { type: "baseOnly", row: alignment.item };
      case "revisedOnly":
        return { type: "revisedOnly", row: alignment.item };
      default: {
        const unreachable: never = alignment;
        return panic("Unhandled content row sequence alignment", { alignment: unreachable });
      }
    }
  });
};

type AlignTableRowsOptions<Block extends FolioContentBlock> = {
  rows: readonly TableRowAlignment<Block>[];
  workSession: FolioContentAlignmentWorkSession;
  moveScopeContext: MoveScopeContext;
  stableIdMismatch: "pair" | "separate";
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
  baseColumnKeys?: ReadonlyMap<number, number> | undefined;
  revisedColumnKeys?: ReadonlyMap<number, number> | undefined;
};

const alignTableRows = <Block extends FolioContentBlock>({
  rows,
  workSession,
  moveScopeContext,
  stableIdMismatch,
  idStability,
  baseColumnKeys,
  revisedColumnKeys,
}: AlignTableRowsOptions<Block>): FolioContentAlignmentStep<Block>[] => {
  const steps: FolioContentAlignmentStep<Block>[] = [];
  const pushRow = (row: readonly Block[], side: "base" | "revised"): void => {
    const location = rowLocation(row);
    if (location) {
      steps.push({ type: side === "base" ? "baseRow" : "revisedRow", blocks: row, location });
    }
  };
  for (const alignment of rows) {
    switch (alignment.type) {
      case "baseOnly":
        pushRow(alignment.row, "base");
        break;
      case "revisedOnly":
        pushRow(alignment.row, "revised");
        break;
      case "pair":
        steps.push(
          ...alignRowCells({
            baseRow: alignment.baseRow,
            revisedRow: alignment.revisedRow,
            workSession,
            moveScopeContext,
            stableIdMismatch,
            idStability,
            baseColumnKeys,
            revisedColumnKeys,
          }),
        );
        break;
      default: {
        const unreachable: never = alignment;
        panic("Unhandled table row alignment", { alignment: unreachable });
      }
    }
  }
  return steps;
};

const rowCellSpansEqual = <Block extends FolioContentBlock>(
  baseRow: readonly Block[],
  revisedRow: readonly Block[],
): boolean => {
  const cells = (row: readonly Block[]): FolioContentTableLocation[] => {
    const byPhysicalIndex = new Map<number, FolioContentTableLocation>();
    for (const block of row) {
      const table = block.table ?? panic("A table row contains a body block");
      if (!byPhysicalIndex.has(table.cellIndex)) {
        byPhysicalIndex.set(table.cellIndex, table);
      }
    }
    return [...byPhysicalIndex.values()];
  };
  const baseCells = cells(baseRow);
  const revisedCells = cells(revisedRow);
  if (baseCells.length !== revisedCells.length) {
    return false;
  }
  return baseCells.every((base, index) => {
    const revised = revisedCells[index];
    return (
      revised !== undefined &&
      base.gridColumnIndex === revised.gridColumnIndex &&
      base.columnSpan === revised.columnSpan &&
      base.rowSpan === revised.rowSpan
    );
  });
};

const rowHasVerticalSpan = <Block extends FolioContentBlock>(row: readonly Block[]): boolean =>
  row.some(({ table }) => table !== undefined && table.rowSpan > 1);

type TableStructurePlan<Block extends FolioContentBlock> = {
  steps: FolioContentAlignmentStep<Block>[];
  representable: boolean;
};

type BuildTablePlanOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  workSession: FolioContentAlignmentWorkSession;
  moveScopeContext: MoveScopeContext;
  stableIdMismatch: "pair" | "separate";
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
};

const buildTablePlan = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  workSession,
  moveScopeContext,
  stableIdMismatch,
  idStability,
}: BuildTablePlanOptions<Block>): TableStructurePlan<Block> => {
  const resolveIdStability = idStability ?? folioContentIdStability;
  const columns = alignTableColumns(baseBlocks, revisedBlocks);
  const rows = pairTableRows({
    baseRows: groupFolioContentTableRows(columns?.baseBlocks ?? baseBlocks),
    revisedRows: groupFolioContentTableRows(columns?.revisedBlocks ?? revisedBlocks),
    workSession,
    idStability: resolveIdStability,
  });
  const representable = rows.every((row) => {
    if (row.type === "pair") {
      return columns !== null || rowCellSpansEqual(row.baseRow, row.revisedRow);
    }
    return !rowHasVerticalSpan(row.row);
  });
  return {
    representable,
    steps: [
      ...(columns?.steps ?? []),
      ...alignTableRows({
        rows,
        workSession,
        moveScopeContext,
        stableIdMismatch,
        idStability: resolveIdStability,
        baseColumnKeys: columns?.baseColumnKeys,
        revisedColumnKeys: columns?.revisedColumnKeys,
      }),
    ],
  };
};

type BuildTableSegmentPlanOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  workSession: FolioContentAlignmentWorkSession;
  moveScopeContext: MoveScopeContext;
  stableIdMismatch: "pair" | "separate";
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
};

const buildTableSegmentPlan = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  workSession,
  moveScopeContext,
  stableIdMismatch,
  idStability,
}: BuildTableSegmentPlanOptions<Block>): TableStructurePlan<Block> => {
  const baseTables = groupTables(baseBlocks);
  const revisedTables = groupTables(revisedBlocks);
  const resolveIdStability = idStability ?? folioContentIdStability;
  const profile = (blocks: Block[]): ProfiledContentSequenceItem<Block[]> => ({
    item: blocks,
    profile: tableStructureProfile(blocks, resolveIdStability),
  });
  const alignedTables = alignProfiledContentSequence({
    base: baseTables.map(profile),
    revised: revisedTables.map(profile),
    workSession,
    // The enclosing outer-table segment already established this container pair.
    pairSoleStructuralSlot: true,
  });
  const steps: FolioContentAlignmentStep<Block>[] = [];
  let representable = true;
  for (const alignment of alignedTables) {
    switch (alignment.type) {
      case "pair": {
        const table = buildTablePlan({
          baseBlocks: alignment.base,
          revisedBlocks: alignment.revised,
          workSession,
          moveScopeContext,
          stableIdMismatch,
          idStability: resolveIdStability,
        });
        steps.push(...table.steps);
        representable &&= table.representable;
        break;
      }
      case "baseOnly":
      case "revisedOnly": {
        const location = alignment.item.at(0)?.table;
        if (location) {
          steps.push({
            type: alignment.type === "baseOnly" ? "baseTable" : "revisedTable",
            blocks: alignment.item,
            location,
          });
        }
        representable = false;
        break;
      }
      default: {
        const unreachable: never = alignment;
        panic("Unhandled nested table sequence alignment", { alignment: unreachable });
      }
    }
  }
  return { steps, representable };
};

type TableDocumentSegment<Block extends FolioContentBlock> = Extract<
  DocumentSegment<Block>,
  { kind: "table" }
>;

const isTableDocumentSegment = <Block extends FolioContentBlock>(
  segment: DocumentSegment<Block>,
): segment is TableDocumentSegment<Block> => segment.kind === "table";

type TableSegmentProfiles<Block extends FolioContentBlock> = {
  items: readonly ProfiledContentSequenceItem<TableDocumentSegment<Block>>[];
  ordinalBySegment: ReadonlyMap<TableDocumentSegment<Block>, number>;
};

const profileTableSegments = <Block extends FolioContentBlock>(
  segments: readonly DocumentSegment<Block>[],
  idStability: (block: Block) => FolioContentIdStability,
): TableSegmentProfiles<Block> => {
  const items: ProfiledContentSequenceItem<TableDocumentSegment<Block>>[] = [];
  const ordinalBySegment = new Map<TableDocumentSegment<Block>, number>();
  for (const segment of segments) {
    if (!isTableDocumentSegment(segment)) {
      continue;
    }
    const profile = tableStructureProfile(segment.blocks, idStability);
    ordinalBySegment.set(segment, items.length);
    items.push({ item: segment, profile });
  }
  return { items, ordinalBySegment };
};

type ContentStructureEvidenceIndex = {
  anchorTextIndexes: ReadonlyMap<string, readonly number[]>;
  exactIndexes: ReadonlyMap<number, readonly number[]>;
  exactKeyByProfile: ReadonlyMap<ContentStructureProfile, number>;
  stableIndexes: ReadonlyMap<string, readonly number[]>;
};

const contentStructureEvidenceIndexes = <Item>(
  base: readonly ProfiledContentSequenceItem<Item>[],
  revised: readonly ProfiledContentSequenceItem<Item>[],
): { base: ContentStructureEvidenceIndex; revised: ContentStructureEvidenceIndex } => {
  const exactKeys = new Map<string, number>();
  const exactKeyByProfile = new Map<ContentStructureProfile, number>();
  let nextExactKey = 0;
  const internExactSignatures = (items: readonly ProfiledContentSequenceItem<Item>[]): void => {
    for (const { profile } of items) {
      if (profile.exactSignature === null) {
        continue;
      }
      let key = exactKeys.get(profile.exactSignature);
      if (key === undefined) {
        key = nextExactKey++;
        exactKeys.set(profile.exactSignature, key);
      }
      exactKeyByProfile.set(profile, key);
    }
  };
  internExactSignatures(base);
  internExactSignatures(revised);

  const index = (
    items: readonly ProfiledContentSequenceItem<Item>[],
  ): ContentStructureEvidenceIndex => {
    const anchorTextIndexes = new Map<string, number[]>();
    const exactIndexes = new Map<number, number[]>();
    const stableIndexes = new Map<string, number[]>();
    items.forEach(({ profile }, itemIndex) => {
      const exactKey = exactKeyByProfile.get(profile);
      if (exactKey !== undefined) {
        const indexes = exactIndexes.get(exactKey);
        if (indexes) {
          indexes.push(itemIndex);
        } else {
          exactIndexes.set(exactKey, [itemIndex]);
        }
      }
      for (const id of profile.stableIds) {
        const indexes = stableIndexes.get(id);
        if (indexes) {
          if (indexes.at(-1) !== itemIndex) {
            indexes.push(itemIndex);
          }
        } else {
          stableIndexes.set(id, [itemIndex]);
        }
      }
      for (const text of profile.anchorTexts) {
        const indexes = anchorTextIndexes.get(text);
        if (indexes) {
          if (indexes.at(-1) !== itemIndex) {
            indexes.push(itemIndex);
          }
        } else {
          anchorTextIndexes.set(text, [itemIndex]);
        }
      }
    });
    return { anchorTextIndexes, exactIndexes, exactKeyByProfile, stableIndexes };
  };

  return { base: index(base), revised: index(revised) };
};

const sortedIndexesLeaveRange = (
  indexes: readonly number[] | undefined,
  start: number,
  end: number,
): boolean =>
  indexes !== undefined && ((indexes.at(0) ?? start) < start || (indexes.at(-1) ?? end - 1) >= end);

const profileHasEvidenceOutsideRange = (
  profile: ContentStructureProfile,
  opposite: ContentStructureEvidenceIndex,
  start: number,
  end: number,
): boolean => {
  const exactKey = opposite.exactKeyByProfile.get(profile);
  if (
    exactKey !== undefined &&
    sortedIndexesLeaveRange(opposite.exactIndexes.get(exactKey), start, end)
  ) {
    return true;
  }
  if (
    profile.anchorTexts.some((text) =>
      sortedIndexesLeaveRange(opposite.anchorTextIndexes.get(text), start, end),
    )
  ) {
    return true;
  }
  return profile.stableIds.some((id) =>
    sortedIndexesLeaveRange(opposite.stableIndexes.get(id), start, end),
  );
};

const profilesShareContentAnchor = (
  base: ContentStructureProfile,
  revised: ContentStructureProfile,
): boolean => {
  const baseTexts = new Set(base.anchorTexts);
  return revised.anchorTexts.some((text) => baseTexts.has(text));
};

type BodyDocumentSegment<Block extends FolioContentBlock> = Extract<
  DocumentSegment<Block>,
  { kind: "body" }
>;

const isBodyDocumentSegment = <Block extends FolioContentBlock>(
  segment: DocumentSegment<Block>,
): segment is BodyDocumentSegment<Block> => segment.kind === "body";

const profileBodySegments = <Block extends FolioContentBlock>(
  segments: readonly DocumentSegment<Block>[],
  idStability: (block: Block) => FolioContentIdStability,
): readonly ProfiledContentSequenceItem<BodyDocumentSegment<Block>>[] =>
  segments.flatMap((segment) =>
    isBodyDocumentSegment(segment)
      ? [
          {
            item: segment,
            profile: createContentStructureProfile({
              blocks: segment.blocks,
              idStability,
              blockStructure: () => [segment.containerPathKey],
            }),
          },
        ]
      : [],
  );

type TrustedBodyPair<Block extends FolioContentBlock> = {
  base: BodyDocumentSegment<Block>;
  revised: BodyDocumentSegment<Block>;
};

const exactBodyPairsInRange = <Block extends FolioContentBlock>(
  base: readonly ProfiledContentSequenceItem<BodyDocumentSegment<Block>>[],
  revised: readonly ProfiledContentSequenceItem<BodyDocumentSegment<Block>>[],
  exactKeyByProfile: ReadonlyMap<ContentStructureProfile, number>,
  baseStart: number,
  baseEnd: number,
  revisedStart: number,
  revisedEnd: number,
): FolioContentBlockPair[] => {
  const uniqueIndexes = (
    items: readonly ProfiledContentSequenceItem<BodyDocumentSegment<Block>>[],
    start: number,
    end: number,
  ): ReadonlyMap<number, number | null> => {
    const indexes = new Map<number, number | null>();
    for (let index = start; index < end; index++) {
      const profile = items[index]?.profile;
      const key = profile ? exactKeyByProfile.get(profile) : undefined;
      if (key === undefined) {
        continue;
      }
      indexes.set(key, indexes.has(key) ? null : index);
    }
    return indexes;
  };
  const baseIndexes = uniqueIndexes(base, baseStart, baseEnd);
  const revisedIndexes = uniqueIndexes(revised, revisedStart, revisedEnd);
  const candidates: FolioContentBlockPair[] = [];
  for (const [key, baseIndex] of baseIndexes) {
    const revisedIndex = revisedIndexes.get(key);
    if (baseIndex !== null && revisedIndex !== undefined && revisedIndex !== null) {
      candidates.push({ baseIndex, revisedIndex });
    }
  }
  return longestIncreasingFolioContentPairs(
    candidates.toSorted(
      (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
    ),
  );
};

const trustedBodyPairs = <Block extends FolioContentBlock>(
  base: readonly ProfiledContentSequenceItem<BodyDocumentSegment<Block>>[],
  revised: readonly ProfiledContentSequenceItem<BodyDocumentSegment<Block>>[],
): TrustedBodyPair<Block>[] => {
  const evidence = contentStructureEvidenceIndexes(base, revised);
  const stableCandidates = [...stableContentSequencePairs(base, revised)]
    .map(
      (pairIndex): FolioContentBlockPair => ({
        baseIndex: Math.floor(pairIndex / revised.length),
        revisedIndex: pairIndex % revised.length,
      }),
    )
    .filter(
      ({ baseIndex, revisedIndex }) =>
        base[baseIndex]?.item.containerPathKey === revised[revisedIndex]?.item.containerPathKey,
    )
    .toSorted(
      (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
    );
  const stablePairs = longestIncreasingFolioContentPairs(stableCandidates);
  const pairs = [...stablePairs];
  let baseStart = 0;
  let revisedStart = 0;
  for (const stable of [...stablePairs, { baseIndex: base.length, revisedIndex: revised.length }]) {
    pairs.push(
      ...exactBodyPairsInRange(
        base,
        revised,
        evidence.base.exactKeyByProfile,
        baseStart,
        stable.baseIndex,
        revisedStart,
        stable.revisedIndex,
      ),
    );
    baseStart = stable.baseIndex + 1;
    revisedStart = stable.revisedIndex + 1;
  }
  return pairs
    .toSorted(
      (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
    )
    .map(({ baseIndex, revisedIndex }) => ({
      base: base[baseIndex]?.item ?? panic("A trusted base body segment is missing"),
      revised: revised[revisedIndex]?.item ?? panic("A trusted revised body segment is missing"),
    }));
};

const segmentGaps = <Block extends FolioContentBlock>(
  segments: readonly DocumentSegment<Block>[],
  anchors: readonly BodyDocumentSegment<Block>[],
): ReadonlyMap<DocumentSegment<Block>, number> => {
  const anchorIndexes = new Map<DocumentSegment<Block>, number>(
    anchors.map((anchor, index) => [anchor, index] as const),
  );
  const gaps = new Map<DocumentSegment<Block>, number>();
  let gap = 0;
  for (const segment of segments) {
    gaps.set(segment, gap);
    if (anchorIndexes.get(segment) === gap) {
      gap += 1;
    }
  }
  return gaps;
};

type PairedTableSegments<Block extends FolioContentBlock> = {
  baseToRevised: ReadonlyMap<TableDocumentSegment<Block>, TableDocumentSegment<Block>>;
  revisedToBase: ReadonlyMap<TableDocumentSegment<Block>, TableDocumentSegment<Block>>;
};

type PairTableSegmentsInGapsOptions<Block extends FolioContentBlock> = {
  base: TableSegmentProfiles<Block>;
  revised: TableSegmentProfiles<Block>;
  baseGaps: ReadonlyMap<DocumentSegment<Block>, number>;
  revisedGaps: ReadonlyMap<DocumentSegment<Block>, number>;
  trustedBodyPairCount: number;
  workSession: FolioContentAlignmentWorkSession;
};

const pairTableSegmentsInGaps = <Block extends FolioContentBlock>({
  base,
  revised,
  baseGaps,
  revisedGaps,
  trustedBodyPairCount,
  workSession,
}: PairTableSegmentsInGapsOptions<Block>): PairedTableSegments<Block> => {
  const groupByGap = (
    items: readonly ProfiledContentSequenceItem<TableDocumentSegment<Block>>[],
    gaps: ReadonlyMap<DocumentSegment<Block>, number>,
  ): ReadonlyMap<number, readonly ProfiledContentSequenceItem<TableDocumentSegment<Block>>[]> => {
    const grouped = new Map<number, ProfiledContentSequenceItem<TableDocumentSegment<Block>>[]>();
    for (const item of items) {
      const gap = gaps.get(item.item) ?? panic("A table segment has no structural gap");
      const entries = grouped.get(gap);
      if (entries) {
        entries.push(item);
      } else {
        grouped.set(gap, [item]);
      }
    }
    return grouped;
  };
  const baseByGap = groupByGap(base.items, baseGaps);
  const revisedByGap = groupByGap(revised.items, revisedGaps);
  const evidence = contentStructureEvidenceIndexes(base.items, revised.items);
  const baseToRevised = new Map<TableDocumentSegment<Block>, TableDocumentSegment<Block>>();
  const revisedToBase = new Map<TableDocumentSegment<Block>, TableDocumentSegment<Block>>();
  for (let gap = 0; gap <= trustedBodyPairCount; gap++) {
    const baseItems = baseByGap.get(gap) ?? [];
    const revisedItems = revisedByGap.get(gap) ?? [];
    if (baseItems.length === 0 || revisedItems.length === 0) {
      continue;
    }
    const baseStart =
      base.ordinalBySegment.get(
        baseItems.at(0)?.item ?? panic("A non-empty base table gap has no first item"),
      ) ?? panic("A base table gap item has no ordinal");
    const revisedStart =
      revised.ordinalBySegment.get(
        revisedItems.at(0)?.item ?? panic("A non-empty revised table gap has no first item"),
      ) ?? panic("A revised table gap item has no ordinal");
    const baseEnd = baseStart + baseItems.length;
    const revisedEnd = revisedStart + revisedItems.length;
    const baseHasExternalEvidence = new Set(
      baseItems
        .filter(({ profile }) =>
          profileHasEvidenceOutsideRange(profile, evidence.revised, revisedStart, revisedEnd),
        )
        .map(({ item }) => item),
    );
    const revisedHasExternalEvidence = new Set(
      revisedItems
        .filter(({ profile }) =>
          profileHasEvidenceOutsideRange(profile, evidence.base, baseStart, baseEnd),
        )
        .map(({ item }) => item),
    );
    const solePairHasContentAnchor =
      baseItems.length === 1 &&
      revisedItems.length === 1 &&
      profilesShareContentAnchor(
        baseItems.at(0)?.profile ?? panic("A sole base table has no profile"),
        revisedItems.at(0)?.profile ?? panic("A sole revised table has no profile"),
      );
    const alignments = alignProfiledContentSequence({
      base: baseItems,
      revised: revisedItems,
      workSession,
      canPair: (baseItem, revisedItem) =>
        !baseHasExternalEvidence.has(baseItem.item) &&
        !revisedHasExternalEvidence.has(revisedItem.item),
      pairSoleStructuralSlot: trustedBodyPairCount > 0 || solePairHasContentAnchor,
    });
    for (const alignment of alignments) {
      if (alignment.type !== "pair") {
        continue;
      }
      baseToRevised.set(alignment.base, alignment.revised);
      revisedToBase.set(alignment.revised, alignment.base);
    }
  }
  return { baseToRevised, revisedToBase };
};

const segmentIndexesByKey = <Block extends FolioContentBlock>(
  segments: readonly DocumentSegment<Block>[],
): ReadonlyMap<string, readonly number[]> => {
  const indexes = new Map<string, number[]>();
  segments.forEach((segment, index) => {
    const key = segment.structuralKey;
    const existing = indexes.get(key);
    if (existing) {
      existing.push(index);
    } else {
      indexes.set(key, [index]);
    }
  });
  return indexes;
};

const firstIndexAfter = (indexes: readonly number[] | undefined, cursor: number): number | null => {
  if (!indexes) {
    return null;
  }
  let low = 0;
  let high = indexes.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if ((indexes[middle] ?? -1) <= cursor) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return indexes[low] ?? null;
};

const segmentsCanPair = <Block extends FolioContentBlock>(
  base: DocumentSegment<Block>,
  revised: DocumentSegment<Block>,
): boolean => {
  if (base.kind !== revised.kind) {
    return false;
  }
  if (base.blocks.length === 0 || revised.blocks.length === 0) {
    return true;
  }
  if (base.containerPathKey === null && revised.containerPathKey === null) {
    return true;
  }
  return base.containerPathKey === revised.containerPathKey;
};

type AlignSegmentsOptions<Block extends FolioContentBlock> = {
  baseSegments: readonly DocumentSegment<Block>[];
  revisedSegments: readonly DocumentSegment<Block>[];
  workSession: FolioContentAlignmentWorkSession;
  idStability: (block: Block) => FolioContentIdStability;
};

const alignSegments = <Block extends FolioContentBlock>({
  baseSegments,
  revisedSegments,
  workSession,
  idStability,
}: AlignSegmentsOptions<Block>): {
  baseSegment: DocumentSegment<Block> | null;
  revisedSegment: DocumentSegment<Block> | null;
}[] => {
  const baseIndexesByKey = segmentIndexesByKey(baseSegments);
  const revisedIndexesByKey = segmentIndexesByKey(revisedSegments);
  const baseIndexBySegment = new Map(
    baseSegments.map((segment, index) => [segment, index] as const),
  );
  const revisedIndexBySegment = new Map(
    revisedSegments.map((segment, index) => [segment, index] as const),
  );
  const bodyPairs = trustedBodyPairs(
    profileBodySegments(baseSegments, idStability),
    profileBodySegments(revisedSegments, idStability),
  );
  const baseGaps = segmentGaps(
    baseSegments,
    bodyPairs.map(({ base }) => base),
  );
  const revisedGaps = segmentGaps(
    revisedSegments,
    bodyPairs.map(({ revised }) => revised),
  );
  const baseTableProfiles = profileTableSegments(baseSegments, idStability);
  const revisedTableProfiles = profileTableSegments(revisedSegments, idStability);
  const tablePairs = pairTableSegmentsInGaps({
    base: baseTableProfiles,
    revised: revisedTableProfiles,
    baseGaps,
    revisedGaps,
    trustedBodyPairCount: bodyPairs.length,
    workSession,
  });
  const baseToRevised = new Map<DocumentSegment<Block>, DocumentSegment<Block>>();
  const revisedToBase = new Map<DocumentSegment<Block>, DocumentSegment<Block>>();
  for (const { base, revised } of bodyPairs) {
    baseToRevised.set(base, revised);
    revisedToBase.set(revised, base);
  }
  for (const [base, revised] of tablePairs.baseToRevised) {
    baseToRevised.set(base, revised);
    revisedToBase.set(revised, base);
  }
  const paired: {
    baseSegment: DocumentSegment<Block> | null;
    revisedSegment: DocumentSegment<Block> | null;
  }[] = [];
  let baseCursor = 0;
  let revisedCursor = 0;
  while (baseCursor < baseSegments.length && revisedCursor < revisedSegments.length) {
    const baseSegment = baseSegments[baseCursor];
    const revisedSegment = revisedSegments[revisedCursor];
    if (!baseSegment || !revisedSegment) {
      break;
    }
    const basePartner = baseToRevised.get(baseSegment);
    const revisedPartner = revisedToBase.get(revisedSegment);
    if (basePartner === revisedSegment && revisedPartner === baseSegment) {
      paired.push({ baseSegment, revisedSegment });
      baseCursor += 1;
      revisedCursor += 1;
      continue;
    }
    if (basePartner) {
      const partnerIndex =
        revisedIndexBySegment.get(basePartner) ??
        panic("A paired revised segment has no document index");
      if (partnerIndex <= revisedCursor || revisedPartner) {
        return panic("Precomputed content segment pairs are not monotone", {
          baseCursor,
          partnerIndex,
          revisedCursor,
        });
      }
      paired.push({ baseSegment: null, revisedSegment });
      revisedCursor += 1;
      continue;
    }
    if (revisedPartner) {
      const partnerIndex =
        baseIndexBySegment.get(revisedPartner) ??
        panic("A paired base segment has no document index");
      if (partnerIndex <= baseCursor) {
        return panic("Precomputed content segment pairs are not monotone", {
          baseCursor,
          partnerIndex,
          revisedCursor,
        });
      }
      paired.push({ baseSegment, revisedSegment: null });
      baseCursor += 1;
      continue;
    }
    if (
      baseSegment.kind === "body" &&
      revisedSegment.kind === "body" &&
      segmentsCanPair(baseSegment, revisedSegment)
    ) {
      paired.push({ baseSegment, revisedSegment });
      baseCursor += 1;
      revisedCursor += 1;
      continue;
    }
    if (!segmentsCanPair(baseSegment, revisedSegment)) {
      const nextBaseIndex = firstIndexAfter(
        baseIndexesByKey.get(revisedSegment.structuralKey),
        baseCursor,
      );
      const nextRevisedIndex = firstIndexAfter(
        revisedIndexesByKey.get(baseSegment.structuralKey),
        revisedCursor,
      );
      const baseDistance =
        nextBaseIndex === null ? Number.POSITIVE_INFINITY : nextBaseIndex - baseCursor;
      const revisedDistance =
        nextRevisedIndex === null ? Number.POSITIVE_INFINITY : nextRevisedIndex - revisedCursor;
      if (revisedDistance < baseDistance) {
        paired.push({ baseSegment: null, revisedSegment });
        revisedCursor += 1;
      } else {
        paired.push({ baseSegment, revisedSegment: null });
        baseCursor += 1;
      }
      continue;
    }
    paired.push({ baseSegment, revisedSegment: null });
    baseCursor += 1;
  }
  for (const segment of baseSegments.slice(baseCursor)) {
    paired.push({ baseSegment: segment, revisedSegment: null });
  }
  for (const segment of revisedSegments.slice(revisedCursor)) {
    paired.push({ baseSegment: null, revisedSegment: segment });
  }
  return paired;
};

const unpairedSegmentSteps = <Block extends FolioContentBlock>(
  segment: DocumentSegment<Block>,
  side: "base" | "revised",
  moveScopeContext: MoveScopeContext,
): FolioContentAlignmentStep<Block>[] => {
  if (segment.kind !== "table") {
    const gap = moveScopeContext.nextGap++;
    return segment.blocks.map((block) =>
      side === "base"
        ? { type: "baseOnly", block, moveScope: { bucket: BODY_MOVE_BUCKET, gap } }
        : { type: "revisedOnly", block, moveScope: { bucket: BODY_MOVE_BUCKET, gap } },
    );
  }
  const location = segment.blocks.at(0)?.table;
  if (!location) {
    return [];
  }
  const outerLocation = { ...location, tableIndex: location.outerTableIndex };
  return [
    {
      type: side === "base" ? "baseTable" : "revisedTable",
      blocks: segment.blocks,
      location: outerLocation,
    },
  ];
};

export type AlignFolioContentStructureOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  workSession?: FolioContentAlignmentWorkSession;
  wholeTableReplacement?: "allow" | "avoid";
  /** Let a structurally aligned slot outrank different stable identifiers. */
  stableIdMismatch?: "pair" | "separate";
  /** @internal Compatibility hook for adapters whose source model predates `idStability`. */
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
};

export const alignFolioContentStructure = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  workSession = createFolioContentAlignmentWorkSession(),
  wholeTableReplacement = "allow",
  stableIdMismatch = "separate",
  idStability,
}: AlignFolioContentStructureOptions<Block>): FolioContentAlignmentStep<Block>[] => {
  const canReplaceWholeTable = wholeTableReplacement === "allow";
  const resolveIdStability = idStability ?? folioContentIdStability;
  const moveScopeContext: MoveScopeContext = {
    nextTableCellBucket: BODY_MOVE_BUCKET + 1,
    nextGap: 0,
  };
  const steps: FolioContentAlignmentStep<Block>[] = [];
  for (const { baseSegment, revisedSegment } of alignSegments({
    baseSegments: splitSegments(baseBlocks),
    revisedSegments: splitSegments(revisedBlocks),
    workSession,
    idStability: resolveIdStability,
  })) {
    if (baseSegment && revisedSegment) {
      if (baseSegment.kind !== "table") {
        steps.push(
          ...scopedAlignmentSteps(
            alignFolioContentBlocks(baseSegment.blocks, revisedSegment.blocks, {
              workSession,
              stableIdMismatch,
              idStability: resolveIdStability,
            }),
            moveScopeContext,
            () => BODY_MOVE_BUCKET,
          ),
        );
        continue;
      }
      const remainingLcsCells = workSession.remainingLcsCells;
      const remainingStructuralTokenLookups = workSession.remainingStructuralTokenLookups;
      const table = buildTableSegmentPlan({
        baseBlocks: baseSegment.blocks,
        revisedBlocks: revisedSegment.blocks,
        workSession,
        moveScopeContext,
        stableIdMismatch,
        idStability: resolveIdStability,
      });
      if (canReplaceWholeTable && !table.representable) {
        workSession.remainingLcsCells = remainingLcsCells;
        workSession.remainingStructuralTokenLookups = remainingStructuralTokenLookups;
        steps.push(...unpairedSegmentSteps(baseSegment, "base", moveScopeContext));
        steps.push(...unpairedSegmentSteps(revisedSegment, "revised", moveScopeContext));
        continue;
      }
      steps.push(...table.steps);
      continue;
    }
    if (baseSegment) {
      steps.push(...unpairedSegmentSteps(baseSegment, "base", moveScopeContext));
      continue;
    }
    if (revisedSegment) {
      steps.push(...unpairedSegmentSteps(revisedSegment, "revised", moveScopeContext));
    }
  }
  return steps;
};

export const contentBlocksShareContainer = (
  left: FolioContentBlock,
  right: FolioContentBlock,
): boolean => {
  if (!contentBlocksShareContainerPath(left, right)) {
    return false;
  }
  if (!left.table || !right.table) {
    return left.table === undefined && right.table === undefined;
  }
  return (
    left.table.tableIndex === right.table.tableIndex &&
    left.table.rowIndex === right.table.rowIndex &&
    left.table.cellIndex === right.table.cellIndex
  );
};
