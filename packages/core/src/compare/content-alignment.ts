/**
 * Pure alignment for representation-neutral content snapshots.
 *
 * The block walk is split by structural container before any textual
 * alignment runs. Body blocks are compared with body blocks; tables are
 * aligned table by table, then row by row and cell by cell. All quadratic
 * structural candidate alignment in one caller-owned work session shares
 * one hard budget; block exact-text anchors are gap-local and non-quadratic.
 */

import { panic } from "better-result";

import { alignParagraphOrdinals, type ParagraphIdentity } from "../docx/paraIdAttribute";
import type {
  FolioContentBlock,
  FolioContentIdStability,
  FolioContentParagraphKind,
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

type BlockPairingTableFacts =
  | { readonly type: "body" }
  | {
      readonly type: "table";
      readonly outerTableIndex: number;
      readonly tableIndex: number;
      readonly rowIndex: number;
      readonly cellIndex: number;
    };

/** Immutable pairing inputs captured once for one block occurrence. */
type BlockPairingFacts = {
  readonly containerPathToken: number;
  readonly id: string;
  readonly idStability: FolioContentIdStability;
  readonly kind: string;
  readonly table: BlockPairingTableFacts;
};

type PreparedAlignmentBlock<Block extends FolioContentBlock> = {
  readonly block: Block;
  readonly index: number;
  readonly pairingFacts: BlockPairingFacts;
};

const NON_WHITESPACE = /\S/u;

type BlockAlignmentStructuralScope = "document" | "pairedTableCell";

const PARAGRAPH_MARK_KIND_FAMILY = {
  heading: "paragraphMark",
  listItem: "paragraphMark",
  paragraph: "paragraphMark",
} as const satisfies Record<FolioContentParagraphKind, "paragraphMark">;

const blockKindsCanPair = (baseKind: string, revisedKind: string): boolean =>
  baseKind === revisedKind ||
  (Object.hasOwn(PARAGRAPH_MARK_KIND_FAMILY, baseKind) &&
    Object.hasOwn(PARAGRAPH_MARK_KIND_FAMILY, revisedKind));

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

const containerPathKeyOf = ({ containerPath }: FolioContentBlock): string | null =>
  containerPath === undefined || containerPath.length === 0
    ? null
    : JSON.stringify(containerPath.map(({ kind, id }) => [kind, id]));

const blockPairingTableFactsOf = ({ table }: FolioContentBlock): BlockPairingTableFacts => {
  if (!table) {
    return { type: "body" };
  }
  return {
    type: "table",
    outerTableIndex: table.outerTableIndex,
    tableIndex: table.tableIndex,
    rowIndex: table.rowIndex,
    cellIndex: table.cellIndex,
  };
};

type PrepareAlignmentBlocksOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  idStability: (block: Block) => FolioContentIdStability;
};

const prepareAlignmentBlocks = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  idStability,
}: PrepareAlignmentBlocksOptions<Block>): {
  base: readonly PreparedAlignmentBlock<Block>[];
  revised: readonly PreparedAlignmentBlock<Block>[];
} => {
  const ROOT_CONTAINER_PATH_TOKEN = 0;
  const containerPathTokens = new Map<string, number>();
  let nextContainerPathToken = ROOT_CONTAINER_PATH_TOKEN + 1;
  const containerPathTokenOf = (block: Block): number => {
    const key = containerPathKeyOf(block);
    if (key === null) {
      return ROOT_CONTAINER_PATH_TOKEN;
    }
    const existing = containerPathTokens.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const token = nextContainerPathToken++;
    containerPathTokens.set(key, token);
    return token;
  };
  const prepare = (blocks: readonly Block[]): PreparedAlignmentBlock<Block>[] =>
    blocks.map((block, index) => ({
      block,
      index,
      pairingFacts: {
        containerPathToken: containerPathTokenOf(block),
        id: block.id,
        idStability: idStability(block),
        kind: block.kind,
        table: blockPairingTableFactsOf(block),
      },
    }));
  return { base: prepare(baseBlocks), revised: prepare(revisedBlocks) };
};

type BlocksCanPairOptions = {
  base: BlockPairingFacts;
  revised: BlockPairingFacts;
  stableIdMismatch: "pair" | "separate";
  structuralScope: BlockAlignmentStructuralScope;
};

/** Every alignment pass shares this complete structural pairing policy. */
const blocksCanPair = ({
  base,
  revised,
  stableIdMismatch,
  structuralScope,
}: BlocksCanPairOptions): boolean => {
  if (
    !blockKindsCanPair(base.kind, revised.kind) ||
    base.containerPathToken !== revised.containerPathToken
  ) {
    return false;
  }
  if (base.table.type !== revised.table.type) {
    return false;
  }
  if (
    structuralScope === "document" &&
    base.table.type === "table" &&
    revised.table.type === "table" &&
    (base.table.outerTableIndex !== revised.table.outerTableIndex ||
      base.table.tableIndex !== revised.table.tableIndex ||
      base.table.rowIndex !== revised.table.rowIndex ||
      base.table.cellIndex !== revised.table.cellIndex)
  ) {
    return false;
  }
  return !(
    stableIdMismatch === "separate" &&
    base.idStability === "stable" &&
    revised.idStability === "stable" &&
    base.id !== revised.id
  );
};

const uniqueStableBlocks = <Block extends FolioContentBlock>(
  blocks: readonly PreparedAlignmentBlock<Block>[],
): ReadonlyMap<string, PreparedAlignmentBlock<Block> | null> => {
  const blocksById = new Map<string, PreparedAlignmentBlock<Block> | null>();
  for (const block of blocks) {
    if (block.pairingFacts.idStability !== "stable") {
      continue;
    }
    blocksById.set(block.pairingFacts.id, blocksById.has(block.pairingFacts.id) ? null : block);
  }
  return blocksById;
};

type PairByStableIdOptions<Block extends FolioContentBlock> = {
  base: readonly PreparedAlignmentBlock<Block>[];
  revised: readonly PreparedAlignmentBlock<Block>[];
  canPair: (
    baseBlock: PreparedAlignmentBlock<Block>,
    revisedBlock: PreparedAlignmentBlock<Block>,
  ) => boolean;
};

/** Stable identity is authoritative only when it is unique on both sides. */
const pairByStableId = <Block extends FolioContentBlock>({
  base,
  revised,
  canPair,
}: PairByStableIdOptions<Block>): FolioContentBlockPair[] => {
  const baseBlockById = uniqueStableBlocks(base);
  const revisedBlockById = uniqueStableBlocks(revised);

  const candidates: FolioContentBlockPair[] = [];
  for (const baseBlock of base) {
    if (baseBlockById.get(baseBlock.pairingFacts.id) !== baseBlock) {
      continue;
    }
    const revisedBlock = revisedBlockById.get(baseBlock.pairingFacts.id);
    if (!revisedBlock) {
      continue;
    }
    if (canPair(baseBlock, revisedBlock)) {
      candidates.push({ baseIndex: baseBlock.index, revisedIndex: revisedBlock.index });
    }
  }
  return longestIncreasingFolioContentPairs(candidates);
};

type PairByResidualIdContinuityOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly PreparedAlignmentBlock<Block>[];
  revisedBlocks: readonly PreparedAlignmentBlock<Block>[];
  baseFrom: number;
  baseTo: number;
  revisedFrom: number;
  revisedTo: number;
  canPair: (
    baseBlock: PreparedAlignmentBlock<Block>,
    revisedBlock: PreparedAlignmentBlock<Block>,
  ) => boolean;
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
  canPair,
}: PairByResidualIdContinuityOptions<Block>): FolioContentBlockPair[] => {
  const uniqueBlocksById = ({
    blocks,
    from,
    to,
  }: {
    blocks: readonly PreparedAlignmentBlock<Block>[];
    from: number;
    to: number;
  }): ReadonlyMap<string, PreparedAlignmentBlock<Block> | null> => {
    const blocksById = new Map<string, PreparedAlignmentBlock<Block> | null>();
    for (let index = from; index < to; index++) {
      const block = blocks[index];
      if (!block) {
        continue;
      }
      blocksById.set(block.pairingFacts.id, blocksById.has(block.pairingFacts.id) ? null : block);
    }
    return blocksById;
  };

  const baseBlocksById = uniqueBlocksById({ blocks: baseBlocks, from: baseFrom, to: baseTo });
  const revisedBlocksById = uniqueBlocksById({
    blocks: revisedBlocks,
    from: revisedFrom,
    to: revisedTo,
  });
  const candidates: FolioContentBlockPair[] = [];
  for (const [id, baseBlock] of baseBlocksById) {
    const revisedBlock = revisedBlocksById.get(id);
    if (!baseBlock || !revisedBlock) {
      continue;
    }
    const baseStability = baseBlock.pairingFacts.idStability;
    const revisedStability = revisedBlock.pairingFacts.idStability;
    if (baseStability === revisedStability || !canPair(baseBlock, revisedBlock)) {
      continue;
    }
    candidates.push({ baseIndex: baseBlock.index, revisedIndex: revisedBlock.index });
  }
  return longestIncreasingFolioContentPairs(candidates);
};

type PairByUniqueExactTextOptions<Block extends FolioContentBlock> = {
  base: readonly PreparedAlignmentBlock<Block>[];
  revised: readonly PreparedAlignmentBlock<Block>[];
  baseFrom: number;
  baseTo: number;
  revisedFrom: number;
  revisedTo: number;
  canPair: (
    baseBlock: PreparedAlignmentBlock<Block>,
    revisedBlock: PreparedAlignmentBlock<Block>,
  ) => boolean;
};

/** Exact anchors are gap-local; repeated or blank text is not identity evidence. */
const pairByUniqueExactText = <Block extends FolioContentBlock>({
  base,
  revised,
  baseFrom,
  baseTo,
  revisedFrom,
  revisedTo,
  canPair,
}: PairByUniqueExactTextOptions<Block>): FolioContentBlockPair[] => {
  const uniqueBlocks = (
    blocks: readonly PreparedAlignmentBlock<Block>[],
    from: number,
    to: number,
  ): ReadonlyMap<string, PreparedAlignmentBlock<Block> | null> => {
    const blocksByText = new Map<string, PreparedAlignmentBlock<Block> | null>();
    for (let index = from; index < to; index++) {
      const block = blocks[index];
      if (!block) {
        continue;
      }
      const text = block.block.text;
      if (!NON_WHITESPACE.test(text)) {
        continue;
      }
      blocksByText.set(text, blocksByText.has(text) ? null : block);
    }
    return blocksByText;
  };
  const baseBlocksByText = uniqueBlocks(base, baseFrom, baseTo);
  const revisedBlocksByText = uniqueBlocks(revised, revisedFrom, revisedTo);
  const candidates: FolioContentBlockPair[] = [];
  for (const [text, baseBlock] of baseBlocksByText) {
    const revisedBlock = revisedBlocksByText.get(text);
    if (!baseBlock || !revisedBlock) {
      continue;
    }
    if (canPair(baseBlock, revisedBlock)) {
      candidates.push({ baseIndex: baseBlock.index, revisedIndex: revisedBlock.index });
    }
  }
  const orderedCandidates = candidates.toSorted(
    (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
  );
  if (orderedCandidates.length === 0) {
    return [];
  }

  // Minimum-tail replacement preserves the previous exact-LCS tie break: for
  // crossing equal-length subsequences, skip the earlier base candidate.
  const tailRevisedIndexes: number[] = [];
  const tailCandidateIndexes: number[] = [];
  const predecessors = new Int32Array(orderedCandidates.length).fill(-1);
  orderedCandidates.forEach((candidate, candidateIndex) => {
    let lower = 0;
    let upper = tailRevisedIndexes.length;
    while (lower < upper) {
      const middle = lower + Math.floor((upper - lower) / 2);
      if ((tailRevisedIndexes[middle] ?? Number.POSITIVE_INFINITY) < candidate.revisedIndex) {
        lower = middle + 1;
      } else {
        upper = middle;
      }
    }
    predecessors[candidateIndex] = lower === 0 ? -1 : (tailCandidateIndexes[lower - 1] ?? -1);
    tailRevisedIndexes[lower] = candidate.revisedIndex;
    tailCandidateIndexes[lower] = candidateIndex;
  });
  const pairs: FolioContentBlockPair[] = [];
  for (
    let candidateIndex = tailCandidateIndexes.at(-1) ?? -1;
    candidateIndex !== -1;
    candidateIndex = predecessors[candidateIndex] ?? -1
  ) {
    const candidate = orderedCandidates[candidateIndex];
    if (candidate) {
      pairs.push(candidate);
    }
  }
  return pairs.toReversed();
};

type PairsInAnchorGapsOptions = {
  baseLength: number;
  revisedLength: number;
  anchors: readonly FolioContentBlockPair[];
  pairGap: (
    baseFrom: number,
    baseTo: number,
    revisedFrom: number,
    revisedTo: number,
  ) => readonly FolioContentBlockPair[];
};

const pairsInAnchorGaps = ({
  baseLength,
  revisedLength,
  anchors,
  pairGap,
}: PairsInAnchorGapsOptions): FolioContentBlockPair[] => {
  const pairs: FolioContentBlockPair[] = [];
  let baseFrom = 0;
  let revisedFrom = 0;
  for (const anchor of [...anchors, { baseIndex: baseLength, revisedIndex: revisedLength }]) {
    pairs.push(...pairGap(baseFrom, anchor.baseIndex, revisedFrom, anchor.revisedIndex));
    baseFrom = anchor.baseIndex + 1;
    revisedFrom = anchor.revisedIndex + 1;
  }
  return pairs;
};

type CrossedExactTextBlocksOptions<Block extends FolioContentBlock> = {
  base: readonly PreparedAlignmentBlock<Block>[];
  revised: readonly PreparedAlignmentBlock<Block>[];
  anchors: readonly FolioContentBlockPair[];
  canPair: (
    baseBlock: PreparedAlignmentBlock<Block>,
    revisedBlock: PreparedAlignmentBlock<Block>,
  ) => boolean;
};

/**
 * Blocks the anchor passes left unpaired whose exact text stands, also
 * unpaired, on the other side.
 *
 * Anchoring is monotone, so content that moved past other content leaves one
 * of its two correspondences unanchored by construction: two paragraphs that
 * swap places produce two exact candidates that cross, and only one survives
 * the increasing subsequence. What is left over is both ends of that crossing.
 * It is not a paragraph rewritten into another, so the positional fallback
 * must not fuse them: a relocation would come back as a replacement neither
 * document contains, and the move pass would never see the removal and the
 * arrival it is there to pair.
 */
const crossedExactTextBlocks = <Block extends FolioContentBlock>({
  base,
  revised,
  anchors,
  canPair,
}: CrossedExactTextBlocksOptions<Block>): {
  base: ReadonlySet<number>;
  revised: ReadonlySet<number>;
} => {
  const uniqueUnanchoredByText = (
    blocks: readonly PreparedAlignmentBlock<Block>[],
    anchored: ReadonlySet<number>,
  ): ReadonlyMap<string, PreparedAlignmentBlock<Block> | null> => {
    const blocksByText = new Map<string, PreparedAlignmentBlock<Block> | null>();
    for (const block of blocks) {
      // Repeated or blank text is not identity evidence, exactly as it is not
      // for the exact-text anchors.
      if (anchored.has(block.index) || !NON_WHITESPACE.test(block.block.text)) {
        continue;
      }
      const { text } = block.block;
      blocksByText.set(text, blocksByText.has(text) ? null : block);
    }
    return blocksByText;
  };
  const baseByText = uniqueUnanchoredByText(
    base,
    new Set(anchors.map(({ baseIndex }) => baseIndex)),
  );
  const revisedByText = uniqueUnanchoredByText(
    revised,
    new Set(anchors.map(({ revisedIndex }) => revisedIndex)),
  );
  const crossedBase = new Set<number>();
  const crossedRevised = new Set<number>();
  for (const [text, baseBlock] of baseByText) {
    const revisedBlock = revisedByText.get(text);
    if (!baseBlock || !revisedBlock || !canPair(baseBlock, revisedBlock)) {
      continue;
    }
    crossedBase.add(baseBlock.index);
    crossedRevised.add(revisedBlock.index);
  }
  return { base: crossedBase, revised: crossedRevised };
};

/**
 * Minimum multiset Dice similarity for two blocks to pair inside a gap. At 0.5 a paragraph still pairs after gaining up to
 * twice its own length (2n / (n + 3n)); below it, more of the pair would read
 * as changed than kept, which a removal beside an insertion says better.
 */
const GAP_PAIR_SIMILARITY_THRESHOLD = 0.5;

/**
 * Gap similarities are compared as integers, so a tie is exact rather than an
 * accident of floating-point summation, and a tie-break can sit below the
 * smallest similarity step.
 */
const GAP_PAIR_SIMILARITY_SCALE = 1_000;

type GapBlockTokens = { counts: ReadonlyMap<string, number>; total: number };

/**
 * Words as case-folded runs of letters, marks and digits: punctuation glued to
 * a word ("paragraph." against "paragraph") and a capital at a sentence start
 * would otherwise count a kept word as changed.
 */
const gapBlockTokens = (text: string): GapBlockTokens => {
  const counts = new Map<string, number>();
  let total = 0;
  for (const match of text.toLowerCase().matchAll(/[\p{L}\p{M}\p{N}]+/gu)) {
    total += 1;
    counts.set(match[0], (counts.get(match[0]) ?? 0) + 1);
  }
  return { counts, total };
};

type GapSimilarity = { status: "measured"; value: number } | { status: "budget-exceeded" };

const gapBlockSimilarity = (
  base: GapBlockTokens,
  revised: GapBlockTokens,
  workSession: FolioContentAlignmentWorkSession,
): GapSimilarity => {
  // A block without words (an empty paragraph) offers no wording to weigh, so
  // it pairs at the threshold: filling an empty paragraph keeps its mark,
  // where a deletion beside an insertion would leave a blank one behind.
  if (base.total === 0 && revised.total === 0) {
    return { status: "measured", value: 1 };
  }
  if (base.total === 0 || revised.total === 0) {
    return { status: "measured", value: GAP_PAIR_SIMILARITY_THRESHOLD };
  }
  const [tokens, counterparts] =
    base.counts.size <= revised.counts.size
      ? [base.counts, revised.counts]
      : [revised.counts, base.counts];
  if (tokens.size > workSession.remainingStructuralTokenLookups) {
    return { status: "budget-exceeded" };
  }
  workSession.remainingStructuralTokenLookups -= tokens.size;
  let shared = 0;
  for (const [token, count] of tokens) {
    shared += Math.min(count, counterparts.get(token) ?? 0);
  }
  return { status: "measured", value: (2 * shared) / (base.total + revised.total) };
};

type PairGapBySimilarityOptions<Block extends FolioContentBlock> = {
  base: readonly PreparedAlignmentBlock<Block>[];
  revised: readonly PreparedAlignmentBlock<Block>[];
  canPair: (base: PreparedAlignmentBlock<Block>, revised: PreparedAlignmentBlock<Block>) => boolean;
  workSession: FolioContentAlignmentWorkSession;
};

type GapSimilarityPairing = {
  pairs: FolioContentBlockPair[];
  /** Gap offsets of the blocks with at least one pairing candidate. */
  candidateBase: ReadonlySet<number>;
  candidateRevised: ReadonlySet<number>;
};

/**
 * The order-preserving pairs of one gap that maximise their summed
 * similarity, each at least `GAP_PAIR_SIMILARITY_THRESHOLD`; offsets are into
 * the gap's slices. Pairing by position instead fuses an inserted block with
 * the neighbour it pushed down, and every block after it with the next one's
 * wording. Null when the work budget refuses the gap.
 */
const pairGapBySimilarity = <Block extends FolioContentBlock>({
  base,
  revised,
  canPair,
  workSession,
}: PairGapBySimilarityOptions<Block>): GapSimilarityPairing | null => {
  const baseCount = base.length;
  const revisedCount = revised.length;
  if (!claimFolioContentAlignmentCells(baseCount, revisedCount, workSession)) {
    return null;
  }
  const revisedTokens = revised.map((block) => gapBlockTokens(block.block.text));
  // Equal display labels break ties between equally similar pairs: repeated
  // wording ("Intentionally omitted.") leaves the label as a block's only identity. The
  // bonuses of every pair together stay below one similarity step.
  const similarityStep = Math.min(baseCount, revisedCount) + 1;
  // -1 marks a cell that may not pair.
  const similarity = new Float64Array(baseCount * revisedCount).fill(-1);
  const candidateBase = new Set<number>();
  const candidateRevised = new Set<number>();
  for (const [baseOffset, baseBlock] of base.entries()) {
    const baseTokens = gapBlockTokens(baseBlock.block.text);
    for (const [revisedOffset, revisedBlock] of revised.entries()) {
      if (!canPair(baseBlock, revisedBlock)) {
        continue;
      }
      const tokens = revisedTokens[revisedOffset] ?? panic("A gap block has no token profile");
      const measured =
        baseBlock.block.text === revisedBlock.block.text
          ? ({ status: "measured", value: 1 } as const)
          : gapBlockSimilarity(baseTokens, tokens, workSession);
      if (measured.status === "budget-exceeded") {
        return null;
      }
      if (measured.value >= GAP_PAIR_SIMILARITY_THRESHOLD) {
        const sameLabel =
          baseBlock.block.displayLabel !== undefined &&
          baseBlock.block.displayLabel === revisedBlock.block.displayLabel;
        similarity[baseOffset * revisedCount + revisedOffset] =
          Math.round(measured.value * GAP_PAIR_SIMILARITY_SCALE) * similarityStep +
          (sameLabel ? 1 : 0);
        candidateBase.add(baseOffset);
        candidateRevised.add(revisedOffset);
      }
    }
  }

  // Scores over suffixes, walked forward: of equally good alignments the walk
  // takes the earliest pair, which is the pairing by position when one exists.
  const width = revisedCount + 1;
  const scores = new Float64Array((baseCount + 1) * width);
  const cellSimilarity = (baseOffset: number, revisedOffset: number): number =>
    similarity[baseOffset * revisedCount + revisedOffset] ?? -1;
  const score = (row: number, column: number): number => scores[row * width + column] ?? 0;
  for (let row = baseCount - 1; row >= 0; row--) {
    for (let column = revisedCount - 1; column >= 0; column--) {
      const cell = cellSimilarity(row, column);
      scores[row * width + column] = Math.max(
        score(row + 1, column),
        score(row, column + 1),
        cell < 0 ? 0 : score(row + 1, column + 1) + cell,
      );
    }
  }

  const pairs: FolioContentBlockPair[] = [];
  let row = 0;
  let column = 0;
  while (row < baseCount && column < revisedCount) {
    const cell = cellSimilarity(row, column);
    if (cell >= 0 && score(row, column) === score(row + 1, column + 1) + cell) {
      pairs.push({ baseIndex: row, revisedIndex: column });
      row += 1;
      column += 1;
    } else if (score(row, column) === score(row + 1, column)) {
      row += 1;
    } else {
      column += 1;
    }
  }
  return { pairs, candidateBase, candidateRevised };
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

type AlignFolioContentBlocksInScopeOptions<Block extends FolioContentBlock> =
  AlignFolioContentBlocksOptions<Block> & {
    structuralScope: BlockAlignmentStructuralScope;
  };

const alignFolioContentBlocksInScope = <Block extends FolioContentBlock>(
  baseBlocks: readonly Block[],
  revisedBlocks: readonly Block[],
  options: AlignFolioContentBlocksInScopeOptions<Block>,
): FolioContentAlignedBlockEvent<Block>[] => {
  const stableIdMismatch = options.stableIdMismatch ?? "separate";
  const idStability = options.idStability ?? folioContentIdStability;
  const workSession = options.workSession ?? createFolioContentAlignmentWorkSession();
  const prepared = prepareAlignmentBlocks({ baseBlocks, revisedBlocks, idStability });
  const canPair = (
    baseBlock: PreparedAlignmentBlock<Block>,
    revisedBlock: PreparedAlignmentBlock<Block>,
  ): boolean =>
    blocksCanPair({
      base: baseBlock.pairingFacts,
      revised: revisedBlock.pairingFacts,
      stableIdMismatch,
      structuralScope: options.structuralScope,
    });
  const stableIdAnchors = pairByStableId({
    base: prepared.base,
    revised: prepared.revised,
    canPair,
  });
  const exactTextAnchors = pairsInAnchorGaps({
    baseLength: prepared.base.length,
    revisedLength: prepared.revised.length,
    anchors: stableIdAnchors,
    pairGap: (baseFrom, baseTo, revisedFrom, revisedTo) =>
      pairByUniqueExactText({
        base: prepared.base,
        revised: prepared.revised,
        baseFrom,
        baseTo,
        revisedFrom,
        revisedTo,
        canPair,
      }),
  });
  const exactAndStableAnchors = [...stableIdAnchors, ...exactTextAnchors].toSorted(
    (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
  );
  const continuityAnchors = pairsInAnchorGaps({
    baseLength: prepared.base.length,
    revisedLength: prepared.revised.length,
    anchors: exactAndStableAnchors,
    pairGap: (baseFrom, baseTo, revisedFrom, revisedTo) =>
      pairByResidualIdContinuity({
        baseBlocks: prepared.base,
        revisedBlocks: prepared.revised,
        baseFrom,
        baseTo,
        revisedFrom,
        revisedTo,
        canPair,
      }),
  });
  const anchors = [...exactAndStableAnchors, ...continuityAnchors].toSorted(
    (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
  );
  const crossed = crossedExactTextBlocks({
    base: prepared.base,
    revised: prepared.revised,
    anchors,
    canPair,
  });
  const events: FolioContentAlignedBlockEvent<Block>[] = [];
  const fusesACrossing = (
    baseBlock: PreparedAlignmentBlock<Block>,
    revisedBlock: PreparedAlignmentBlock<Block>,
  ): boolean =>
    baseBlock.block.text !== revisedBlock.block.text &&
    crossed.base.has(baseBlock.index) &&
    crossed.revised.has(revisedBlock.index);

  const emitPositionalGap = (
    baseFrom: number,
    baseTo: number,
    revisedFrom: number,
    revisedTo: number,
  ): void => {
    const pairedCount = Math.min(baseTo - baseFrom, revisedTo - revisedFrom);
    for (let offset = 0; offset < pairedCount; offset++) {
      const baseBlock = prepared.base[baseFrom + offset];
      const revisedBlock = prepared.revised[revisedFrom + offset];
      if (baseBlock && revisedBlock) {
        if (fusesACrossing(baseBlock, revisedBlock) || !canPair(baseBlock, revisedBlock)) {
          events.push({ type: "baseOnly", block: baseBlock.block });
          events.push({ type: "revisedOnly", block: revisedBlock.block });
        } else {
          events.push({
            type: "pair",
            baseBlock: baseBlock.block,
            revisedBlock: revisedBlock.block,
          });
        }
      }
    }
    for (let index = baseFrom + pairedCount; index < baseTo; index++) {
      const block = prepared.base[index]?.block;
      if (block) {
        events.push({ type: "baseOnly", block });
      }
    }
    for (let index = revisedFrom + pairedCount; index < revisedTo; index++) {
      const block = prepared.revised[index]?.block;
      if (block) {
        events.push({ type: "revisedOnly", block });
      }
    }
  };

  /**
   * A gap's blocks pair by similarity, whether or not its sides are equal in
   * length: an equal gap can hide an insertion beside a deletion, and pairing
   * it by position reads each kept block as a rewrite of its neighbour. The
   * similar pairs anchor the rest; blocks between two anchors pair by position
   * only when neither side offered any candidate and the counts match, which
   * reads a block rewritten beyond recognition as the modification it is.
   */
  const emitGap = (
    baseFrom: number,
    baseTo: number,
    revisedFrom: number,
    revisedTo: number,
  ): void => {
    const pairing = pairGapBySimilarity({
      base: prepared.base.slice(baseFrom, baseTo),
      revised: prepared.revised.slice(revisedFrom, revisedTo),
      canPair: (baseBlock, revisedBlock) =>
        !fusesACrossing(baseBlock, revisedBlock) && canPair(baseBlock, revisedBlock),
      workSession,
    });
    if (pairing === null) {
      emitPositionalGap(baseFrom, baseTo, revisedFrom, revisedTo);
      return;
    }
    const { pairs, candidateBase, candidateRevised } = pairing;
    let baseCursor = baseFrom;
    let revisedCursor = revisedFrom;
    const hasCandidate = (candidates: ReadonlySet<number>, from: number, to: number): boolean => {
      for (let offset = from; offset < to; offset++) {
        if (candidates.has(offset)) {
          return true;
        }
      }
      return false;
    };
    const emitUnpairedUntil = (baseEnd: number, revisedEnd: number): void => {
      if (
        baseEnd - baseCursor === revisedEnd - revisedCursor &&
        !hasCandidate(candidateBase, baseCursor - baseFrom, baseEnd - baseFrom) &&
        !hasCandidate(candidateRevised, revisedCursor - revisedFrom, revisedEnd - revisedFrom)
      ) {
        emitPositionalGap(baseCursor, baseEnd, revisedCursor, revisedEnd);
        baseCursor = baseEnd;
        revisedCursor = revisedEnd;
        return;
      }
      for (; baseCursor < baseEnd; baseCursor++) {
        const block = prepared.base[baseCursor]?.block;
        if (block) {
          events.push({ type: "baseOnly", block });
        }
      }
      for (; revisedCursor < revisedEnd; revisedCursor++) {
        const block = prepared.revised[revisedCursor]?.block;
        if (block) {
          events.push({ type: "revisedOnly", block });
        }
      }
    };
    for (const pair of pairs) {
      emitUnpairedUntil(baseFrom + pair.baseIndex, revisedFrom + pair.revisedIndex);
      const baseBlock = prepared.base[baseCursor]?.block;
      const revisedBlock = prepared.revised[revisedCursor]?.block;
      if (baseBlock && revisedBlock) {
        events.push({ type: "pair", baseBlock, revisedBlock });
      }
      baseCursor += 1;
      revisedCursor += 1;
    }
    emitUnpairedUntil(baseTo, revisedTo);
  };

  let baseCursor = 0;
  let revisedCursor = 0;
  for (const anchor of anchors) {
    emitGap(baseCursor, anchor.baseIndex, revisedCursor, anchor.revisedIndex);
    const baseBlock = prepared.base[anchor.baseIndex]?.block;
    const revisedBlock = prepared.revised[anchor.revisedIndex]?.block;
    if (baseBlock && revisedBlock) {
      events.push({ type: "pair", baseBlock, revisedBlock });
    }
    baseCursor = anchor.baseIndex + 1;
    revisedCursor = anchor.revisedIndex + 1;
  }
  emitGap(baseCursor, prepared.base.length, revisedCursor, prepared.revised.length);
  return events;
};

export const alignFolioContentBlocks = <Block extends FolioContentBlock>(
  baseBlocks: readonly Block[],
  revisedBlocks: readonly Block[],
  options: AlignFolioContentBlocksOptions<Block> = {},
): FolioContentAlignedBlockEvent<Block>[] =>
  alignFolioContentBlocksInScope(baseBlocks, revisedBlocks, {
    ...options,
    structuralScope: "document",
  });

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

/**
 * Which uninterrupted run of its own cell each block belongs to.
 *
 * A cell holds paragraphs and nested tables, and the nested table splits the
 * cell's paragraphs into runs: the ones before it and the ones after it are
 * separate sequences, because nothing moves a block from one side of a table
 * to the other. Aligned as one sequence they pair across it, and the plan that
 * follows can neither move the base paragraph past the nested table nor delete
 * the one it left behind when that paragraph ends the cell. So the runs are
 * aligned separately, the way rows and cells already are.
 *
 * Document order is the whole input: a cell's blocks are contiguous until a
 * block of another cell interrupts them, and inside one table segment the only
 * cell that can interrupt another is one of a table nested in it.
 */
const cellRunOrdinals = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
): ReadonlyMap<Block, number> => {
  const ordinals = new Map<Block, number>();
  const runByCell = new Map<string, number>();
  let previousKey: string | null = null;
  for (const block of blocks) {
    const { table } = block;
    if (!table) {
      continue;
    }
    const key = `${String(table.tableIndex)}:${String(table.rowIndex)}:${String(table.cellIndex)}`;
    const seen = runByCell.get(key) ?? -1;
    // Leaving the cell and coming back is what a nested table between two of
    // its paragraphs looks like in document order, and it starts a new run.
    const run = key === previousKey ? seen : seen + 1;
    runByCell.set(key, run);
    ordinals.set(block, run);
    previousKey = key;
  }
  return ordinals;
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

/**
 * The heuristic halves of a profile, or the fact that a cap stopped them.
 *
 * A cap is a statement about cost, never about content. Reading a cap as "no
 * anchors, no tokens" reads "unknown" as "unequal", which is how a table one
 * block past the cap came to compare unequal with itself.
 *
 * Anchors are positive evidence and never a completeness claim: a `computed`
 * set may still stop early when the token budget runs out, so a text missing
 * from it means nothing.
 */
type ContentStructureAnchors =
  | { status: "computed"; texts: readonly string[] }
  | { status: "skipped-over-cap" };

type ContentStructureTokens =
  | { status: "computed"; counts: ReadonlyMap<string, number>; total: number }
  | { status: "skipped-over-cap" };

type ContentStructureProfile = {
  anchors: ContentStructureAnchors;
  blockIds: readonly string[];
  containerIdentity: ContentContainerIdentityProfile | null;
  /**
   * Content identity for the whole sequence, at every size. The structure used
   * to be retained verbatim, which a sequence past a cap could not afford, so
   * exactness was the first thing a cap took away. A digest folds the same
   * structure into constant memory, so it survives every cap and identical
   * input pairs before any heuristic is consulted.
   */
  contentDigest: string;
  physicalCellCount: number | null;
  stableIds: readonly string[];
  tokens: ContentStructureTokens;
};

/** Four independent 32-bit lanes, so a digest collision needs four at once. */
const CONTENT_DIGEST_SEEDS = [0x81_1c_9d_c5, 0x9e_37_79_b9, 0x85_eb_ca_6b, 0xc2_b2_ae_35] as const;
const CONTENT_DIGEST_PRIMES = [0x01_00_01_93, 0x5b_d1_e9_95, 0x27_d4_eb_2f, 0x16_56_67_b1] as const;

const foldContentDigest = (lanes: Int32Array, text: string): void => {
  for (let index = 0; index < text.length; index++) {
    const unit = text.charCodeAt(index);
    for (let lane = 0; lane < lanes.length; lane++) {
      lanes[lane] = Math.imul((lanes[lane] ?? 0) ^ unit, CONTENT_DIGEST_PRIMES[lane] ?? 1);
    }
  }
  // Length-terminate, so concatenation cannot forge a different split.
  for (let lane = 0; lane < lanes.length; lane++) {
    lanes[lane] = Math.imul((lanes[lane] ?? 0) ^ text.length, CONTENT_DIGEST_PRIMES[lane] ?? 1);
  }
};

const renderContentDigest = (lanes: Int32Array, blockCount: number): string =>
  `${String(blockCount)}:${[...lanes].map((lane) => (lane >>> 0).toString(36)).join(":")}`;

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

  // The digest is a fact about the blocks and costs one pass with no retention,
  // so it is taken before any cap is consulted.
  const lanes = Int32Array.from(CONTENT_DIGEST_SEEDS);
  for (const block of blocks) {
    foldContentDigest(lanes, JSON.stringify([...blockStructure(block), block.kind, block.text]));
  }
  const identity = {
    blockIds,
    containerIdentity,
    contentDigest: renderContentDigest(lanes, blocks.length),
    physicalCellCount: physicalCellCount ?? null,
    stableIds,
  };
  const overCap = {
    ...identity,
    anchors: { status: "skipped-over-cap" },
    tokens: { status: "skipped-over-cap" },
  } as const;

  if (blocks.length > MAX_CONTENT_STRUCTURE_PROFILE_BLOCKS) {
    return overCap;
  }

  let textCodeUnits = 0;
  for (const block of blocks) {
    if (
      block.kind.length > MAX_CONTENT_STRUCTURE_PROFILE_KIND_CODE_UNITS ||
      block.text.length > MAX_CONTENT_STRUCTURE_PROFILE_TEXT_CODE_UNITS - textCodeUnits
    ) {
      return overCap;
    }
    textCodeUnits += block.text.length;
  }

  const tokenCounts = new Map<string, number>();
  const anchorTexts = new Set<string>();
  let tokenCount = 0;
  let tokensComplete = true;
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
    ...identity,
    anchors: { status: "computed", texts: [...anchorTexts] },
    tokens: tokensComplete
      ? { status: "computed", counts: tokenCounts, total: tokenCount }
      : { status: "skipped-over-cap" },
  };
};

/** How alike two sequences read, or that nothing was retained to judge it by. */
type ContentStructureSimilarity = { status: "measured"; value: number } | { status: "unknown" };

const UNKNOWN_CONTENT_STRUCTURE_SIMILARITY = { status: "unknown" } as const;

const contentStructureProfileSimilarity = (
  base: ContentStructureProfile,
  revised: ContentStructureProfile,
  workSession: FolioContentAlignmentWorkSession,
): ContentStructureSimilarity => {
  if (base.tokens.status === "skipped-over-cap" || revised.tokens.status === "skipped-over-cap") {
    return UNKNOWN_CONTENT_STRUCTURE_SIMILARITY;
  }
  if (base.tokens.total === 0 || revised.tokens.total === 0) {
    // No words on one side: the measure is undefined, not zero.
    return UNKNOWN_CONTENT_STRUCTURE_SIMILARITY;
  }
  const [tokens, counterparts] =
    base.tokens.counts.size <= revised.tokens.counts.size
      ? [base.tokens.counts, revised.tokens.counts]
      : [revised.tokens.counts, base.tokens.counts];
  if (tokens.size > workSession.remainingStructuralTokenLookups) {
    return UNKNOWN_CONTENT_STRUCTURE_SIMILARITY;
  }
  workSession.remainingStructuralTokenLookups -= tokens.size;
  let shared = 0;
  for (const [token, count] of tokens) {
    shared += Math.min(count, counterparts.get(token) ?? 0);
  }
  return { status: "measured", value: (2 * shared) / (base.tokens.total + revised.tokens.total) };
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

/**
 * A container's blocks as the paragraph identities the save reasons about: a
 * stable id is one the package wrote, a positional id one folio minted for a
 * paragraph that arrived without one.
 */
const containerParagraphIdentities = (
  profile: Extract<ContentContainerIdentityProfile, { status: "available" }>,
): ParagraphIdentity[] =>
  profile.blockIds.map((paraId, ordinal) =>
    profile.idStabilities[ordinal] === "stable"
      ? { type: "authored", paraId, ordinal }
      : { type: "minted", paraId, ordinal },
  );

/**
 * Whether two containers are the same container, read from the paragraph ids
 * their blocks carry.
 *
 * This used to ask for a positional-to-stable *transition* with identical ids
 * at every index, which is the shape a save produced when it stamped folio's
 * minted ids into the package it wrote. That made the comparison depend on a
 * side effect of the save rather than on the documents in front of it: once a
 * save stopped persisting a minted id, both sides read positional, the
 * transition never fired, and an edited row was reported deleted and inserted
 * instead of edited.
 *
 * The question is the one {@link alignParagraphOrdinals} answers, and it is
 * asked through the owner the save asks it through, so the two cannot come to
 * disagree about which paragraph is which. A comparison sees two different
 * documents, so absence of contradiction is not enough on its own: most of the
 * container has to actively confirm the ordinals, or every container of equal
 * size would pair with every other on the strength of saying nothing. A
 * positional id is derived from the paragraph's own text and position, so an
 * unedited block keeps it across the edit and supplies that confirmation, while
 * an edited block's id changes with its text. A majority is therefore "the
 * blocks this edit did not touch", and it is what separates a row with an edit
 * in it from a row replaced by a different row, which shares no id at all.
 *
 * This signal only has to carry the containers similarity cannot reach: a pair
 * it declines is still weighed on content, so a two-block container with one
 * block edited pairs on the half it kept rather than on a bare majority.
 */
const contentContainersShareIdentity = (
  base: ContentContainerIdentityProfile | null,
  revised: ContentContainerIdentityProfile | null,
): boolean => {
  if (
    base === null ||
    revised === null ||
    base.status !== "available" ||
    revised.status !== "available" ||
    base.blockIds.length === 0
  ) {
    return false;
  }
  const alignment = alignParagraphOrdinals(
    containerParagraphIdentities(base),
    containerParagraphIdentities(revised),
  );
  return alignment.consistent && alignment.sharedOrdinals * 2 > base.blockIds.length;
};

/**
 * Whether two containers carry the same block ids in the same order, over
 * their whole length.
 *
 * Stricter than {@link contentContainersShareIdentity}, which answers which
 * container this is while blocks come and go. This answers whether there is
 * anything left to decide: nothing moved within either container, so paired
 * with equal content they are one container seen twice. An identity a cap
 * withheld is unknown, never equal.
 */
const contentContainerIdsAreEqual = (
  base: ContentContainerIdentityProfile | null,
  revised: ContentContainerIdentityProfile | null,
): boolean =>
  base !== null &&
  revised !== null &&
  base.status === "available" &&
  revised.status === "available" &&
  base.blockIds.length > 0 &&
  base.blockIds.length === revised.blockIds.length &&
  base.blockIds.every((id, index) => id === revised.blockIds[index]);

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
      (left, right) => left.baseIndex - right.baseIndex || left.revisedIndex - right.revisedIndex,
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
  // A container whose paragraph ids line the two sides up is evidence in its own
  // right. Text and stable-id anchors constrain it when present; requiring an
  // anchor first loses every surviving row when each one also contains an edit.
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
      contentContainersShareIdentity(
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
  soleShiftedPairResiduePolicy?: "conservative" | "row-count-evidence";
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
  soleShiftedPairResiduePolicy = "conservative",
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
    let hasProvenanceTransition = contentContainersShareIdentity(
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
  const internContentDigest = (digest: string): number => {
    const existing = exactSignatureKeys.get(digest);
    if (existing !== undefined) {
      return existing;
    }
    const key = nextExactSignatureKey++;
    exactSignatureKeys.set(digest, key);
    return key;
  };
  const baseExactSignatureKeys = base.map(({ profile }) =>
    internContentDigest(profile.contentDigest),
  );
  const revisedExactSignatureKeys = revised.map(({ profile }) =>
    internContentDigest(profile.contentDigest),
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
    const canPairIndexes = ({ baseIndex, revisedIndex }: FolioContentBlockPair): boolean => {
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
      const exact = baseExactSignatureKeys[baseIndex] === revisedExactSignatureKeys[revisedIndex];
      const stable = stablePairs.has(pairIndex);
      const profileSimilarity: ContentStructureSimilarity = exact
        ? { status: "measured", value: 1 }
        : contentStructureProfileSimilarity(baseItem.profile, revisedItem.profile, workSession);
      // A cap that hid the tokens leaves similarity unknown, which asserts
      // nothing either way: it contributes no evidence and refutes none.
      const similar =
        profileSimilarity.status === "measured"
          ? Math.max(
              0,
              Math.min(1, profileSimilarity.value * similarityFactor(baseItem, revisedItem)),
            )
          : 0;
      const similarEnough =
        profileSimilarity.status === "measured" && similar >= CONTENT_STRUCTURE_PAIR_SIMILARITY;
      const stableAtSamePosition = stable && baseIndex === revisedIndex;
      const persistedAtSamePosition =
        baseIndex === revisedIndex && provenanceTransitionAtSamePosition[baseIndex] === 1;
      const shiftedPersisted = baseIndex !== revisedIndex && persistedPairs.has(pairIndex);
      const persisted = persistedAtSamePosition || shiftedPersisted;
      const soleStructuralSlot =
        pairSoleStructuralSlot && base.length === 1 && revised.length === 1;
      if (!exact && !stableAtSamePosition && !persisted && !soleStructuralSlot && !similarEnough) {
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
  let solePairIsIdentity = false;
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
      solePairIsIdentity =
        baseExactSignatureKeys[baseIndex] === revisedExactSignatureKeys[revisedIndex] &&
        contentContainerIdsAreEqual(
          base[baseIndex]?.profile.containerIdentity ?? null,
          revised[revisedIndex]?.profile.containerIdentity ?? null,
        );
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
  if (
    !solePairIsIdentity &&
    (soleShiftedPairResiduePolicy === "conservative" || base.length === revised.length) &&
    pairCount === 1 &&
    solePairIsShifted &&
    hasBaseOnly &&
    hasRevisedOnly
  ) {
    // One content MATCH cannot establish a shifted container mapping when it
    // also strands containers on both sides: that shape reads equally as
    // content moving between a deletion and an insertion. A changed container
    // count is one exception, since it supports retaining a surviving row
    // between a row insertion and a deletion.
    //
    // Identity is the other, and it is not a match at all. A pair whose
    // content digests are equal AND whose blocks carry the same ids in the
    // same order is the same container, so there is no mapping left to infer.
    // Both halves are needed: equal digests alone are two containers that read
    // alike, which a table of boilerplate rows has many of, and equal ids
    // alone are ids that agree after a reorder without the content backing it,
    // which is exactly what a positional id minted from text and position does
    // when rows move. Only a similarity-scored pair is split here.
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
  cellRuns: ReadonlyMap<Block, number>;
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
  cellRuns,
}: AlignRowCellsOptions<Block>): FolioContentAlignmentStep<Block>[] => {
  const byCell = (
    row: readonly Block[],
    columnKeys: ReadonlyMap<number, number> | undefined,
  ): Map<number, Map<number, Block[]>> => {
    const cells = new Map<number, Map<number, Block[]>>();
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
      const runs = cells.get(cellIndex) ?? new Map<number, Block[]>();
      cells.set(cellIndex, runs);
      const run = cellRuns.get(block) ?? 0;
      const blocks = runs.get(run);
      if (blocks) {
        blocks.push(block);
      } else {
        runs.set(run, [block]);
      }
    }
    return cells;
  };

  const ascending = (left: number, right: number): number => left - right;
  const baseCells = byCell(baseRow, baseColumnKeys);
  const revisedCells = byCell(revisedRow, revisedColumnKeys);
  const cellIndexes = [...new Set([...baseCells.keys(), ...revisedCells.keys()])].toSorted(
    ascending,
  );
  const steps: FolioContentAlignmentStep<Block>[] = [];
  for (const cellIndex of cellIndexes) {
    const baseRuns = baseCells.get(cellIndex) ?? new Map<number, Block[]>();
    const revisedRuns = revisedCells.get(cellIndex) ?? new Map<number, Block[]>();
    const runOrdinals = [...new Set([...baseRuns.keys(), ...revisedRuns.keys()])].toSorted(
      ascending,
    );
    for (const run of runOrdinals) {
      steps.push(
        ...alignCellRun({
          baseBlocks: baseRuns.get(run) ?? [],
          revisedBlocks: revisedRuns.get(run) ?? [],
          workSession,
          moveScopeContext,
          stableIdMismatch,
          idStability,
        }),
      );
    }
  }
  return steps;
};

type AlignCellRunOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  workSession: FolioContentAlignmentWorkSession;
  moveScopeContext: MoveScopeContext;
  stableIdMismatch: "pair" | "separate";
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
};

const alignCellRun = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  workSession,
  moveScopeContext,
  stableIdMismatch,
  idStability,
}: AlignCellRunOptions<Block>): FolioContentAlignmentStep<Block>[] => {
  const aligned = alignFolioContentBlocksInScope(baseBlocks, revisedBlocks, {
    workSession,
    idStability,
    stableIdMismatch,
    structuralScope: "pairedTableCell",
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
  return scopedAlignmentSteps(aligned, moveScopeContext, bucketForBlock);
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
    soleShiftedPairResiduePolicy: "row-count-evidence",
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
  cellRuns: ReadonlyMap<Block, number>;
};

const alignTableRows = <Block extends FolioContentBlock>({
  rows,
  workSession,
  moveScopeContext,
  stableIdMismatch,
  idStability,
  baseColumnKeys,
  revisedColumnKeys,
  cellRuns,
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
            cellRuns,
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
  cellRuns: ReadonlyMap<Block, number>;
};

const buildTablePlan = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  workSession,
  moveScopeContext,
  stableIdMismatch,
  idStability,
  cellRuns,
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
        cellRuns,
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
  // Document order is lost once the segment is grouped by table, so the runs
  // are read off the segment as it arrives.
  const cellRuns = new Map([...cellRunOrdinals(baseBlocks), ...cellRunOrdinals(revisedBlocks)]);
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
          cellRuns,
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
      let key = exactKeys.get(profile.contentDigest);
      if (key === undefined) {
        key = nextExactKey++;
        exactKeys.set(profile.contentDigest, key);
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
      for (const text of profile.anchors.status === "computed" ? profile.anchors.texts : []) {
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

/**
 * Whether the evidence lies outside the range, and only outside it.
 *
 * Evidence inside the range is at least as good a counterpart as evidence
 * beyond it, so a match that is also present locally does not make the local
 * pair suspect. Without the second test, two structurally identical tables in
 * one document cancel each other: each shares every exact signature and anchor
 * text with the other by construction, so each looks like it belongs
 * elsewhere, and comparing the document with itself reports both deleted and
 * both re-inserted.
 */
const sortedIndexesLeaveRange = (
  indexes: readonly number[] | undefined,
  start: number,
  end: number,
): boolean => {
  if (indexes === undefined) {
    return false;
  }
  if ((indexes.at(0) ?? start) >= start && (indexes.at(-1) ?? end - 1) < end) {
    return false;
  }
  return !indexes.some((index) => index >= start && index < end);
};

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
    profile.anchors.status === "computed" &&
    profile.anchors.texts.some((text) =>
      sortedIndexesLeaveRange(opposite.anchorTextIndexes.get(text), start, end),
    )
  ) {
    return true;
  }
  return profile.stableIds.some((id) =>
    sortedIndexesLeaveRange(opposite.stableIndexes.get(id), start, end),
  );
};

/**
 * Whether two sequences share a text anchor, or whether a cap left it unknown.
 *
 * The caller pairs a sole structural slot on "shares" alone. "unknown" is not
 * a match either: a profile a cap truncated is no evidence in either
 * direction, which is why identity rests on the digest rather than on this.
 */
type ContentAnchorOverlap = "shares" | "disjoint" | "unknown";

const profilesShareContentAnchor = (
  base: ContentStructureProfile,
  revised: ContentStructureProfile,
): ContentAnchorOverlap => {
  if (base.anchors.status === "skipped-over-cap" || revised.anchors.status === "skipped-over-cap") {
    return "unknown";
  }
  const baseTexts = new Set(base.anchors.texts);
  return revised.anchors.texts.some((text) => baseTexts.has(text)) ? "shares" : "disjoint";
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
    const solePairAnchorOverlap: ContentAnchorOverlap =
      baseItems.length === 1 && revisedItems.length === 1
        ? profilesShareContentAnchor(
            baseItems.at(0)?.profile ?? panic("A sole base table has no profile"),
            revisedItems.at(0)?.profile ?? panic("A sole revised table has no profile"),
          )
        : "disjoint";
    const alignments = alignProfiledContentSequence({
      base: baseItems,
      revised: revisedItems,
      workSession,
      canPair: (baseItem, revisedItem) =>
        !baseHasExternalEvidence.has(baseItem.item) &&
        !revisedHasExternalEvidence.has(revisedItem.item),
      pairSoleStructuralSlot: trustedBodyPairCount > 0 || solePairAnchorOverlap === "shares",
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
