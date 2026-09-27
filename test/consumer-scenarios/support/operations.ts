/**
 * Operations a model would write against what it just read: well-formed ones
 * built from the current blocks, and the mistakes models make (stale or
 * foreign ids, wrong hashes, malformed arguments). Operations carry no `id`,
 * as `suggest_changes` takes them; `withIds` adds ids for a core batch.
 */

import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  hashFolioAIBlockText,
  isFolioDocumentOperationModeSupported,
} from "@stll/folio-core/server";

import { type Random, sentence } from "./random.ts";

export type Mode = "direct" | "tracked-changes" | "suggested";
export const MODES: readonly Mode[] = ["direct", "tracked-changes", "suggested"];

export type Block = {
  id: string;
  kind: string;
  text: string;
  listReference?: { numId: number; level: number };
  table?: unknown;
};

export type Operation = { type: string } & Record<string, unknown>;

const inTable = (block: Block): boolean => block.table !== undefined;

const wordsOf = (text: string): { word: string; start: number }[] =>
  [...text.matchAll(/[\p{L}\p{N}]{3,}/gu)].map((match) => ({
    word: match[0],
    start: match.index,
  }));

type Generator = (blocks: readonly Block[], random: Random) => Operation | null;

const withText = (blocks: readonly Block[]) => blocks.filter((block) => wordsOf(block.text).length);

const numberingRefs = (blocks: readonly Block[]) =>
  blocks.flatMap((block) => (block.listReference ? [block.listReference] : []));

const range = (block: Block, random: Random) => {
  const words = wordsOf(block.text);
  if (words.length === 0) return null;
  const { word, start } = random.pick(words);
  return createFolioAITextRangeHandle({
    blockId: block.id,
    text: block.text,
    startOffset: start,
    endOffset: start + word.length,
  });
};

const paragraphProperties = (blocks: readonly Block[], random: Random) => {
  const refs = numberingRefs(blocks);
  return random.pick([
    { styleId: "Heading2" },
    { styleId: null },
    { alignment: "center" },
    { alignment: null },
    { spacing: { spaceBefore: 120, spaceAfter: 120 } },
    { numbering: { start: "new", kind: random.pick(["numbered", "bullet"]) } },
    ...(refs.length > 0 ? [{ numbering: random.pick(refs) }, { numbering: null }] : []),
  ]);
};

/** Well-formed operations, by type. A generator returns null when nothing fits. */
export const GENERATORS: Record<string, Generator> = {
  replaceInBlock: (blocks, random) => {
    const candidates = withText(blocks);
    if (candidates.length === 0) return null;
    const block = random.pick(candidates);
    return {
      type: "replaceInBlock",
      blockId: block.id,
      find: random.pick(wordsOf(block.text)).word,
      replace: random.pick(["revised", "amended", "the Customer", ""]),
      precondition: { blockTextHash: hashFolioAIBlockText(block.text) },
    };
  },
  replaceRange: (blocks, random) => {
    const candidates = withText(blocks);
    if (candidates.length === 0) return null;
    const handle = range(random.pick(candidates), random);
    return handle && { type: "replaceRange", range: handle, replace: "updated" };
  },
  formatRange: (blocks, random) => {
    const candidates = withText(blocks);
    if (candidates.length === 0) return null;
    const handle = range(random.pick(candidates), random);
    return (
      handle && {
        type: "formatRange",
        range: handle,
        formatting: random.pick([{ bold: true }, { italic: true }, { underline: true }]),
      }
    );
  },
  commentOnRange: (blocks, random) => {
    const candidates = withText(blocks);
    if (candidates.length === 0) return null;
    const handle = range(random.pick(candidates), random);
    return handle && { type: "commentOnRange", range: handle, comment: { text: sentence(random) } };
  },
  insertAfterBlock: (blocks, random) => {
    if (blocks.length === 0) return null;
    const refs = numberingRefs(blocks);
    const extra = random.pick([
      {},
      { styleId: "Heading2" },
      { alignment: "right" },
      { numbering: { start: "new", kind: random.pick(["numbered", "bullet"]) } },
      ...(refs.length > 0 ? [{ numbering: random.pick(refs) }] : []),
    ]);
    return {
      type: "insertAfterBlock",
      blockId: random.pick(blocks).id,
      text: sentence(random),
      ...extra,
    };
  },
  insertBeforeBlock: (blocks, random) =>
    blocks.length === 0
      ? null
      : { type: "insertBeforeBlock", blockId: random.pick(blocks).id, text: sentence(random) },
  replaceBlock: (blocks, random) => {
    const candidates = blocks.filter((block) => !inTable(block));
    if (candidates.length === 0) return null;
    return { type: "replaceBlock", blockId: random.pick(candidates).id, text: sentence(random) };
  },
  deleteBlock: (blocks, random) => {
    const candidates = blocks.filter((block) => !inTable(block));
    if (candidates.length < 3) return null;
    return { type: "deleteBlock", blockId: random.pick(candidates).id };
  },
  splitBlock: (blocks, random) => {
    const candidates = blocks.filter((block) => !inTable(block) && wordsOf(block.text).length > 1);
    if (candidates.length === 0) return null;
    const block = random.pick(candidates);
    const words = wordsOf(block.text);
    return { type: "splitBlock", blockId: block.id, offset: random.pick(words.slice(1)).start };
  },
  mergeBlockWithNext: (blocks, random) => {
    const candidates = blocks.filter(
      (block, index) =>
        !inTable(block) && blocks[index + 1] !== undefined && !inTable(blocks[index + 1] as Block),
    );
    if (candidates.length === 0) return null;
    return { type: "mergeBlockWithNext", blockId: random.pick(candidates).id, separator: " " };
  },
  setBlockParagraphProperties: (blocks, random) => {
    const candidates = blocks.filter((block) => !inTable(block));
    if (candidates.length === 0) return null;
    return {
      type: "setBlockParagraphProperties",
      blockId: random.pick(candidates).id,
      properties: paragraphProperties(blocks, random),
    };
  },
  commentOnBlock: (blocks, random) => {
    const candidates = withText(blocks);
    if (candidates.length === 0) return null;
    const block = random.pick(candidates);
    return {
      type: "commentOnBlock",
      blockId: block.id,
      quote: random.pick(wordsOf(block.text)).word,
      comment: { text: sentence(random) },
    };
  },
  insertTable: (blocks, random) => {
    const candidates = blocks.filter((block) => !inTable(block));
    if (candidates.length === 0) return null;
    return {
      type: "insertTable",
      blockId: random.pick(candidates).id,
      rows: [
        ["Term", "Value"],
        ["Period", "12 months"],
      ],
    };
  },
  insertTableRow: (blocks, random) => {
    const cells = blocks.filter(inTable);
    return cells.length === 0
      ? null
      : { type: "insertTableRow", blockId: random.pick(cells).id, position: "after" };
  },
  deleteTableRow: (blocks, random) => {
    const cells = blocks.filter(inTable);
    return cells.length === 0 ? null : { type: "deleteTableRow", blockId: random.pick(cells).id };
  },
  insertTableColumn: (blocks, random) => {
    const cells = blocks.filter(inTable);
    return cells.length === 0
      ? null
      : { type: "insertTableColumn", blockId: random.pick(cells).id, position: "after" };
  },
  deleteTable: (blocks, random) => {
    const cells = blocks.filter(inTable);
    return cells.length === 0 ? null : { type: "deleteTable", blockId: random.pick(cells).id };
  },
};

/** Whether `suggest_changes` / a core batch accepts `type` in `mode`. */
export const supports = (type: string, mode: Mode): boolean =>
  isFolioDocumentOperationModeSupported(type as never, mode);

/** A random well-formed operation `mode` supports, or null. */
export const randomOperation = (
  blocks: readonly Block[],
  mode: Mode,
  random: Random,
  types: readonly string[] = Object.keys(GENERATORS),
): Operation | null => {
  const usable = types.filter((type) => supports(type, mode));
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const generator = GENERATORS[random.pick(usable)];
    const operation = generator?.(blocks, random);
    if (operation) return operation;
  }
  return null;
};

/** A block id this document never had, shaped like a real one. */
export const foreignBlockId = (random: Random): string =>
  Array.from({ length: 8 }, () => random.pick([..."0123456789ABCDEF"])).join("");

/**
 * The mistakes a model makes, as arguments to `suggest_changes`. Each must be
 * refused (with an issue) or leave a document that saves. `staleIds` are ids
 * the model read earlier that no longer name a block.
 */
export const MISTAKES: Record<
  string,
  (blocks: readonly Block[], random: Random, staleIds: readonly string[]) => unknown
> = {
  staleBlockId: (_blocks, random, staleIds) => ({
    operations: [
      {
        type: "replaceBlock",
        blockId: staleIds.length > 0 ? random.pick(staleIds) : foreignBlockId(random),
        text: sentence(random),
      },
    ],
  }),
  foreignBlockId: (_blocks, random) => ({
    operations: [{ type: "insertAfterBlock", blockId: foreignBlockId(random), text: "x" }],
  }),
  wrongPrecondition: (blocks, random) =>
    blocks.length === 0
      ? { operations: [] }
      : {
          operations: [
            {
              type: "replaceBlock",
              blockId: random.pick(blocks).id,
              text: sentence(random),
              precondition: { blockTextHash: "hdeadbeef" },
            },
          ],
        },
  staleRange: (blocks, random) => {
    const block = withText(blocks)[0];
    const handle = block && range(block, random);
    return {
      operations: [
        {
          type: "replaceRange",
          range: handle ? { ...handle, selectedTextHash: "hstale00" } : { type: "textRange" },
          replace: "x",
        },
      ],
    };
  },
  findNotInBlock: (blocks, random) => ({
    operations: [
      {
        type: "replaceInBlock",
        blockId: blocks.length > 0 ? random.pick(blocks).id : "00000000",
        find: "zzz-not-there",
        replace: "x",
      },
    ],
  }),
  splitOutOfRange: (blocks, random) => ({
    operations: [
      {
        type: "splitBlock",
        blockId: blocks.length > 0 ? random.pick(blocks).id : "00000000",
        offset: random.pick([-1, 10_000, 0.5]),
      },
    ],
  }),
  numIdAsString: (blocks, random) => ({
    operations: [
      {
        type: "insertAfterBlock",
        blockId: blocks.length > 0 ? random.pick(blocks).id : "00000000",
        text: "x",
        numbering: { numId: "1", level: "0" },
      },
    ],
  }),
  unknownType: () => ({ operations: [{ type: "rewriteEverything", text: "x" }] }),
  missingFields: (blocks, random) => ({
    operations: [{ type: random.pick(["replaceBlock", "insertAfterBlock", "deleteBlock"]) }],
  }),
  unknownKey: (blocks, random) => ({
    operations: [
      {
        type: "replaceBlock",
        blockId: blocks.length > 0 ? random.pick(blocks).id : "00000000",
        text: "x",
        colour: "red",
      },
    ],
  }),
  notAnArray: () => ({ operations: "replace everything" }),
  emptyBatch: () => ({ operations: [] }),
  tooMany: (blocks) => ({
    operations: Array.from({ length: 51 }, () => ({
      type: "insertAfterBlock",
      blockId: blocks[0]?.id ?? "00000000",
      text: "x",
    })),
  }),
  nullArguments: () => null,
  styleThatDoesNotExist: (blocks, random) => ({
    operations: [
      {
        type: "insertAfterBlock",
        blockId: blocks.length > 0 ? random.pick(blocks).id : "00000000",
        text: "x",
        styleId: "NoSuchStyle",
      },
    ],
  }),
  levelBeyondDefinition: (blocks, random) => {
    const refs = numberingRefs(blocks);
    return {
      operations: [
        {
          type: "insertAfterBlock",
          blockId: blocks.length > 0 ? random.pick(blocks).id : "00000000",
          text: "x",
          ...(refs.length > 0 ? { numbering: { numId: random.pick(refs).numId, level: 8 } } : {}),
        },
      ],
    };
  },
};

/**
 * Mistakes the package can recognise before applying anything, by the issue
 * code `suggest_changes` must refuse each with. Accepting one would save a
 * document that silently ignores what the model asked for.
 */
export const REFUSED_MISTAKES: Readonly<Partial<Record<keyof typeof MISTAKES, string>>> = {
  styleThatDoesNotExist: "missingStyle",
};

/** A numbering reference to an instance the package does not define (#1103). */
export const undefinedNumbering = (blocks: readonly Block[]) => {
  const used = new Set(numberingRefs(blocks).map((ref) => ref.numId));
  let numId = 1;
  while (used.has(numId)) numId += 1;
  return { numId: numId + 40, level: 0 };
};

/** Give each operation an id and wrap them as a core batch. */
export const coreBatch = (operations: readonly Operation[], mode: Mode) => ({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  mode,
  operations: operations.map((operation, index) => ({ id: `op-${index + 1}`, ...operation })),
});

// ---------------------------------------------------------------------------
// Collisions
// ---------------------------------------------------------------------------

type Collision = (blocks: readonly Block[]) => Operation[] | null;

const prose = (blocks: readonly Block[]) =>
  blocks.filter((block) => !inTable(block) && wordsOf(block.text).length > 1);

/** The first prose block with words whose next block is prose with words too. */
const withNext = (blocks: readonly Block[]): [Block, Block] | null => {
  for (let index = 0; index + 1 < blocks.length; index += 1) {
    const block = blocks[index] as Block;
    const next = blocks[index + 1] as Block;
    if (
      !inTable(block) &&
      !inTable(next) &&
      wordsOf(block.text).length > 1 &&
      wordsOf(next.text).length > 0
    ) {
      return [block, next];
    }
  }
  return null;
};

const handleOf = (block: Block, start: number, end: number) =>
  createFolioAITextRangeHandle({
    blockId: block.id,
    text: block.text,
    startOffset: start,
    endOffset: end,
  });

const wordRange = (block: Block, which: number) => {
  const word = wordsOf(block.text).at(which);
  return word ? handleOf(block, word.start, word.start + word.word.length) : null;
};

const replaceWord = (block: Block, which = 0): Operation => ({
  type: "replaceInBlock",
  blockId: block.id,
  find: wordsOf(block.text).at(which)?.word ?? "",
  replace: "revised",
});

const ZWJ = String.fromCodePoint(0x200d);
const ASTRAL = /[\u{10000}-\u{10FFFF}]/u;

/** The user-perceived characters of `text` that hold an astral code point, as UTF-16 ranges. */
const emojiOf = (text: string): { start: number; end: number }[] =>
  [...new Intl.Segmenter("en", { granularity: "grapheme" }).segment(text)]
    .filter(({ segment }) => ASTRAL.test(segment))
    .map(({ segment, index }) => ({ start: index, end: index + segment.length }));

const emojiBlocks = (blocks: readonly Block[]) =>
  blocks.filter(
    (block) => !inTable(block) && emojiOf(block.text).length > 0 && wordsOf(block.text).length > 0,
  );

type Location = { tableIndex: number; rowIndex: number; cellIndex: number; rowSpan: number };
const locationOf = (block: Block): Location | undefined => block.table as Location | undefined;

/** The paragraphs of the anchor's table row. */
const rowCells = (blocks: readonly Block[], anchor: Block) =>
  blocks.filter(
    (block) =>
      locationOf(block)?.tableIndex === locationOf(anchor)?.tableIndex &&
      locationOf(block)?.rowIndex === locationOf(anchor)?.rowIndex,
  );

const cellCount = (blocks: readonly Block[]) =>
  new Set(blocks.map((block) => locationOf(block)?.cellIndex)).size;

const rowCount = (blocks: readonly Block[], anchor: Block) =>
  new Set(
    blocks
      .filter((block) => locationOf(block)?.tableIndex === locationOf(anchor)?.tableIndex)
      .map((block) => locationOf(block)?.rowIndex),
  ).size;

/** The last table cell of the document: a new row or column there follows the others. */
const lastCell = (blocks: readonly Block[]) => blocks.findLast(inTable);

const cellTexts = (count: number, label: string) =>
  Array.from({ length: count }, (_, index) => `${label} ${index + 1}`);

/**
 * Batches aimed where batch claims, offsets and payloads meet: two
 * operations on one block, an edit inside a block another operation deletes,
 * ranges next to surrogate pairs and inside a ZWJ sequence, and table
 * payloads that must all land. Deterministic: each takes the first block
 * that fits, or returns null.
 */
export const COLLISIONS: Record<string, Collision> = {
  replaceThenDeleteSameBlock: (blocks) => {
    const block = prose(blocks)[0];
    return block ? [replaceWord(block), { type: "deleteBlock", blockId: block.id }] : null;
  },
  deleteThenReplaceSameBlock: (blocks) => {
    const block = prose(blocks)[0];
    return block ? [{ type: "deleteBlock", blockId: block.id }, replaceWord(block)] : null;
  },
  replaceThenDeleteNext: (blocks) => {
    const pair = withNext(blocks);
    return pair ? [replaceWord(pair[0]), { type: "deleteBlock", blockId: pair[1].id }] : null;
  },
  deleteThenReplaceNext: (blocks) => {
    const pair = withNext(blocks);
    return pair ? [{ type: "deleteBlock", blockId: pair[0].id }, replaceWord(pair[1])] : null;
  },
  twoRangesInOneBlock: (blocks) => {
    const block = prose(blocks)[0];
    const first = block && wordRange(block, 0);
    const last = block && wordRange(block, -1);
    return first && last
      ? [
          { type: "replaceRange", range: last, replace: "second" },
          { type: "replaceRange", range: first, replace: "first" },
        ]
      : null;
  },
  formatThenReplaceSameRange: (blocks) => {
    const block = prose(blocks)[0];
    const handle = block && wordRange(block, 0);
    return handle
      ? [
          { type: "formatRange", range: handle, formatting: { bold: true } },
          { type: "replaceRange", range: handle, replace: "changed" },
        ]
      : null;
  },
  commentAndFormatSameRange: (blocks) => {
    const block = prose(blocks)[0];
    const handle = block && wordRange(block, -1);
    return handle
      ? [
          { type: "commentOnRange", range: handle, comment: { text: "Check this word." } },
          { type: "formatRange", range: handle, formatting: { italic: true } },
        ]
      : null;
  },
  splitThenEditSecondHalf: (blocks) => {
    const block = prose(blocks).find((candidate) => wordsOf(candidate.text).length > 2);
    const second = block && wordsOf(block.text)[1];
    const last = block && wordRange(block, -1);
    return block && second && last
      ? [
          { type: "splitBlock", blockId: block.id, offset: second.start },
          { type: "replaceRange", range: last, replace: "tail" },
        ]
      : null;
  },
  mergeThenEditNext: (blocks) => {
    const pair = withNext(blocks);
    return pair
      ? [{ type: "mergeBlockWithNext", blockId: pair[0].id, separator: " " }, replaceWord(pair[1])]
      : null;
  },
  replaceBlockThenInsertAfter: (blocks) => {
    const block = prose(blocks)[0];
    return block
      ? [
          { type: "replaceBlock", blockId: block.id, text: "A rewritten clause." },
          { type: "insertAfterBlock", blockId: block.id, text: "An added clause." },
        ]
      : null;
  },
  twoInsertsAfterOneBlock: (blocks) => {
    const block = prose(blocks)[0];
    return block
      ? [
          { type: "insertAfterBlock", blockId: block.id, text: "First added clause." },
          { type: "insertAfterBlock", blockId: block.id, text: "Second added clause." },
        ]
      : null;
  },
  insertAroundOneBlock: (blocks) => {
    const block = prose(blocks).at(-1);
    return block
      ? [
          { type: "insertAfterBlock", blockId: block.id, text: "Clause after." },
          { type: "insertBeforeBlock", blockId: block.id, text: "Clause before." },
        ]
      : null;
  },
  restyleThenEdit: (blocks) => {
    const block = prose(blocks).find((candidate) => candidate.kind === "paragraph");
    return block
      ? [
          {
            type: "setBlockParagraphProperties",
            blockId: block.id,
            properties: { styleId: "Heading2" },
          },
          replaceWord(block, -1),
        ]
      : null;
  },
  deleteTwoNeighbours: (blocks) => {
    const pair = withNext(blocks);
    return pair
      ? [
          { type: "deleteBlock", blockId: pair[1].id },
          { type: "deleteBlock", blockId: pair[0].id },
        ]
      : null;
  },
  insertStyledAndNumbered: (blocks) => {
    const block = prose(blocks)[0];
    return block
      ? [
          {
            type: "insertAfterBlock",
            blockId: block.id,
            text: "A second-level heading.",
            styleId: "Heading2",
          },
          {
            type: "insertAfterBlock",
            blockId: block.id,
            text: "A new bullet.",
            numbering: { start: "new", kind: "bullet" },
          },
        ]
      : null;
  },
  tableRowWithTexts: (blocks) => {
    const anchor = lastCell(blocks);
    return anchor
      ? [
          {
            type: "insertTableRow",
            blockId: anchor.id,
            position: "after",
            cellTexts: cellTexts(cellCount(rowCells(blocks, anchor)), "Row cell"),
          },
        ]
      : null;
  },
  tableRowBeforeWithTexts: (blocks) => {
    const anchor = lastCell(blocks);
    return anchor
      ? [
          {
            type: "insertTableRow",
            blockId: anchor.id,
            position: "before",
            cellTexts: cellTexts(cellCount(rowCells(blocks, anchor)), "Earlier cell"),
          },
        ]
      : null;
  },
  tableColumnWithTexts: (blocks) => {
    const anchor = lastCell(blocks);
    return anchor
      ? [
          {
            type: "insertTableColumn",
            blockId: anchor.id,
            position: "after",
            cellTexts: cellTexts(rowCount(blocks, anchor), "Column cell"),
          },
        ]
      : null;
  },
  tableRowThroughVerticalMerge: (blocks) => {
    // A row after one where a vertical merge starts: the merge grows through
    // the new row, which has one cell fewer; every text must still land.
    const anchor = blocks.find((block) => (locationOf(block)?.rowSpan ?? 1) > 1);
    if (!anchor) return null;
    const cells = rowCells(blocks, anchor);
    const merged = cells.filter((block) => (locationOf(block)?.rowSpan ?? 1) > 1).length;
    return [
      {
        type: "insertTableRow",
        blockId: anchor.id,
        position: "after",
        cellTexts: cellTexts(cellCount(cells) - merged, "Merged row cell"),
      },
    ];
  },
  deleteRowThenEditAnotherRow: (blocks) => {
    const cells = blocks.filter((block) => inTable(block) && wordsOf(block.text).length > 0);
    const first = cells[0];
    const inFirstRow = new Set(first ? rowCells(blocks, first).map(({ id }) => id) : []);
    const other = cells.find((cell) => !inFirstRow.has(cell.id));
    return first && other
      ? [{ type: "deleteTableRow", blockId: other.id }, replaceWord(first)]
      : null;
  },
  replaceBesideEmoji: (blocks) => {
    // In every block with an emoji: the word right before the first one, up
    // to it, and the word right after the last one, from it.
    const operations: Operation[] = [];
    for (const block of emojiBlocks(blocks)) {
      const emoji = emojiOf(block.text);
      const first = emoji[0] as { start: number };
      const last = emoji.at(-1) as { end: number };
      const before = wordsOf(block.text).findLast(
        (word) => word.start + word.word.length <= first.start,
      );
      const after = wordsOf(block.text).find((word) => word.start >= last.end);
      const left = before && handleOf(block, before.start, first.start);
      if (left) operations.push({ type: "replaceRange", range: left, replace: "Signed " });
      const right = after && handleOf(block, last.end, after.start + after.word.length);
      if (right) operations.push({ type: "replaceRange", range: right, replace: " endorsed" });
    }
    return operations.length > 0 ? operations : null;
  },
  replaceEmoji: (blocks) => {
    const operations: Operation[] = [];
    for (const block of emojiBlocks(blocks)) {
      const { start, end } = emojiOf(block.text)[0] as { start: number; end: number };
      const handle = handleOf(block, start, end);
      if (handle) operations.push({ type: "replaceRange", range: handle, replace: "(ok)" });
    }
    return operations.length > 0 ? operations : null;
  },
  formatAndCommentOverEmoji: (blocks) => {
    const operations: Operation[] = [];
    for (const block of emojiBlocks(blocks)) {
      const handle = handleOf(block, 0, block.text.length);
      if (!handle) continue;
      operations.push(
        { type: "formatRange", range: handle, formatting: { bold: true } },
        { type: "commentOnRange", range: handle, comment: { text: "Around the emoji." } },
      );
    }
    return operations.length > 0 ? operations : null;
  },
  rangeInsideSurrogatePair: (blocks) => {
    const block = emojiBlocks(blocks)[0];
    const at = block && emojiOf(block.text)[0]?.start;
    if (!block || at === undefined) return null;
    // Half an emoji: `createFolioAITextRangeHandle` will not build it; a model may.
    const end = Math.min(block.text.length, at + 4);
    const handle = {
      type: "textRange",
      story: "main",
      blockId: block.id,
      startOffset: at + 1,
      endOffset: end,
      selectedTextHash: hashFolioAIBlockText(block.text.slice(at + 1, end)),
    };
    return [{ type: "replaceRange", range: handle, replace: "x" }];
  },
  rangeInsideZwjSequence: (blocks) => {
    const block = blocks.find((candidate) => candidate.text.includes(ZWJ));
    const join = block ? block.text.indexOf(ZWJ) : -1;
    // The first person of a family: a whole code point, half a character.
    const handle = block && join >= 2 ? handleOf(block, join - 2, join) : null;
    return handle ? [{ type: "replaceRange", range: handle, replace: "\u{1F469}" }] : null;
  },
};
