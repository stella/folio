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
