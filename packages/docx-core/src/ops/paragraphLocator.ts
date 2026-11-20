import type { BlockContent, DocumentBody, Paragraph } from "../model/document";
import { blockListAt, type ParagraphLocation, storyParagraphs } from "./blocks";
import { idKey } from "./ids";
import { changedListBounds } from "./listSplice";
import { recentValues } from "./recentValues";

/**
 * Finding a story's paragraph by its id without locating every paragraph of
 * the story.
 *
 * Operations never change a block list in place: each returns a new list
 * that keeps the blocks it left. So the paragraphs a top-level block holds
 * are located once per block object, and which top-level blocks hold each id
 * is indexed once per block list, derived from a list indexed before it
 * along the blocks that differ between the two.
 */

/** Each id a top-level block holds, with where it sits inside the block, as if the block stood first. */
const ownLocations = new WeakMap<
  BlockContent,
  ReadonlyMap<string, readonly ParagraphLocation[]>
>();

const locationsIn = (
  block: BlockContent,
): ReadonlyMap<string, readonly ParagraphLocation[]> => {
  let known = ownLocations.get(block);
  if (known === undefined) {
    const byKey = new Map<string, ParagraphLocation[]>();
    for (const location of storyParagraphs({ content: [block] })) {
      const id = location.paragraph.paraId;
      if (id === undefined) continue;
      const key = idKey(id);
      const held = byKey.get(key);
      if (held === undefined) byKey.set(key, [location]);
      else held.push(location);
    }
    known = byKey;
    ownLocations.set(block, known);
  }
  return known;
};

type Holders = readonly BlockContent[];
type Leaf = Map<string, Holders>;
type SuffixPage = Map<string, Leaf>;
type PrefixPage = Map<string, SuffixPage>;
type HolderPages = Map<string, PrefixPage>;

type IndexRoot = {
  type: "root";
  pages: HolderPages;
  other: Map<string, Holders>;
  depth: 0;
};

/**
 * Each overlay contains the keys whose holders changed, including removals.
 * The shared root uses a fixed number of bounded pages per hexadecimal key.
 */
type IndexOverlay = {
  type: "overlay";
  parent: IndexNode;
  holders: Map<string, Holders>;
  depth: number;
};
type IndexNode = IndexRoot | IndexOverlay;

const MAX_CHAIN = 32;

const HEX_ID = /^[0-9A-F]{8}$/;
const isHexId = (key: string) => key.length === 8 && HEX_ID.test(key);

const holdersIn = (node: IndexNode, key: string): Holders => {
  let at = node;
  while (at.type === "overlay") {
    const held = at.holders.get(key);
    if (held !== undefined) return held;
    at = at.parent;
  }
  if (!isHexId(key)) return at.other.get(key) ?? [];
  return (
    at.pages
      .get(key.slice(0, 2))
      ?.get(key.slice(2, 4))
      ?.get(key.slice(4, 6))
      ?.get(key.slice(6, 8)) ?? []
  );
};

const draftPage = <Key, Value>(
  original: Map<Key, Value> | undefined,
  drafts: WeakSet<object>,
): Map<Key, Value> => {
  if (original !== undefined && drafts.has(original)) return original;
  const copy = new Map(original);
  drafts.add(copy);
  return copy;
};

const changedRoot = (
  changes: ReadonlyMap<string, Holders>,
  base?: IndexRoot,
): IndexRoot => {
  const drafts = new WeakSet<object>();
  let pages = base?.pages ?? new Map<string, PrefixPage>();
  let other = base?.other ?? new Map<string, Holders>();
  for (const [key, held] of changes) {
    if (!isHexId(key)) {
      other = draftPage(other, drafts);
      if (held.length === 0) other.delete(key);
      else other.set(key, held);
      continue;
    }
    pages = draftPage(pages, drafts);
    const first = key.slice(0, 2);
    const second = key.slice(2, 4);
    const third = key.slice(4, 6);
    const last = key.slice(6, 8);
    const prefix = draftPage(pages.get(first), drafts);
    const suffix = draftPage(prefix.get(second), drafts);
    const leaf = draftPage(suffix.get(third), drafts);
    if (held.length === 0) leaf.delete(last);
    else leaf.set(last, held);
    if (leaf.size === 0) suffix.delete(third);
    else suffix.set(third, leaf);
    if (suffix.size === 0) prefix.delete(second);
    else prefix.set(second, suffix);
    if (prefix.size === 0) pages.delete(first);
    else pages.set(first, prefix);
  }
  return { type: "root", pages, other, depth: 0 };
};

// A splice on either side of a block preserves its position from the opposite end.
const positions = new WeakMap<
  BlockContent,
  { index: number; fromEnd: number }
>();
const rememberPosition = (
  blocks: readonly BlockContent[],
  index: number,
): void => {
  const block = blocks[index];
  if (block !== undefined)
    positions.set(block, { index, fromEnd: blocks.length - index });
};
const blockPosition = (
  blocks: readonly BlockContent[],
  block: BlockContent,
): number => {
  const known = positions.get(block);
  if (known !== undefined) {
    if (blocks[known.index] === block) return known.index;
    const shifted = blocks.length - known.fromEnd;
    if (blocks[shifted] === block) return shifted;
  }
  const index = blocks.indexOf(block);
  rememberPosition(blocks, index);
  return index;
};

const rootOf = (blocks: readonly BlockContent[]): IndexNode => {
  const holders = new Map<string, BlockContent[]>();
  for (const [index, block] of blocks.entries()) {
    rememberPosition(blocks, index);
    for (const [key, held] of locationsIn(block))
      for (let count = 0; count < held.length; count += 1) {
        const list = holders.get(key);
        if (list === undefined) holders.set(key, [block]);
        else list.push(block);
      }
  }
  return changedRoot(holders);
};

/** Retain only the newest holders per changed key when bounding the chain. */
const settled = (node: IndexOverlay): IndexRoot => {
  const changes = new Map<string, Holders>();
  let at: IndexNode = node;
  while (at.type === "overlay") {
    for (const [key, held] of at.holders)
      if (!changes.has(key)) changes.set(key, held);
    at = at.parent;
  }
  return changedRoot(changes, at);
};

const nodes = new WeakMap<readonly BlockContent[], IndexNode>();
const recent = recentValues<readonly BlockContent[]>();

/**
 * The index of `blocks` from `previous`'s: the ids the differing blocks hold
 * get the holders they have now; every other id keeps its holders.
 */
const derivedFrom = (
  previous: readonly BlockContent[],
  base: IndexNode,
  blocks: readonly BlockContent[],
): IndexNode | undefined => {
  const { start, previousEnd, end } = changedListBounds(previous, blocks);
  const removed = previous.slice(start, previousEnd);
  const added = blocks.slice(start, end);
  // A list that shares little with the base is counted in full.
  if (removed.length + added.length > blocks.length / 2 + 8) return undefined;
  const holders = new Map<string, BlockContent[]>();
  const holdersOf = (key: string): BlockContent[] => {
    let held = holders.get(key);
    if (held === undefined) {
      held = [...holdersIn(base, key)];
      holders.set(key, held);
    }
    return held;
  };
  for (const block of removed)
    for (const [key, located] of locationsIn(block)) {
      const held = holdersOf(key);
      for (let count = 0; count < located.length; count += 1) {
        const at = held.indexOf(block);
        if (at >= 0) held.splice(at, 1);
      }
    }
  for (const [offset, block] of added.entries()) {
    rememberPosition(blocks, start + offset);
    for (const [key, located] of locationsIn(block)) {
      const held = holdersOf(key);
      for (let count = 0; count < located.length; count += 1) held.push(block);
    }
  }
  const node: IndexOverlay = {
    type: "overlay",
    parent: base,
    holders,
    depth: base.depth + 1,
  };
  return node.depth >= MAX_CHAIN ? settled(node) : node;
};

const indexOf = (blocks: readonly BlockContent[]): IndexNode => {
  let node = nodes.get(blocks);
  if (node === undefined) {
    for (const previous of recent.values()) {
      const base = nodes.get(previous);
      if (base === undefined) continue;
      node = derivedFrom(previous, base, blocks);
      break;
    }
    node ??= rootOf(blocks);
    nodes.set(blocks, node);
  }
  recent.remember(blocks);
  return node;
};

/** A location inside a block standing first, moved to where the block stands. */
const placed = (
  location: ParagraphLocation,
  block: number,
): ParagraphLocation => {
  const [first, ...rest] = location.list;
  return first === undefined
    ? { list: [], index: block, paragraph: location.paragraph }
    : {
        list: [{ ...first, block }, ...rest],
        index: location.index,
        paragraph: location.paragraph,
      };
};

/**
 * The paragraph of a story body that `blockId` names, compared as hex, as
 * the first of `storyParagraphs(body)` carrying it.
 */
export const findStoryParagraph = (
  body: DocumentBody,
  blockId: string,
): ParagraphLocation | undefined => {
  const key = idKey(blockId);
  const held = holdersIn(indexOf(body.content), key);
  if (held.length === 0) return undefined;
  const [block] = held;
  const located = block === undefined ? undefined : locationsIn(block).get(key);
  // A repeated id is found as a walk in document order finds it.
  if (held.length > 1 || block === undefined || located?.length !== 1)
    return storyParagraphs(body).find(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === key,
    );
  const [location] = located;
  return location === undefined
    ? undefined
    : placed(location, blockPosition(body.content, block));
};

/** The paragraph and its physical siblings, within the same body, wrapper or cell. */
export const locateParagraphInBlockList = (
  body: DocumentBody,
  blockId: string,
):
  | { paragraph: Paragraph; blocks: readonly BlockContent[]; index: number }
  | undefined => {
  const at =
    blockId === ""
      ? storyParagraphs(body).find(
          ({ paragraph }) => (paragraph.paraId ?? "") === "",
        )
      : findStoryParagraph(body, blockId);
  if (at === undefined) return undefined;
  return {
    paragraph: at.paragraph,
    blocks: blockListAt(body.content, at.list),
    index: at.index,
  };
};
