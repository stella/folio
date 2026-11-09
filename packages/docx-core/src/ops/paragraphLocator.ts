import type { BlockContent, DocumentBody } from "../model/document";
import { type ParagraphLocation, storyParagraphs } from "./blocks";
import { idKey } from "./ids";

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

/**
 * The top-level blocks holding each id: every one (a root) or the ids whose
 * holders changed from the list it was derived from. Lookups take the first
 * answer up the chain, which {@link MAX_CHAIN} bounds.
 */
type IndexNode = {
  parent: IndexNode | undefined;
  holders: Map<string, readonly BlockContent[]>;
  depth: number;
};

const MAX_CHAIN = 32;

/** How many recently indexed block lists a new index may derive from. */
const RECENT_BASES = 8;

const holdersIn = (node: IndexNode, key: string): readonly BlockContent[] => {
  for (let at: IndexNode | undefined = node; at !== undefined; at = at.parent) {
    const held = at.holders.get(key);
    if (held !== undefined) return held;
  }
  return [];
};

const rootOf = (blocks: readonly BlockContent[]): IndexNode => {
  const holders = new Map<string, BlockContent[]>();
  for (const block of blocks)
    for (const [key, held] of locationsIn(block))
      for (let count = 0; count < held.length; count += 1) {
        const list = holders.get(key);
        if (list === undefined) holders.set(key, [block]);
        else list.push(block);
      }
  return { parent: undefined, holders, depth: 0 };
};

/** Turns a node into a root holding every id's holders. */
const settled = (node: IndexNode): IndexNode => {
  const holders = new Map<string, readonly BlockContent[]>();
  const chain: IndexNode[] = [];
  for (let at: IndexNode | undefined = node; at !== undefined; at = at.parent)
    chain.push(at);
  for (const at of chain.toReversed())
    for (const [key, held] of at.holders) holders.set(key, held);
  for (const [key, held] of holders) if (held.length === 0) holders.delete(key);
  return { parent: undefined, holders, depth: 0 };
};

const nodes = new WeakMap<readonly BlockContent[], IndexNode>();
const recent: WeakRef<readonly BlockContent[]>[] = [];

const remember = (blocks: readonly BlockContent[]): void => {
  if (recent[0]?.deref() === blocks) return;
  const alive = recent.filter((reference) => {
    const held = reference.deref();
    return held !== undefined && held !== blocks;
  });
  recent.length = 0;
  recent.push(new WeakRef(blocks), ...alive.slice(0, RECENT_BASES - 1));
};

/**
 * The index of `blocks` from `previous`'s: the ids the differing blocks hold
 * get the holders they have now; every other id keeps its holders.
 */
const derivedFrom = (
  previous: readonly BlockContent[],
  base: IndexNode,
  blocks: readonly BlockContent[],
): IndexNode | undefined => {
  let prefix = 0;
  while (
    prefix < blocks.length &&
    prefix < previous.length &&
    blocks[prefix] === previous[prefix]
  )
    prefix += 1;
  let suffix = 0;
  while (
    suffix < blocks.length - prefix &&
    suffix < previous.length - prefix &&
    blocks[blocks.length - 1 - suffix] ===
      previous[previous.length - 1 - suffix]
  )
    suffix += 1;
  const removed = previous.slice(prefix, previous.length - suffix);
  const added = blocks.slice(prefix, blocks.length - suffix);
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
  for (const block of added)
    for (const [key, located] of locationsIn(block)) {
      const held = holdersOf(key);
      for (let count = 0; count < located.length; count += 1) held.push(block);
    }
  const node: IndexNode = { parent: base, holders, depth: base.depth + 1 };
  return node.depth >= MAX_CHAIN ? settled(node) : node;
};

const indexOf = (blocks: readonly BlockContent[]): IndexNode => {
  let node = nodes.get(blocks);
  if (node === undefined) {
    for (const reference of recent) {
      const previous = reference.deref();
      const base = previous === undefined ? undefined : nodes.get(previous);
      if (previous === undefined || base === undefined) continue;
      node = derivedFrom(previous, base, blocks);
      break;
    }
    node ??= rootOf(blocks);
    nodes.set(blocks, node);
  }
  remember(blocks);
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
    : placed(location, body.content.indexOf(block));
};
