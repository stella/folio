import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";

import type { FolioRevisionStamp } from "../ai-edits/apply";
import { buildCleanBlockText } from "../ai-edits/clean-text";
import { sourceDocumentOf } from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { resolveAllChangesInHeadlessStateWithMapping } from "../prosemirror/commands/comments";
import { canonicalJson } from "../utils/canonicalJson";
import { prepareTargetInlineAtom } from "./inline-atom-resources";

export type MatchInlineAtomsOptions = {
  state: EditorState;
  targetSnapshot: FolioAIEditSnapshot;
  revisionStamp: FolioRevisionStamp;
  /** First id reserved for this comparison, before planned operations allocated ids. */
  originalRevisionIdSeed: number;
  author: string;
  maxRanges: number;
};

export type MatchInlineAtomsResult =
  | {
      status: "matched";
      transaction: Transaction;
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
      rangeCount: number;
    }
  | { status: "unalignable" }
  | { status: "budget-exceeded" };

type TextBlock = { node: PMNode; from: number };

type InlineAtom = {
  node: PMNode;
  from: number;
  offset: number;
  key: string;
};

type AtomBlock = TextBlock & {
  cleanText: string;
  offsets: readonly number[];
  supported: readonly InlineAtom[];
  unsupportedTopology: readonly string[];
};

type InsertAction = {
  kind: "insert";
  from: number;
  node: PMNode;
  targetBlockId: string | undefined;
};
type DeleteAction = { kind: "delete"; from: number; to: number; targetBlockId: string | undefined };
type Action = InsertAction | DeleteAction;

// Inline atom groups normally contain only a few zero-width carriers. Keep the
// exact ordered match bounded so a hostile document cannot make comparison
// quadratic; larger groups retain only stable unique anchors.
const MAX_INLINE_ATOM_LCS_CELLS = 16_384;

const isSupportedAtom = (node: PMNode): boolean =>
  node.isInline &&
  node.isAtom &&
  (node.type.name === "field" || node.type.name === "image" || node.type.name === "pageBreakRun");

const hasSupportedAtom = (doc: PMNode): boolean => {
  let found = false;
  doc.descendants((node) => {
    if (isSupportedAtom(node)) {
      found = true;
      return false;
    }
    return !found;
  });
  return found;
};

const textBlocksOf = (doc: PMNode): TextBlock[] => {
  const blocks: TextBlock[] = [];
  doc.descendants((node, from) => {
    if (!node.isTextblock) return true;
    blocks.push({ node, from });
    return false;
  });
  return blocks;
};

const offsetAt = ({
  offsets,
  position,
}: {
  offsets: readonly number[];
  position: number;
}): number | null => {
  let low = 0;
  let high = offsets.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    const candidate = offsets[middle];
    if (candidate === undefined || candidate < position) low = middle + 1;
    else high = middle;
  }
  return Math.min(low, offsets.length - 1);
};

const atomKey = (node: PMNode): string =>
  canonicalJson({ type: node.type.name, attrs: node.attrs, content: node.content.toJSON() });

const atomBlockOf = ({ node, from }: TextBlock): AtomBlock | null => {
  const clean = buildCleanBlockText(node, from);
  const supported: InlineAtom[] = [];
  const unsupportedTopology: string[] = [];
  let unalignable = false;
  node.descendants((child, relativePosition) => {
    if (child.isText) return false;
    if (!child.isInline || !child.isAtom) return true;
    const position = from + 1 + relativePosition;
    const offset = offsetAt({ offsets: clean.offsets, position });
    if (offset === null) {
      unalignable = true;
      return false;
    }
    if (!isSupportedAtom(child)) {
      // The layout-only projection is derived from pageBreakRun. Its presence
      // cannot block reconciliation of the serializable carrier itself.
      if (child.type.name === "renderedPageBreak") return false;
      unsupportedTopology.push(`${offset}:${child.type.name}:${canonicalJson(child.attrs)}`);
      return false;
    }
    const prepared = prepareTargetInlineAtom(child);
    if (!prepared) {
      unalignable = true;
      return false;
    }
    supported.push({ node: prepared, from: position, offset, key: atomKey(prepared) });
    return false;
  });
  return unalignable
    ? null
    : { node, from, cleanText: clean.text, offsets: clean.offsets, supported, unsupportedTopology };
};

const targetBlockIdLookup = (snapshot: FolioAIEditSnapshot) => {
  const anchors = Object.values(snapshot.anchors).toSorted(
    (left, right) => left.from - right.from || right.to - left.to,
  );
  return (position: number): string | undefined => {
    let low = 0;
    let high = anchors.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if ((anchors[middle]?.from ?? Number.POSITIVE_INFINITY) <= position) low = middle + 1;
      else high = middle;
    }
    let owner: (typeof anchors)[number] | undefined;
    for (let index = low - 1; index >= 0; index--) {
      const anchor = anchors[index];
      if (!anchor || anchor.to < position) continue;
      if (!owner || anchor.to - anchor.from < owner.to - owner.from) owner = anchor;
    }
    return owner?.id;
  };
};

const atomBlocksOf = (doc: PMNode): AtomBlock[] | null => {
  const blocks: AtomBlock[] = [];
  for (const block of textBlocksOf(doc)) {
    const atoms = atomBlockOf(block);
    if (!atoms) return null;
    blocks.push(atoms);
  }
  return blocks;
};

const sameBlockTopology = (left: AtomBlock, right: AtomBlock): boolean =>
  left.node.type === right.node.type &&
  left.cleanText === right.cleanText &&
  canonicalJson(left.unsupportedTopology) === canonicalJson(right.unsupportedTopology);

const matchingAtomGroup = ({
  live,
  target,
  liveStart,
  targetStart,
}: {
  live: readonly InlineAtom[];
  target: readonly InlineAtom[];
  liveStart: number;
  targetStart: number;
}): readonly [number, number][] => {
  if (
    live.length === target.length &&
    live.every((atom, index) => atom.key === target[index]?.key)
  ) {
    return live.map((_, index) => [liveStart + index, targetStart + index]);
  }
  if (live.length === 0 || target.length === 0) return [];

  if (live.length <= Math.floor(MAX_INLINE_ATOM_LCS_CELLS / target.length)) {
    const width = target.length + 1;
    const table = new Uint16Array((live.length + 1) * width);
    for (let liveIndex = live.length - 1; liveIndex >= 0; liveIndex--) {
      for (let targetIndex = target.length - 1; targetIndex >= 0; targetIndex--) {
        const index = liveIndex * width + targetIndex;
        if (live[liveIndex]?.key === target[targetIndex]?.key) {
          table[index] = (table[(liveIndex + 1) * width + targetIndex + 1] ?? 0) + 1;
          continue;
        }
        table[index] = Math.max(
          table[(liveIndex + 1) * width + targetIndex] ?? 0,
          table[liveIndex * width + targetIndex + 1] ?? 0,
        );
      }
    }
    const matches: [number, number][] = [];
    let liveIndex = 0;
    let targetIndex = 0;
    while (liveIndex < live.length && targetIndex < target.length) {
      if (live[liveIndex]?.key === target[targetIndex]?.key) {
        matches.push([liveStart + liveIndex, targetStart + targetIndex]);
        liveIndex++;
        targetIndex++;
        continue;
      }
      const skipLive = table[(liveIndex + 1) * width + targetIndex] ?? 0;
      const skipTarget = table[liveIndex * width + targetIndex + 1] ?? 0;
      if (skipLive >= skipTarget) liveIndex++;
      else targetIndex++;
    }
    return matches;
  }

  const targetIndexes = new Map<string, number | null>();
  for (const [index, atom] of target.entries()) {
    targetIndexes.set(atom.key, targetIndexes.has(atom.key) ? null : index);
  }
  const liveKeyCounts = new Map<string, number>();
  for (const atom of live) {
    liveKeyCounts.set(atom.key, (liveKeyCounts.get(atom.key) ?? 0) + 1);
  }
  const matches: [number, number][] = [];
  let previousTargetIndex = -1;
  for (const [liveIndex, atom] of live.entries()) {
    const targetIndex = targetIndexes.get(atom.key);
    if (
      targetIndex === undefined ||
      targetIndex === null ||
      targetIndex <= previousTargetIndex ||
      liveKeyCounts.get(atom.key) !== 1
    ) {
      continue;
    }
    matches.push([liveStart + liveIndex, targetStart + targetIndex]);
    previousTargetIndex = targetIndex;
  }
  return matches;
};

const matchingAtoms = ({
  live,
  target,
}: {
  live: readonly InlineAtom[];
  target: readonly InlineAtom[];
}): readonly [number, number][] => {
  const matched: [number, number][] = [];
  let left = 0;
  let right = 0;
  while (left < live.length || right < target.length) {
    const liveOffset = live[left]?.offset;
    const targetOffset = target[right]?.offset;
    const offset = Math.min(
      liveOffset ?? Number.POSITIVE_INFINITY,
      targetOffset ?? Number.POSITIVE_INFINITY,
    );
    if (!Number.isFinite(offset)) break;
    const liveStart = left;
    const targetStart = right;
    while (live[left]?.offset === offset) left++;
    while (target[right]?.offset === offset) right++;
    const liveGroup = live.slice(liveStart, left);
    const targetGroup = target.slice(targetStart, right);
    matched.push(
      ...matchingAtomGroup({ live: liveGroup, target: targetGroup, liveStart, targetStart }),
    );
  }
  return matched;
};

const insertionAnchors = ({
  live,
  target,
  matches,
}: {
  live: readonly InlineAtom[];
  target: readonly InlineAtom[];
  matches: readonly [number, number][];
}): ReadonlyMap<number, number> => {
  const liveIndexByTargetIndex = new Map<number, number>();
  for (const [liveIndex, targetIndex] of matches) {
    liveIndexByTargetIndex.set(targetIndex, liveIndex);
  }
  const anchors = new Map<number, number>();
  let groupStart = 0;
  while (groupStart < target.length) {
    const offset = target[groupStart]?.offset;
    let groupEnd = groupStart + 1;
    while (target[groupEnd]?.offset === offset) groupEnd++;
    let nextMatchedLiveIndex: number | undefined;
    for (let index = groupEnd - 1; index >= groupStart; index--) {
      const matchedLiveIndex = liveIndexByTargetIndex.get(index);
      if (matchedLiveIndex !== undefined) nextMatchedLiveIndex = matchedLiveIndex;
      else if (nextMatchedLiveIndex !== undefined)
        anchors.set(index, live[nextMatchedLiveIndex]!.from);
    }
    let previousMatchedLiveIndex: number | undefined;
    for (let index = groupStart; index < groupEnd; index++) {
      const matchedLiveIndex = liveIndexByTargetIndex.get(index);
      if (matchedLiveIndex !== undefined) {
        previousMatchedLiveIndex = matchedLiveIndex;
        continue;
      }
      if (!anchors.has(index) && previousMatchedLiveIndex !== undefined) {
        const previous = live[previousMatchedLiveIndex];
        if (previous) anchors.set(index, previous.from + previous.node.nodeSize);
      }
    }
    groupStart = groupEnd;
  }
  return anchors;
};

const mappedSourcePosition = ({
  mapping,
  position,
}: {
  mapping: ReturnType<typeof resolveAllChangesInHeadlessStateWithMapping>["mapping"];
  position: number;
}): number | null => {
  const inverse = mapping.invert();
  const right = inverse.mapResult(position, 1);
  if (!right.deleted) return right.pos;
  const left = inverse.mapResult(position, -1);
  return left.deleted ? null : left.pos;
};

const sourceTextBlockIds = (source: PMNode): ReadonlyMap<string, TextBlock | null> => {
  const blocks = new Map<string, TextBlock | null>();
  for (const block of textBlocksOf(source)) {
    const paraId = block.node.attrs["paraId"];
    if (typeof paraId !== "string" || paraId.length === 0) continue;
    blocks.set(paraId, blocks.has(paraId) ? null : block);
  }
  return blocks;
};

const sameParagraphSourcePosition = ({
  sourceBlocks,
  reviewed,
  offset,
}: {
  sourceBlocks: ReadonlyMap<string, TextBlock | null>;
  reviewed: AtomBlock;
  offset: number;
}): number | null => {
  const paraId = reviewed.node.attrs["paraId"];
  if (typeof paraId !== "string" || paraId.length === 0) return null;
  const source = sourceBlocks.get(paraId);
  if (!source || source.node.type !== reviewed.node.type) return null;
  const clean = buildCleanBlockText(source.node, source.from);
  return clean.text === reviewed.cleanText ? (clean.offsets[offset] ?? null) : null;
};

/**
 * Restore field, image, and page-break atoms omitted by text-only comparison operations.
 * Positions come from the legacy review resolver's mapping, never a guessed
 * paragraph correspondence: paragraph-mark deletions can merge paragraphs.
 */
export const matchInlineAtoms = ({
  state,
  targetSnapshot,
  revisionStamp,
  originalRevisionIdSeed,
  author,
  maxRanges,
}: MatchInlineAtomsOptions): MatchInlineAtomsResult => {
  if (!Number.isSafeInteger(maxRanges) || maxRanges < 0) return { status: "budget-exceeded" };
  const targetDocument = sourceDocumentOf(targetSnapshot);
  if (!hasSupportedAtom(state.doc) && !hasSupportedAtom(targetDocument)) {
    return {
      status: "matched",
      transaction: state.tr,
      nextRevisionId: revisionStamp.idSeed,
      changedTargetBlockIds: [],
      rangeCount: 0,
    };
  }

  const reviewed = resolveAllChangesInHeadlessStateWithMapping(state, "accept");
  const targetBlockIdAt = targetBlockIdLookup(targetSnapshot);
  let sourceBlocks: ReadonlyMap<string, TextBlock | null> | undefined;
  const liveBlocks = atomBlocksOf(reviewed.state.doc);
  const targetBlocks = atomBlocksOf(targetDocument);
  if (!liveBlocks || !targetBlocks || liveBlocks.length !== targetBlocks.length) {
    return { status: "unalignable" };
  }

  const actions: Action[] = [];
  for (const [index, live] of liveBlocks.entries()) {
    const target = targetBlocks[index];
    if (!target) return { status: "unalignable" };
    if (live.supported.length === 0 && target.supported.length === 0) continue;
    if (!sameBlockTopology(live, target)) return { status: "unalignable" };
    const matches = matchingAtoms({ live: live.supported, target: target.supported });
    const matchedLive = new Set(matches.map(([liveIndex]) => liveIndex));
    const matchedTarget = new Set(matches.map(([, targetIndex]) => targetIndex));
    const anchors = insertionAnchors({
      live: live.supported,
      target: target.supported,
      matches,
    });
    for (const [liveIndex, atom] of live.supported.entries()) {
      if (matchedLive.has(liveIndex)) continue;
      const from = mappedSourcePosition({ mapping: reviewed.mapping, position: atom.from });
      if (from === null) return { status: "unalignable" };
      actions.push({
        kind: "delete",
        from,
        to: from + atom.node.nodeSize,
        targetBlockId: targetBlockIdAt(target.from),
      });
    }
    for (const [targetIndex, atom] of target.supported.entries()) {
      if (matchedTarget.has(targetIndex)) continue;
      const position = anchors.get(targetIndex) ?? live.offsets[atom.offset];
      if (position === undefined) return { status: "unalignable" };
      const from =
        mappedSourcePosition({ mapping: reviewed.mapping, position }) ??
        sameParagraphSourcePosition({
          sourceBlocks: (sourceBlocks ??= sourceTextBlockIds(state.doc)),
          reviewed: live,
          offset: atom.offset,
        });
      if (from === null) return { status: "unalignable" };
      actions.push({
        kind: "insert",
        from,
        node: atom.node,
        targetBlockId: targetBlockIdAt(target.from),
      });
    }
  }
  if (actions.length > maxRanges) return { status: "budget-exceeded" };

  const insertionType = state.schema.marks["insertion"];
  const deletionType = state.schema.marks["deletion"];
  if (!insertionType || !deletionType) return { status: "unalignable" };
  const transaction = state.tr;
  let nextRevisionId = revisionStamp.idSeed;
  const changedTargetBlockIds = new Set<string>();
  const ordered = actions.toSorted(
    (left, right) => right.from - left.from || (left.kind === "delete" ? -1 : 1),
  );
  for (const action of ordered) {
    if (action.kind === "insert") {
      const at = transaction.mapping.map(action.from, 1);
      const marked = action.node.mark(
        insertionType
          .create({ revisionId: nextRevisionId++, author, date: revisionStamp.date })
          .addToSet(action.node.marks),
      );
      transaction.insert(at, marked);
      if (action.targetBlockId) changedTargetBlockIds.add(action.targetBlockId);
      continue;
    }
    const from = transaction.mapping.map(action.from, 1);
    const to = transaction.mapping.map(action.to, -1);
    const node = transaction.doc.nodeAt(from);
    if (!node || node.nodeSize !== to - from || !isSupportedAtom(node))
      return { status: "unalignable" };
    const insertion = node.marks.find(({ type }) => type === insertionType);
    if (insertion) {
      const revisionId = insertion.attrs["revisionId"];
      if (typeof revisionId !== "number" || revisionId < originalRevisionIdSeed)
        return { status: "unalignable" };
      transaction.delete(from, to);
      if (action.targetBlockId) changedTargetBlockIds.add(action.targetBlockId);
      continue;
    }
    if (node.marks.some(({ type }) => type === deletionType)) return { status: "unalignable" };
    transaction.addMark(
      from,
      to,
      deletionType.create({ revisionId: nextRevisionId++, author, date: revisionStamp.date }),
    );
    if (action.targetBlockId) changedTargetBlockIds.add(action.targetBlockId);
  }
  return {
    status: "matched",
    transaction,
    nextRevisionId,
    changedTargetBlockIds: [...changedTargetBlockIds],
    rangeCount: actions.length,
  };
};

/** Compare supported inline atom identity after accept or reject projection. */
export const sameInlineAtoms = (leftDocument: PMNode, rightDocument: PMNode): boolean => {
  if (!hasSupportedAtom(leftDocument) && !hasSupportedAtom(rightDocument)) return true;
  const leftBlocks = atomBlocksOf(leftDocument);
  const rightBlocks = atomBlocksOf(rightDocument);
  if (!leftBlocks || !rightBlocks || leftBlocks.length !== rightBlocks.length) return false;
  return leftBlocks.every((left, index) => {
    const right = rightBlocks[index];
    if (left.supported.length === 0 && right?.supported.length === 0) return true;
    return (
      right !== undefined &&
      sameBlockTopology(left, right) &&
      canonicalJson(left.supported.map(({ offset, key }) => ({ offset, key }))) ===
        canonicalJson(right.supported.map(({ offset, key }) => ({ offset, key })))
    );
  });
};
