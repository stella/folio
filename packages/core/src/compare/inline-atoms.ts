import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";

import type { FolioRevisionStamp } from "../ai-edits/apply";
import { buildCleanBlockText, type BuildCleanBlockTextOptions } from "../ai-edits/clean-text";
import { sourceDocumentOf } from "../ai-edits/snapshot";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { resolveAllChangesInHeadlessStateWithMapping } from "../prosemirror/commands/comments";
import { runFormattingInlineAtomResultText } from "../prosemirror/runFormattingInlineCarriers";
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
  unsupportedTopology: readonly { offset: number; key: string }[];
};

type InsertAction = {
  kind: "insert";
  from: number;
  node: PMNode;
  targetBlockId: string | undefined;
};
type DeleteAction = { kind: "delete"; from: number; to: number; targetBlockId: string | undefined };
type ReplaceTextAction = {
  kind: "replace-text";
  from: number;
  to: number;
  node: PMNode;
  textDisposition: "inserted" | "retained";
  targetBlockId: string | undefined;
};
type ReplaceAtomAction = {
  kind: "replace-atom";
  from: number;
  to: number;
  text: string;
  atomDisposition: "inserted" | "retained";
  marks: PMNode["marks"];
  targetBlockId: string | undefined;
};
type Action = InsertAction | DeleteAction | ReplaceTextAction | ReplaceAtomAction;

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

/**
 * Clean-text offset nearest a document position.
 *
 * `buildCleanBlockText` returns one offset per character plus one for
 * end-of-block, so there is always an offset to land on. The result was typed
 * `number | null` and never was: the caller's refusal branch could not run,
 * while an empty `offsets` would have produced `-1` rather than the `null` the
 * type promised.
 */
const offsetAt = ({
  offsets,
  position,
}: {
  offsets: readonly number[];
  position: number;
}): number => {
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

/**
 * Attributes a conversion mints for itself rather than reads from the package.
 *
 * `toProseDoc` salts text-box anchor ids with a per-load nonce so a pasted span
 * cannot hijack a real anchor. A comparison reads two independent conversions,
 * so a minted handle differs on every run, and a key holding one makes a
 * document differ from itself. Only document facts belong in a key.
 *
 * The guard against a node type gaining a minted attribute without being added
 * here is not this list but `compareIdentity.property.test.ts`, which compares
 * generated packages carrying each atom with themselves.
 */
const CONVERSION_LOCAL_ATTRS: Readonly<Record<string, readonly string[]>> = {
  textBoxAnchor: ["anchorId"],
};

const documentFactAttrs = (node: PMNode): Record<string, unknown> => {
  const local = CONVERSION_LOCAL_ATTRS[node.type.name];
  if (!local) return node.attrs;
  return Object.fromEntries(Object.entries(node.attrs).filter(([name]) => !local.includes(name)));
};

const atomKey = (node: PMNode): string =>
  canonicalJson({
    type: node.type.name,
    attrs: documentFactAttrs(node),
    content: node.content.toJSON(),
  });

const atomBlockOf = (
  { node, from }: TextBlock,
  fieldResults: BuildCleanBlockTextOptions["fieldResults"],
): AtomBlock => {
  const clean = buildCleanBlockText(node, from, { fieldResults });
  const supported: InlineAtom[] = [];
  const unsupportedTopology: { offset: number; key: string }[] = [];
  node.descendants((child, relativePosition) => {
    if (child.isText) return false;
    if (!child.isInline || !child.isAtom) return true;
    const position = from + 1 + relativePosition;
    const offset = offsetAt({ offsets: clean.offsets, position });
    if (!isSupportedAtom(child)) {
      // The layout-only projection is derived from pageBreakRun. Its presence
      // cannot block reconciliation of the serializable carrier itself.
      if (child.type.name === "renderedPageBreak") return false;
      unsupportedTopology.push({
        offset,
        key: `${child.type.name}:${canonicalJson(documentFactAttrs(child))}`,
      });
      return false;
    }
    const prepared = prepareTargetInlineAtom(child);
    if (!prepared) {
      // An atom this comparison cannot detach from its package — an image with
      // no embedded media, a drawing that is a chart rather than a picture — is
      // one it cannot restore, not one that makes the story unalignable. Its
      // identity still has to match, so it joins the topology both sides are
      // compared on instead of abandoning the alignment for the whole story.
      unsupportedTopology.push({
        offset,
        key: `${child.type.name}:${canonicalJson(documentFactAttrs(child))}`,
      });
      return false;
    }
    supported.push({ node: prepared, from: position, offset, key: atomKey(prepared) });
    return false;
  });
  return {
    node,
    from,
    cleanText: clean.text,
    offsets: clean.offsets,
    supported,
    unsupportedTopology,
  };
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

const atomBlocksOf = (
  doc: PMNode,
  fieldResults: BuildCleanBlockTextOptions["fieldResults"],
): AtomBlock[] => textBlocksOf(doc).map((block) => atomBlockOf(block, fieldResults));

const sameBlockTopology = (left: AtomBlock, right: AtomBlock): boolean => {
  if (left.node.type !== right.node.type || left.cleanText !== right.cleanText) return false;
  const supportedOffsets = new Set([
    ...left.supported.map(({ offset }) => offset),
    ...right.supported.map(({ offset }) => offset),
  ]);
  const relevantTopology = ({ unsupportedTopology }: AtomBlock) =>
    unsupportedTopology.filter(({ offset }) => supportedOffsets.has(offset));
  return canonicalJson(relevantTopology(left)) === canonicalJson(relevantTopology(right));
};

const supportedAtomProjection = (block: AtomBlock) =>
  canonicalJson(block.supported.map(({ offset, key }) => ({ offset, key })));

const sameSupportedAtoms = (left: AtomBlock, right: AtomBlock): boolean =>
  left.node.type === right.node.type &&
  supportedAtomProjection(left) === supportedAtomProjection(right);

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
  fieldResults,
}: {
  sourceBlocks: ReadonlyMap<string, TextBlock | null>;
  reviewed: AtomBlock;
  offset: number;
  fieldResults: BuildCleanBlockTextOptions["fieldResults"];
}): number | null => {
  const paraId = reviewed.node.attrs["paraId"];
  if (typeof paraId !== "string" || paraId.length === 0) return null;
  const source = sourceBlocks.get(paraId);
  if (!source || source.node.type !== reviewed.node.type) return null;
  const clean = buildCleanBlockText(source.node, source.from, { fieldResults });
  return clean.text === reviewed.cleanText ? (clean.offsets[offset] ?? null) : null;
};

const textRangeDisposition = ({
  doc,
  from,
  to,
  expectedText,
  originalRevisionIdSeed,
}: {
  doc: PMNode;
  from: number;
  to: number;
  expectedText: string;
  originalRevisionIdSeed: number;
}): ReplaceTextAction["textDisposition"] | null => {
  if (from >= to || doc.textBetween(from, to) !== expectedText) return null;
  let disposition: ReplaceTextAction["textDisposition"] | undefined;
  let covered = 0;
  let invalid = false;
  doc.nodesBetween(from, to, (node, position) => {
    if (!node.isText) {
      if (position > from && position < to) invalid = true;
      return !invalid;
    }
    const overlapFrom = Math.max(from, position);
    const overlapTo = Math.min(to, position + node.nodeSize);
    if (overlapFrom >= overlapTo) return false;
    covered += overlapTo - overlapFrom;
    if (node.marks.some(({ type }) => type.name === "deletion")) {
      invalid = true;
      return false;
    }
    const insertion = node.marks.find(({ type }) => type.name === "insertion");
    let current: ReplaceTextAction["textDisposition"] = "retained";
    if (insertion) {
      const revisionId = insertion.attrs["revisionId"];
      if (typeof revisionId !== "number" || revisionId < originalRevisionIdSeed) {
        invalid = true;
        return false;
      }
      current = "inserted";
    }
    if (disposition !== undefined && disposition !== current) {
      invalid = true;
      return false;
    }
    disposition = current;
    return false;
  });
  return !invalid && covered === to - from ? (disposition ?? null) : null;
};

const atomDisposition = ({
  atom,
  originalRevisionIdSeed,
}: {
  atom: InlineAtom;
  originalRevisionIdSeed: number;
}): ReplaceAtomAction["atomDisposition"] | null => {
  if (atom.node.marks.some(({ type }) => type.name === "deletion")) return null;
  const insertion = atom.node.marks.find(({ type }) => type.name === "insertion");
  if (!insertion) return "retained";
  const revisionId = insertion.attrs["revisionId"];
  return typeof revisionId === "number" && revisionId >= originalRevisionIdSeed ? "inserted" : null;
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
  const liveBlocks = atomBlocksOf(reviewed.state.doc, "text");
  const targetBlocks = atomBlocksOf(targetDocument, "text");
  if (liveBlocks.length !== targetBlocks.length) {
    return { status: "unalignable" };
  }

  const actions: Action[] = [];
  let omittedLiveBlocks: readonly AtomBlock[] | undefined;
  let omittedTargetBlocks: readonly AtomBlock[] | undefined;
  for (const [index, fullLive] of liveBlocks.entries()) {
    const fullTarget = targetBlocks[index];
    if (!fullTarget) return { status: "unalignable" };
    if (fullLive.supported.length === 0 && fullTarget.supported.length === 0) continue;
    // An unsupported zero-width neighbor cannot make an unchanged supported
    // carrier need reconciliation. Package-local comment ids are the common
    // case: they may be rebound independently while the page break beside
    // them remains identical.
    if (sameSupportedAtoms(fullLive, fullTarget)) continue;
    let live = fullLive;
    let target = fullTarget;
    let fieldResults: BuildCleanBlockTextOptions["fieldResults"] = "text";
    if (!sameBlockTopology(live, target)) {
      omittedLiveBlocks ??= atomBlocksOf(reviewed.state.doc, "omitted");
      omittedTargetBlocks ??= atomBlocksOf(targetDocument, "omitted");
      const omittedLive = omittedLiveBlocks[index];
      const omittedTarget = omittedTargetBlocks[index];
      if (!omittedLive || !omittedTarget || !sameBlockTopology(omittedLive, omittedTarget)) {
        return { status: "unalignable" };
      }
      live = omittedLive;
      target = omittedTarget;
      fieldResults = "omitted";
    }
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
      const resultText = runFormattingInlineAtomResultText(atom.node);
      const targetFieldOwnsResult = target.supported.some(
        (candidate) =>
          candidate.offset === atom.offset &&
          runFormattingInlineAtomResultText(candidate.node) === resultText,
      );
      if (fieldResults === "text" && resultText && !targetFieldOwnsResult) {
        const disposition = atomDisposition({ atom, originalRevisionIdSeed });
        if (disposition === null) return { status: "unalignable" };
        actions.push({
          kind: "replace-atom",
          from,
          to: from + atom.node.nodeSize,
          text: resultText,
          atomDisposition: disposition,
          marks: atom.node.marks,
          targetBlockId: targetBlockIdAt(target.from),
        });
        continue;
      }
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
          fieldResults,
        });
      if (from === null) return { status: "unalignable" };
      const resultText = runFormattingInlineAtomResultText(atom.node);
      const liveFieldOwnsResult = live.supported.some(
        (candidate) =>
          candidate.offset === atom.offset &&
          runFormattingInlineAtomResultText(candidate.node) === resultText,
      );
      if (fieldResults === "text" && resultText && !liveFieldOwnsResult) {
        const reviewedTo = live.offsets[atom.offset + resultText.length];
        const to =
          reviewedTo === undefined
            ? null
            : mappedSourcePosition({ mapping: reviewed.mapping, position: reviewedTo });
        if (to === null) return { status: "unalignable" };
        const textDisposition = textRangeDisposition({
          doc: state.doc,
          from,
          to,
          expectedText: resultText,
          originalRevisionIdSeed,
        });
        if (textDisposition === null) return { status: "unalignable" };
        actions.push({
          kind: "replace-text",
          from,
          to,
          node: atom.node,
          textDisposition,
          targetBlockId: targetBlockIdAt(target.from),
        });
        continue;
      }
      actions.push({
        kind: "insert",
        from,
        node: atom.node,
        targetBlockId: targetBlockIdAt(target.from),
      });
    }
  }
  const rangeCount = actions.reduce(
    (count, action) =>
      count +
      ((action.kind === "replace-text" && action.textDisposition === "retained") ||
      (action.kind === "replace-atom" && action.atomDisposition === "retained")
        ? 2
        : 1),
    0,
  );
  if (rangeCount > maxRanges) return { status: "budget-exceeded" };

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
    if (action.kind === "replace-text") {
      const from = transaction.mapping.map(action.from, 1);
      const to = transaction.mapping.map(action.to, -1);
      if (from >= to) return { status: "unalignable" };
      if (action.textDisposition === "inserted") {
        transaction.delete(from, to);
      } else {
        transaction.addMark(
          from,
          to,
          deletionType.create({ revisionId: nextRevisionId++, author, date: revisionStamp.date }),
        );
      }
      const at = transaction.mapping.map(action.from, -1);
      transaction.insert(
        at,
        action.node.mark(
          insertionType
            .create({ revisionId: nextRevisionId++, author, date: revisionStamp.date })
            .addToSet(action.node.marks),
        ),
      );
      if (action.targetBlockId) changedTargetBlockIds.add(action.targetBlockId);
      continue;
    }
    if (action.kind === "replace-atom") {
      const from = transaction.mapping.map(action.from, 1);
      const to = transaction.mapping.map(action.to, -1);
      const node = transaction.doc.nodeAt(from);
      if (!node || node.nodeSize !== to - from || node.type.name !== "field") {
        return { status: "unalignable" };
      }
      if (action.atomDisposition === "inserted") {
        transaction.delete(from, to);
      } else {
        transaction.addMark(
          from,
          to,
          deletionType.create({ revisionId: nextRevisionId++, author, date: revisionStamp.date }),
        );
      }
      const at = transaction.mapping.map(action.from, -1);
      transaction.insert(
        at,
        state.schema.text(
          action.text,
          insertionType
            .create({ revisionId: nextRevisionId++, author, date: revisionStamp.date })
            .addToSet(action.marks),
        ),
      );
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
    rangeCount,
  };
};

/** Compare supported inline atom identity after accept or reject projection. */
export const sameInlineAtoms = (leftDocument: PMNode, rightDocument: PMNode): boolean => {
  if (!hasSupportedAtom(leftDocument) && !hasSupportedAtom(rightDocument)) return true;
  const leftBlocks = atomBlocksOf(leftDocument, "omitted");
  const rightBlocks = atomBlocksOf(rightDocument, "omitted");
  if (leftBlocks.length !== rightBlocks.length) return false;
  return leftBlocks.every((left, index) => {
    const right = rightBlocks[index];
    if (left.supported.length === 0 && right?.supported.length === 0) return true;
    return (
      right !== undefined && left.cleanText === right.cleanText && sameSupportedAtoms(left, right)
    );
  });
};
