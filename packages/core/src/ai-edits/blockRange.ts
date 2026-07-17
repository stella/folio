/**
 * Helpers shared by the adapters' AI-edit imperative API methods (React and
 * Vue both resolve block/revision ranges through here, so the behavior can
 * never drift).
 *
 * Framework-neutral so the bounds-checking logic that protects
 * `TextSelection.between` from "endpoint not pointing into a node with
 * inline content" can be unit-tested without spinning up a real PM view.
 */

import type { Node as PMNode } from "prosemirror-model";

import { buildCleanBlockText } from "./clean-text";
import { createFolioAIEditSnapshot, hashFolioAIBlockText } from "./snapshot";
import type { FolioAIBlockAnchor, FolioAIEditSnapshot, FolioAITextRangeHandle } from "./types";
import { findParagraphByParaId } from "../prosemirror/utils/findParagraphByParaId";
import { getFolioParaIdFromBlockId, getSequentialFolioBlockIdIndex } from "../types/block-id";

export type DocPositionRange = { from: number; to: number };

type ResolveFolioAIBlockRangeOptions = {
  blockId: string;
  doc: PMNode;
  snapshot?: FolioAIEditSnapshot | null | undefined;
};

export const resolveFolioAIBlockRange = ({
  blockId,
  doc,
  snapshot,
}: ResolveFolioAIBlockRangeOptions): DocPositionRange | null => {
  const paraId = getFolioParaIdFromBlockId(blockId);
  if (paraId !== null) {
    const liveRange = findParagraphByParaId(doc, paraId);
    if (liveRange !== null) {
      return { from: liveRange.from, to: liveRange.to };
    }
  }

  const resolvedSnapshot = snapshot ?? createFolioAIEditSnapshot(doc);
  const anchor =
    resolvedSnapshot.anchors[blockId] ?? resolveSequentialBlockAnchor(blockId, resolvedSnapshot);
  if (!anchor) {
    return null;
  }
  return clampRangeToDocSize(doc.content.size, anchor);
};

type ResolveFolioAITextRangeOptions = {
  range: FolioAITextRangeHandle;
  doc: PMNode;
  snapshot?: FolioAIEditSnapshot | null | undefined;
};

/** Resolve and validate a stable text-range handle against the live document. */
export const resolveFolioAITextRange = ({
  range,
  doc,
  snapshot,
}: ResolveFolioAITextRangeOptions): DocPositionRange | null => {
  const blockRange = resolveFolioAIBlockRange({
    blockId: range.blockId,
    doc,
    snapshot,
  });
  if (blockRange === null) {
    return null;
  }
  const blockNode = doc.nodeAt(blockRange.from);
  if (!blockNode?.isTextblock) {
    return null;
  }

  const cleanBlock = buildCleanBlockText(blockNode, blockRange.from);
  const from = cleanBlock.offsets[range.startOffset];
  const to = cleanBlock.offsets[range.endOffset];
  if (from === undefined || to === undefined) {
    return null;
  }
  const selectedText = cleanBlock.text.slice(range.startOffset, range.endOffset);
  if (hashFolioAIBlockText(selectedText) !== range.selectedTextHash) {
    return null;
  }
  return clampRangeToDocSize(doc.content.size, { from, to });
};

type ResolvePassageRangeOptions = {
  blockId: string;
  text: string;
  doc: PMNode;
  snapshot?: FolioAIEditSnapshot | null | undefined;
};

/**
 * Resolve a free-text passage inside a block to a `{ from, to }` PM range.
 *
 * First resolves the block the same way {@link resolveFolioAIBlockRange} does
 * (live paraId lookup, then snapshot/seq fallback). Then matches `text` against
 * the block's post-tracked-changes clean text (so the offsets it finds still
 * map back to live PM positions on a doc with pending redlines) and maps the
 * match back through the clean-text `offsets`.
 *
 * Matching normalizes whitespace runs in the needle to `\s+` (typography inserts
 * non-breaking / narrow spaces that the caller's quoted text won't reproduce
 * exactly) and tries case-sensitive first, then case-insensitive; the first
 * match wins. Returns null when the block does not resolve, is not a textblock,
 * the needle is empty, or the text does not occur in the block.
 */
export const resolvePassageRange = ({
  blockId,
  text,
  doc,
  snapshot,
}: ResolvePassageRangeOptions): DocPositionRange | null => {
  const blockRange = resolveFolioAIBlockRange({ blockId, doc, snapshot });
  if (blockRange === null) {
    return null;
  }
  const blockNode = doc.nodeAt(blockRange.from);
  if (!blockNode?.isTextblock) {
    return null;
  }

  const needleSource = buildPassageNeedleSource(text);
  if (needleSource === null) {
    return null;
  }

  const cleanBlock = buildCleanBlockText(blockNode, blockRange.from);
  const match = matchPassage(cleanBlock.text, needleSource);
  if (match === null) {
    return null;
  }

  const from = cleanBlock.offsets[match.start];
  const to = cleanBlock.offsets[match.end];
  if (from === undefined || to === undefined) {
    return null;
  }
  return clampRangeToDocSize(doc.content.size, { from, to });
};

/**
 * Turn a caller-supplied passage into a regex source that tolerates whitespace
 * variance. The needle is trimmed first (quoted passages routinely carry
 * accidental leading/trailing whitespace that the block text won't reproduce),
 * then regex metacharacters are escaped so the literal text matches, then
 * whitespace runs collapse to `\s+` (mirrors the anonymization matcher).
 * Returns null for empty / whitespace-only input so an empty needle can never
 * match at offset 0 and paint a zero-width highlight.
 */
const buildPassageNeedleSource = (text: string): string | null => {
  const trimmed = text.trim();
  if (trimmed.length === 0) {
    return null;
  }
  return trimmed.replaceAll(/[\\^$.*+?()[\]{}|]/gu, "\\$&").replaceAll(/\s+/gu, "\\s+");
};

type PassageMatch = { start: number; end: number };

/**
 * Case-sensitive match first, then case-insensitive fallback. The first hit in
 * each pass wins; `index` and length are clean-text offsets the caller maps back
 * through `offsets`.
 */
const matchPassage = (haystack: string, needleSource: string): PassageMatch | null => {
  for (const flags of ["u", "iu"]) {
    const match = new RegExp(needleSource, flags).exec(haystack);
    if (match !== null) {
      return { start: match.index, end: match.index + match[0].length };
    }
  }
  return null;
};

/**
 * Resolve a `seq-NNNN` fallback id by document position.
 *
 * A sequential id is only minted for a paragraph the source DOCX left
 * without a `w14:paraId`. The live editor's `ParaIdAllocator` fills
 * that gap with a fresh random hex paraId, so a snapshot of the live
 * document keys the block by that hex and never reproduces the
 * server's `seq-NNNN`: the direct anchor lookup misses, and a
 * paraId-based live lookup can't match either (no node carries a
 * `seq-` paraId). The seq number is the block's 1-based position in
 * the same non-empty-block walk the server extractor and
 * `createFolioAIEditSnapshot` share, so it indexes the snapshot's
 * ordered `blocks` directly.
 */
const resolveSequentialBlockAnchor = (
  blockId: string,
  snapshot: FolioAIEditSnapshot,
): FolioAIBlockAnchor | undefined => {
  const index = getSequentialFolioBlockIdIndex(blockId);
  if (index === null) {
    return undefined;
  }
  const block = snapshot.blocks.at(index - 1);
  return block ? snapshot.anchors[block.id] : undefined;
};

/**
 * Clamp a `{from, to}` pair so both endpoints fit inside a document of
 * `docSize` (in PM content positions). Block-boundary snapshots and stale
 * range data sometimes produce a `to` one past the last inline position;
 * `view.state.doc.resolve(...)` rejects that with
 * "Position … out of range", and `TextSelection.between` doesn't help — it
 * needs *valid* resolved positions. Clamping before resolution is the cheap
 * defensive step.
 *
 * Order is preserved: if both endpoints exceed `docSize`, the returned
 * `from` may equal `to`, yielding a cursor selection at the doc end.
 */
export function clampRangeToDocSize(docSize: number, range: DocPositionRange): DocPositionRange {
  return {
    from: Math.min(Math.max(range.from, 0), docSize),
    to: Math.min(Math.max(range.to, 0), docSize),
  };
}
