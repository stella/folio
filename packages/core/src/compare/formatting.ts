/**
 * Character-aligned inline-formatting diff of two text-equal blocks.
 *
 * Owned here and consumed by both the redline generator and
 * {@link ./compare.compareDocx}, so a formatting-only difference is described
 * the same way in the generated tracked changes and in the change list.
 */

import type {
  FolioContentBlock,
  FolioContentInlineFormattingChange,
  FolioContentRun,
} from "./content-types";
import {
  changedFolioContentProperties,
  sameFolioContentPropertySet,
  sameFolioContentPropertyChanges,
} from "./content-properties";
import type { WordDiffSegment } from "./text-diff";
import { panic } from "better-result";

/** One run of characters whose supported inline formatting differs. */
export type InlineFormattingSegment = {
  /** Zero-based UTF-16 offset into the block's visible text. */
  startOffset: number;
  endOffset: number;
  formatting: FolioContentInlineFormattingChange;
};

const changedSupportedFormatting = (
  base: FolioContentRun,
  target: FolioContentRun,
): FolioContentInlineFormattingChange =>
  Object.freeze({
    authored: changedFolioContentProperties(base.authoredFormatting, target.authoredFormatting),
    effective: changedFolioContentProperties(base.effectiveFormatting, target.effectiveFormatting),
  });

const sameInlineFormatting = (
  left: FolioContentInlineFormattingChange,
  right: FolioContentInlineFormattingChange,
): boolean =>
  sameFolioContentPropertyChanges(left.authored, right.authored) &&
  sameFolioContentPropertyChanges(left.effective, right.effective);

const hasInlineFormatting = (formatting: FolioContentInlineFormattingChange): boolean =>
  formatting.authored.length > 0 || formatting.effective.length > 0;

/**
 * A block's runs, or `null` when they cannot describe the block's text.
 * Non-text inline content (a field, an image) leaves the concatenated run text
 * shorter than the block text; attributing formatting by offset would then
 * point at the wrong characters, so the caller must back off instead.
 */
const runsForBlock = (block: FolioContentBlock): readonly FolioContentRun[] | null => {
  const runs =
    block.runs.length > 0
      ? block.runs
      : [
          Object.freeze({
            text: block.text,
            authoredFormatting: Object.freeze([]),
            effectiveFormatting: Object.freeze([]),
          }),
        ];
  let offset = 0;
  for (const run of runs) {
    if (!block.text.startsWith(run.text, offset)) return null;
    offset += run.text.length;
  }
  return offset === block.text.length ? runs : null;
};

type RunCursor = {
  runs: readonly FolioContentRun[];
  index: number;
  runStart: number;
};

const runAtOffset = (cursor: RunCursor, offset: number): FolioContentRun | null => {
  while (cursor.index < cursor.runs.length) {
    const run = cursor.runs[cursor.index];
    if (!run) return null;
    const runEnd = cursor.runStart + run.text.length;
    if (offset < runEnd) return run;
    cursor.index++;
    cursor.runStart = runEnd;
  }
  return null;
};

type SameFormattingRangeOptions = {
  base: RunCursor;
  revised: RunCursor;
  baseStart: number;
  revisedStart: number;
  length: number;
};

/** Exact authored and effective formatting equality over two text-equal ranges. */
const sameFormattingRange = ({
  base,
  revised,
  baseStart,
  revisedStart,
  length,
}: SameFormattingRangeOptions): boolean => {
  let baseOffset = baseStart;
  let revisedOffset = revisedStart;
  let remaining = length;
  while (remaining > 0) {
    const baseRun = runAtOffset(base, baseOffset);
    const revisedRun = runAtOffset(revised, revisedOffset);
    if (!baseRun || !revisedRun) return false;
    if (
      !sameFolioContentPropertySet(baseRun.authoredFormatting, revisedRun.authoredFormatting) ||
      !sameFolioContentPropertySet(baseRun.effectiveFormatting, revisedRun.effectiveFormatting)
    ) {
      return false;
    }
    const compared = Math.min(
      remaining,
      base.runStart + baseRun.text.length - baseOffset,
      revised.runStart + revisedRun.text.length - revisedOffset,
    );
    if (compared <= 0) return false;
    remaining -= compared;
    baseOffset += compared;
    revisedOffset += compared;
  }
  return true;
};

type PositionedSegment = {
  baseStart: number;
  revisedStart: number;
};

const positionSegments = (
  segments: readonly WordDiffSegment[],
  baseStart: number,
  revisedStart: number,
): readonly PositionedSegment[] => {
  const positioned: PositionedSegment[] = [];
  let baseOffset = baseStart;
  let revisedOffset = revisedStart;
  for (const segment of segments) {
    positioned.push({ baseStart: baseOffset, revisedStart: revisedOffset });
    if (segment.type !== "ins") baseOffset += segment.text.length;
    if (segment.type !== "del") revisedOffset += segment.text.length;
  }
  return positioned;
};

const leadingWhitespace = (text: string): string => /^\s+/u.exec(text)?.[0] ?? "";

type BoundaryRotation = {
  changeStart: number;
  equalIndex: number;
  whitespace: string;
};

type BoundaryCandidate = BoundaryRotation & {
  currentBaseStart: number;
  currentRevisedStart: number;
  alternativeBaseStart: number;
  alternativeRevisedStart: number;
};

const pushWordSegment = (segments: WordDiffSegment[], segment: WordDiffSegment): void => {
  if (segment.text.length === 0) return;
  const previous = segments.at(-1);
  if (previous?.type === segment.type) {
    previous.text += segment.text;
    return;
  }
  segments.push(segment);
};

type AlignBoundaryWhitespaceOptions = {
  baseBlock: FolioContentBlock;
  revisedBlock: FolioContentBlock;
  baseStart: number;
  revisedStart: number;
  segments: readonly WordDiffSegment[];
};

/**
 * Resolve the ownership of an unchanged separator beside an edit.
 *
 * Word tokens carry leading whitespace, so duplicate words can leave the
 * surviving separator paired with the correct word but the wrong run. There
 * is an equally short alignment which keeps that separator with the preceding
 * unchanged text. Choose it only when its authored and effective formatting
 * are exactly equal and the current pairing is not; text-only ties retain the
 * established deterministic LCS result.
 *
 * This is a linear post-pass over already bounded segments and run spans. It
 * never moves non-whitespace text or creates an opposite-direction change.
 *
 * @internal
 */
export const alignBoundaryWhitespaceToFormatting = ({
  baseBlock,
  revisedBlock,
  baseStart,
  revisedStart,
  segments,
}: AlignBoundaryWhitespaceOptions): WordDiffSegment[] => {
  if (segments.length < 2) return [...segments];
  const singleBaseRun = baseBlock.runs.length === 1 ? baseBlock.runs[0] : undefined;
  const singleRevisedRun = revisedBlock.runs.length === 1 ? revisedBlock.runs[0] : undefined;
  if (
    (baseBlock.runs.length === 0 && revisedBlock.runs.length === 0) ||
    (singleBaseRun &&
      singleRevisedRun &&
      sameFolioContentPropertySet(
        singleBaseRun.authoredFormatting,
        singleRevisedRun.authoredFormatting,
      ) &&
      sameFolioContentPropertySet(
        singleBaseRun.effectiveFormatting,
        singleRevisedRun.effectiveFormatting,
      ))
  ) {
    return [...segments];
  }

  const positioned = positionSegments(segments, baseStart, revisedStart);
  const candidates: BoundaryCandidate[] = [];
  for (let equalIndex = 1; equalIndex < segments.length; equalIndex++) {
    const equal = segments[equalIndex];
    if (equal?.type !== "equal") continue;
    const whitespace = leadingWhitespace(equal.text);
    if (whitespace.length === 0) continue;

    let changeStart = equalIndex - 1;
    while (changeStart > 0 && segments[changeStart - 1]?.type !== "equal") changeStart--;
    const changed = segments.slice(changeStart, equalIndex);
    let deletionIndex = -1;
    let insertionIndex = -1;
    let ambiguous = false;
    for (const [index, segment] of changed.entries()) {
      if (segment.type === "del") {
        if (deletionIndex !== -1) ambiguous = true;
        deletionIndex = index;
      } else if (segment.type === "ins") {
        if (insertionIndex !== -1) ambiguous = true;
        insertionIndex = index;
      }
    }
    if (ambiguous) continue;
    if (deletionIndex === -1 && insertionIndex === -1) continue;
    if (insertionIndex !== -1 && deletionIndex > insertionIndex) continue;
    const deletion = deletionIndex === -1 ? undefined : changed[deletionIndex];
    const insertion = insertionIndex === -1 ? undefined : changed[insertionIndex];
    if (
      (deletion && leadingWhitespace(deletion.text) !== whitespace) ||
      (insertion && leadingWhitespace(insertion.text) !== whitespace)
    ) {
      continue;
    }
    if (
      ![deletion, insertion].some(
        (segment) => segment !== undefined && /\S/u.test(segment.text.slice(whitespace.length)),
      )
    ) {
      continue;
    }

    const current = positioned[equalIndex];
    const deletionPosition =
      deletionIndex === -1 ? undefined : positioned[changeStart + deletionIndex];
    const insertionPosition =
      insertionIndex === -1 ? undefined : positioned[changeStart + insertionIndex];
    if (!current) continue;
    candidates.push({
      changeStart,
      equalIndex,
      whitespace,
      currentBaseStart: current.baseStart,
      currentRevisedStart: current.revisedStart,
      alternativeBaseStart: deletionPosition?.baseStart ?? current.baseStart,
      alternativeRevisedStart: insertionPosition?.revisedStart ?? current.revisedStart,
    });
  }

  if (candidates.length === 0) return [...segments];
  const baseRuns = runsForBlock(baseBlock);
  const revisedRuns = runsForBlock(revisedBlock);
  if (!baseRuns || !revisedRuns) return [...segments];
  const cursor = (runs: readonly FolioContentRun[]): RunCursor => ({
    runs,
    index: 0,
    runStart: 0,
  });
  const currentBase = cursor(baseRuns);
  const currentRevised = cursor(revisedRuns);
  const alternativeBase = cursor(baseRuns);
  const alternativeRevised = cursor(revisedRuns);
  const rotations: BoundaryRotation[] = [];
  for (const candidate of candidates) {
    const currentIsExact = sameFormattingRange({
      base: currentBase,
      revised: currentRevised,
      baseStart: candidate.currentBaseStart,
      revisedStart: candidate.currentRevisedStart,
      length: candidate.whitespace.length,
    });
    if (currentIsExact) continue;
    const alternativeIsExact = sameFormattingRange({
      base: alternativeBase,
      revised: alternativeRevised,
      baseStart: candidate.alternativeBaseStart,
      revisedStart: candidate.alternativeRevisedStart,
      length: candidate.whitespace.length,
    });
    if (alternativeIsExact) rotations.push(candidate);
  }
  if (rotations.length === 0) return [...segments];

  const aligned = segments.map((segment) => ({ ...segment }));
  for (let index = rotations.length - 1; index >= 0; index--) {
    const rotation = rotations[index];
    if (!rotation) continue;
    const changed = aligned.slice(rotation.changeStart, rotation.equalIndex);
    const equal = aligned[rotation.equalIndex];
    if (equal?.type !== "equal") return panic("A whitespace rotation lost its equal boundary");
    const replacement: WordDiffSegment[] = [{ type: "equal", text: rotation.whitespace }];
    for (const segment of changed) {
      replacement.push({
        type: segment.type,
        text: `${segment.text.slice(rotation.whitespace.length)}${rotation.whitespace}`,
      });
    }
    replacement.push({ type: "equal", text: equal.text.slice(rotation.whitespace.length) });
    aligned.splice(
      rotation.changeStart,
      rotation.equalIndex - rotation.changeStart + 1,
      ...replacement,
    );
  }

  const coalesced: WordDiffSegment[] = [];
  for (const segment of aligned) pushWordSegment(coalesced, segment);
  return coalesced;
};

export type InlineFormattingPairedRange = {
  baseStart: number;
  baseEnd: number;
  revisedStart: number;
  revisedEnd: number;
};

export type PairedInlineFormattingSegment = InlineFormattingPairedRange & {
  formatting: FolioContentInlineFormattingChange;
};

type PairedInlineFormattingSegmentsOptions = {
  baseBlock: FolioContentBlock;
  revisedBlock: FolioContentBlock;
  equalRanges: readonly InlineFormattingPairedRange[];
  /** Refuse (return `null`) rather than build more segments than this. */
  maxSegments: number;
};

/**
 * Compare formatting over monotone, text-equal base/revised ranges in one
 * forward walk of both blocks' run streams.
 *
 * @internal
 */
export const pairedInlineFormattingSegments = ({
  baseBlock,
  revisedBlock,
  equalRanges,
  maxSegments,
}: PairedInlineFormattingSegmentsOptions): PairedInlineFormattingSegment[] | null => {
  const baseRuns = runsForBlock(baseBlock);
  const revisedRuns = runsForBlock(revisedBlock);
  if (!baseRuns || !revisedRuns) return [];

  const segments: PairedInlineFormattingSegment[] = [];
  const baseCursor: RunCursor = { runs: baseRuns, index: 0, runStart: 0 };
  const revisedCursor: RunCursor = { runs: revisedRuns, index: 0, runStart: 0 };
  let previousBaseEnd = 0;
  let previousRevisedEnd = 0;

  for (const range of equalRanges) {
    if (
      range.baseStart < 0 ||
      range.revisedStart < 0 ||
      range.baseStart < previousBaseEnd ||
      range.revisedStart < previousRevisedEnd ||
      range.baseEnd > baseBlock.text.length ||
      range.revisedEnd > revisedBlock.text.length ||
      range.baseEnd - range.baseStart !== range.revisedEnd - range.revisedStart
    ) {
      return panic("Inline-formatting ranges must be bounded, monotone, and equally sized");
    }
    previousBaseEnd = range.baseEnd;
    previousRevisedEnd = range.revisedEnd;
    let baseOffset = range.baseStart;
    let revisedOffset = range.revisedStart;
    while (baseOffset < range.baseEnd) {
      const baseRun = runAtOffset(baseCursor, baseOffset);
      const revisedRun = runAtOffset(revisedCursor, revisedOffset);
      if (!baseRun || !revisedRun) break;
      const baseRunEnd = baseCursor.runStart + baseRun.text.length;
      const revisedRunEnd = revisedCursor.runStart + revisedRun.text.length;
      const length = Math.min(
        range.baseEnd - baseOffset,
        range.revisedEnd - revisedOffset,
        baseRunEnd - baseOffset,
        revisedRunEnd - revisedOffset,
      );
      if (length <= 0) break;
      const formatting = changedSupportedFormatting(baseRun, revisedRun);
      if (hasInlineFormatting(formatting)) {
        const previous = segments.at(-1);
        if (
          previous &&
          previous.baseEnd === baseOffset &&
          previous.revisedEnd === revisedOffset &&
          sameInlineFormatting(previous.formatting, formatting)
        ) {
          previous.baseEnd += length;
          previous.revisedEnd += length;
        } else {
          if (segments.length >= maxSegments) return null;
          segments.push({
            baseStart: baseOffset,
            baseEnd: baseOffset + length,
            revisedStart: revisedOffset,
            revisedEnd: revisedOffset + length,
            formatting,
          });
        }
      }
      baseOffset += length;
      revisedOffset += length;
    }
  }
  return segments;
};

type InlineFormattingSegmentsOptions = {
  baseBlock: FolioContentBlock;
  targetBlock: FolioContentBlock;
  /** Refuse (return `null`) rather than build more segments than this. */
  maxSegments: number;
};

/**
 * Segments where `targetBlock`'s supported run formatting differs from
 * `baseBlock`'s, for two blocks that carry the same text. Returns `null` only
 * when the diff would exceed `maxSegments`.
 *
 * A block whose runs cannot be aligned to its text reports no segments rather
 * than guessing: attributing formatting to the wrong characters is worse than
 * missing a formatting-only change, and the caller has no offset it could
 * trust instead.
 */
export const inlineFormattingSegments = ({
  baseBlock,
  targetBlock,
  maxSegments,
}: InlineFormattingSegmentsOptions): InlineFormattingSegment[] | null => {
  if (baseBlock.text !== targetBlock.text || baseBlock.text.length === 0) return [];
  const paired = pairedInlineFormattingSegments({
    baseBlock,
    revisedBlock: targetBlock,
    equalRanges: [
      {
        baseStart: 0,
        baseEnd: baseBlock.text.length,
        revisedStart: 0,
        revisedEnd: targetBlock.text.length,
      },
    ],
    maxSegments,
  });
  return (
    paired?.map(({ revisedStart, revisedEnd, formatting }) => ({
      startOffset: revisedStart,
      endOffset: revisedEnd,
      formatting,
    })) ?? null
  );
};
