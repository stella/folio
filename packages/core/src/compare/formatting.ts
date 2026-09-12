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
  sameFolioContentPropertyChanges,
} from "./content-properties";
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

export type InlineFormattingPairedRange = {
  baseStart: number;
  baseEnd: number;
  revisedStart: number;
  revisedEnd: number;
};

export type PairedInlineFormattingSegment = InlineFormattingPairedRange & {
  formatting: FolioContentInlineFormattingChange;
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
