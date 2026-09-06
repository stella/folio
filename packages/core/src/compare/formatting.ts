/**
 * Character-aligned inline-formatting diff of two text-equal blocks.
 *
 * Owned here and consumed by both the redline generator and
 * {@link ./compare.compareDocx}, so a formatting-only difference is described
 * the same way in the generated tracked changes and in the change list.
 */

import type {
  FolioAIBlock,
  FolioAIBlockPreviewRun,
  FolioAIInlineFormatting,
} from "../ai-edits/types";

/** One run of characters whose supported inline formatting differs. */
export type InlineFormattingSegment = {
  /** Zero-based UTF-16 offset into the block's visible text. */
  startOffset: number;
  endOffset: number;
  /** Only the properties that differ, set to the target document's value. */
  formatting: FolioAIInlineFormatting;
};

const changedSupportedFormatting = (
  base: FolioAIBlockPreviewRun,
  target: FolioAIBlockPreviewRun,
): FolioAIInlineFormatting => ({
  ...(Boolean(base.bold) !== Boolean(target.bold) && { bold: Boolean(target.bold) }),
  ...(Boolean(base.italic) !== Boolean(target.italic) && { italic: Boolean(target.italic) }),
  ...(Boolean(base.underline) !== Boolean(target.underline) && {
    underline: Boolean(target.underline),
  }),
});

const sameInlineFormatting = (
  left: FolioAIInlineFormatting,
  right: FolioAIInlineFormatting,
): boolean =>
  left.bold === right.bold && left.italic === right.italic && left.underline === right.underline;

const hasInlineFormatting = (formatting: FolioAIInlineFormatting): boolean =>
  formatting.bold !== undefined ||
  formatting.italic !== undefined ||
  formatting.underline !== undefined;

/**
 * A block's runs, or `null` when they cannot describe the block's text.
 * Non-text inline content (a field, an image) leaves the concatenated run text
 * shorter than the block text; attributing formatting by offset would then
 * point at the wrong characters, so the caller must back off instead.
 */
const previewRunsForBlock = (block: FolioAIBlock): readonly FolioAIBlockPreviewRun[] | null => {
  const runs = block.previewRuns ?? [{ text: block.text }];
  return runs.map(({ text }) => text).join("") === block.text ? runs : null;
};

type InlineFormattingSegmentsOptions = {
  baseBlock: FolioAIBlock;
  targetBlock: FolioAIBlock;
  /** Refuse (return `null`) rather than build more segments than this. */
  maxSegments: number;
};

/**
 * Segments where `targetBlock`'s bold / italic / underline differs from
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
  const baseRuns = previewRunsForBlock(baseBlock);
  const targetRuns = previewRunsForBlock(targetBlock);
  if (!baseRuns || !targetRuns || baseBlock.text.length === 0) {
    return [];
  }

  const segments: InlineFormattingSegment[] = [];
  let baseRunIndex = 0;
  let targetRunIndex = 0;
  let baseRunOffset = 0;
  let targetRunOffset = 0;
  let textOffset = 0;

  while (baseRunIndex < baseRuns.length && targetRunIndex < targetRuns.length) {
    const baseRun = baseRuns[baseRunIndex];
    const targetRun = targetRuns[targetRunIndex];
    if (!baseRun || !targetRun) {
      break;
    }
    const length = Math.min(
      baseRun.text.length - baseRunOffset,
      targetRun.text.length - targetRunOffset,
    );
    if (length <= 0) {
      if (baseRunOffset >= baseRun.text.length) {
        baseRunIndex++;
        baseRunOffset = 0;
      }
      if (targetRunOffset >= targetRun.text.length) {
        targetRunIndex++;
        targetRunOffset = 0;
      }
      continue;
    }

    const formatting = changedSupportedFormatting(baseRun, targetRun);
    if (hasInlineFormatting(formatting)) {
      const previous = segments.at(-1);
      if (
        previous &&
        previous.endOffset === textOffset &&
        sameInlineFormatting(previous.formatting, formatting)
      ) {
        previous.endOffset += length;
      } else {
        if (segments.length >= maxSegments) {
          return null;
        }
        segments.push({
          startOffset: textOffset,
          endOffset: textOffset + length,
          formatting,
        });
      }
    }

    textOffset += length;
    baseRunOffset += length;
    targetRunOffset += length;
    if (baseRunOffset >= baseRun.text.length) {
      baseRunIndex++;
      baseRunOffset = 0;
    }
    if (targetRunOffset >= targetRun.text.length) {
      targetRunIndex++;
      targetRunOffset = 0;
    }
  }

  return segments;
};
