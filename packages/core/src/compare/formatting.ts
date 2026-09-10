/**
 * Character-aligned inline-formatting diff of two text-equal blocks.
 *
 * Owned here and consumed by both the redline generator and
 * {@link ./compare.compareDocx}, so a formatting-only difference is described
 * the same way in the generated tracked changes and in the change list.
 */

import type {
  FolioContentBlock,
  FolioContentInlineFormattingPatch,
  FolioContentRun,
} from "./content-types";

const HEX_COLOR = /^#?(?:[0-9A-Fa-f]{3}|[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/u;

const normalizeInlineFormattingColor = (color: string): string =>
  HEX_COLOR.test(color)
    ? color.replace(/^#/u, "").toUpperCase()
    : color;

const normalizeEffectiveInlineFormattingColor = (
  color: string | undefined,
): string | undefined =>
  color === undefined ? undefined : normalizeInlineFormattingColor(color);

const normalizeDirectInlineFormattingColor = (
  color: string | null | undefined,
): string | null | undefined =>
  color === null || color === undefined ? color : normalizeInlineFormattingColor(color);

const changedStringValue = (
  baseEffective: string | undefined,
  targetEffective: string | undefined,
  baseAuthored: string | null | undefined,
  targetAuthored: string | null | undefined,
): string | null | undefined => {
  if (baseAuthored !== targetAuthored) {
    return targetAuthored ?? null;
  }
  return baseEffective === targetEffective ? undefined : (targetEffective ?? null);
};

const changedNumberValue = (
  baseEffective: number | undefined,
  targetEffective: number | undefined,
  baseAuthored: number | null | undefined,
  targetAuthored: number | null | undefined,
): number | null | undefined => {
  if (baseAuthored !== targetAuthored) {
    return targetAuthored ?? null;
  }
  return baseEffective === targetEffective ? undefined : (targetEffective ?? null);
};

/** One run of characters whose supported inline formatting differs. */
export type InlineFormattingSegment = {
  /** Zero-based UTF-16 offset into the block's visible text. */
  startOffset: number;
  endOffset: number;
  /** Differing properties set to the target value; null removes a direct property. */
  formatting: FolioContentInlineFormattingPatch;
};

const changedSupportedFormatting = (
  base: FolioContentRun,
  target: FolioContentRun,
): FolioContentInlineFormattingPatch => {
  const baseDirect = base.directFormatting ?? {};
  const targetDirect = target.directFormatting ?? {};
  const formatting: FolioContentInlineFormattingPatch = {};

  const changedBoolean = (property: "bold" | "italic" | "underline" | "strike") => {
    if (baseDirect[property] !== targetDirect[property]) {
      return targetDirect[property] ?? null;
    }
    return Boolean(base[property]) === Boolean(target[property])
      ? undefined
      : Boolean(target[property]);
  };
  for (const property of ["bold", "italic", "underline", "strike"] as const) {
    const changed = changedBoolean(property);
    if (changed !== undefined) {
      formatting[property] = changed;
    }
  }

  const fontFamily = changedStringValue(
    base.fontFamily,
    target.fontFamily,
    baseDirect.fontFamily,
    targetDirect.fontFamily,
  );
  if (fontFamily !== undefined) {
    formatting.fontFamily = fontFamily;
  }

  const fontSizePt = changedNumberValue(
    base.fontSizePt,
    target.fontSizePt,
    baseDirect.fontSizePt,
    targetDirect.fontSizePt,
  );
  if (fontSizePt !== undefined) {
    formatting.fontSizePt = fontSizePt;
  }

  const color = changedStringValue(
    normalizeEffectiveInlineFormattingColor(base.color),
    normalizeEffectiveInlineFormattingColor(target.color),
    normalizeDirectInlineFormattingColor(baseDirect.color),
    normalizeDirectInlineFormattingColor(targetDirect.color),
  );
  if (color !== undefined) {
    formatting.color = color;
  }

  return formatting;
};

const sameInlineFormatting = (
  left: FolioContentInlineFormattingPatch,
  right: FolioContentInlineFormattingPatch,
): boolean =>
  left.bold === right.bold &&
  left.italic === right.italic &&
  left.underline === right.underline &&
  left.strike === right.strike &&
  left.fontFamily === right.fontFamily &&
  left.fontSizePt === right.fontSizePt &&
  left.color === right.color;

const hasInlineFormatting = (formatting: FolioContentInlineFormattingPatch): boolean =>
  formatting.bold !== undefined ||
  formatting.italic !== undefined ||
  formatting.underline !== undefined ||
  formatting.strike !== undefined ||
  formatting.fontFamily !== undefined ||
  formatting.fontSizePt !== undefined ||
  formatting.color !== undefined;

/**
 * A block's runs, or `null` when they cannot describe the block's text.
 * Non-text inline content (a field, an image) leaves the concatenated run text
 * shorter than the block text; attributing formatting by offset would then
 * point at the wrong characters, so the caller must back off instead.
 */
const previewRunsForBlock = (block: FolioContentBlock): readonly FolioContentRun[] | null => {
  const runs = block.previewRuns ?? [{ text: block.text }];
  return runs.map(({ text }) => text).join("") === block.text ? runs : null;
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
