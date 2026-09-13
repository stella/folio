/**
 * Canonical interpretation of the inline presentation carried by comparison
 * snapshots.
 *
 * Effective and explicitly authored values are separate facts. Every caller
 * compares them through the descriptor grammar below, so adding a supported
 * property cannot update redline planning while leaving verification behind.
 */

import { panic } from "better-result";

import type {
  FolioContentBlock,
  FolioContentFormatRange,
  FolioContentInlineComparisonResult,
  FolioContentInlineFormatting,
  FolioContentInlineFormattingPatch,
  FolioContentRun,
} from "../../compare/content-types";

type InlinePresentationScalar = boolean | number | string;

const NO_PRESENTATION_DIFFERENCE = Symbol("no-inline-presentation-difference");

const HEX_COLOR = /^#?(?:[0-9A-Fa-f]{3}|[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/u;

const normalizeInlinePresentationColor = (color: string): string =>
  HEX_COLOR.test(color) ? color.replace(/^#/u, "").toUpperCase() : color;

type InlineFormattingProperty = keyof FolioContentInlineFormatting;
type InlineRunFormattingProperty = Exclude<keyof FolioContentRun, "directFormatting" | "text">;
type InlineFormattingPropertyOf<Value extends InlinePresentationScalar> = {
  [Property in InlineFormattingProperty]-?: Exclude<
    FolioContentInlineFormatting[Property],
    null | undefined
  > extends Value
    ? Property
    : never;
}[InlineFormattingProperty];

type InlineBooleanProperty = InlineFormattingPropertyOf<boolean>;
type InlineNumberProperty = InlineFormattingPropertyOf<number>;
type InlineStringProperty = Exclude<InlineFormattingPropertyOf<string>, "color">;
type InlineColorProperty = Extract<InlineFormattingProperty, "color">;

type SamePropertySet<Left, Right> =
  Exclude<Left, Right> extends never
    ? Exclude<Right, Left> extends never
      ? unknown
      : never
    : never;

/**
 * The one grammar for every modeled run-presentation property. The
 * `SamePropertySet` constraint also rejects a field added to `FolioContentRun`
 * without a corresponding authored-formatting representation, or vice versa.
 */
const INLINE_PRESENTATION_GRAMMAR = {
  boolean: ["bold", "italic", "underline", "strike"],
  string: ["fontFamily"],
  number: ["fontSizePt"],
  color: ["color"],
} as const;

/**
 * Boolean properties share one loop. The remaining semantic families each
 * have one optimized slot; adding another property makes this contract fail
 * until the hot path below gains an explicit slot for it.
 */
type GrammarProperty =
  (typeof INLINE_PRESENTATION_GRAMMAR)[keyof typeof INLINE_PRESENTATION_GRAMMAR][number];

const INLINE_PRESENTATION_HOT_PATH_GRAMMAR = INLINE_PRESENTATION_GRAMMAR satisfies Readonly<{
  boolean: readonly InlineBooleanProperty[];
  string: readonly [InlineStringProperty];
  number: readonly [InlineNumberProperty];
  color: readonly [InlineColorProperty];
}> &
  SamePropertySet<InlineFormattingProperty, GrammarProperty> &
  SamePropertySet<InlineFormattingProperty, InlineRunFormattingProperty>;

/** Derived from the descriptor grammar; tests use it to prove full coverage. */
export const CANONICAL_INLINE_PRESENTATION_PROPERTIES = Object.freeze([
  ...INLINE_PRESENTATION_GRAMMAR.boolean,
  ...INLINE_PRESENTATION_GRAMMAR.string,
  ...INLINE_PRESENTATION_GRAMMAR.number,
  ...INLINE_PRESENTATION_GRAMMAR.color,
] satisfies InlineFormattingProperty[]);

const canonicalColor = (value: string | null | undefined): string | null | undefined =>
  value === null || value === undefined ? value : normalizeInlinePresentationColor(value);

const writeCanonicalInlinePresentationDifferences = (
  base: FolioContentRun | undefined,
  revised: FolioContentRun | undefined,
  patch: FolioContentInlineFormattingPatch | null,
): boolean => {
  let changed = false;
  const baseDirect = base?.directFormatting ?? {};
  const revisedDirect = revised?.directFormatting ?? {};
  for (const property of INLINE_PRESENTATION_HOT_PATH_GRAMMAR.boolean) {
    const baseAuthored = baseDirect[property];
    const revisedAuthored = revisedDirect[property];
    const baseEffective = base?.[property] === true;
    const revisedEffective = revised?.[property] === true;
    let value;
    if (baseAuthored !== revisedAuthored) {
      value = revisedAuthored ?? null;
    } else if (baseEffective === revisedEffective) {
      value = NO_PRESENTATION_DIFFERENCE;
    } else {
      value = revisedEffective;
    }
    if (value !== NO_PRESENTATION_DIFFERENCE) {
      changed = true;
      if (patch !== null) patch[property] = value;
    }
  }
  const stringProperty = INLINE_PRESENTATION_HOT_PATH_GRAMMAR.string[0];
  let stringValue;
  if (baseDirect[stringProperty] !== revisedDirect[stringProperty]) {
    stringValue = revisedDirect[stringProperty] ?? null;
  } else if (base?.[stringProperty] === revised?.[stringProperty]) {
    stringValue = NO_PRESENTATION_DIFFERENCE;
  } else {
    stringValue = revised?.[stringProperty] ?? null;
  }
  if (stringValue !== NO_PRESENTATION_DIFFERENCE) {
    changed = true;
    if (patch !== null) patch[stringProperty] = stringValue;
  }
  const numberProperty = INLINE_PRESENTATION_HOT_PATH_GRAMMAR.number[0];
  let numberValue;
  if (baseDirect[numberProperty] !== revisedDirect[numberProperty]) {
    numberValue = revisedDirect[numberProperty] ?? null;
  } else if (base?.[numberProperty] === revised?.[numberProperty]) {
    numberValue = NO_PRESENTATION_DIFFERENCE;
  } else {
    numberValue = revised?.[numberProperty] ?? null;
  }
  if (numberValue !== NO_PRESENTATION_DIFFERENCE) {
    changed = true;
    if (patch !== null) patch[numberProperty] = numberValue;
  }
  const colorProperty = INLINE_PRESENTATION_HOT_PATH_GRAMMAR.color[0];
  const baseAuthoredColor = canonicalColor(baseDirect[colorProperty]);
  const revisedAuthoredColor = canonicalColor(revisedDirect[colorProperty]);
  const baseEffectiveColor = canonicalColor(base?.[colorProperty]);
  const revisedEffectiveColor = canonicalColor(revised?.[colorProperty]);
  let color;
  if (baseAuthoredColor !== revisedAuthoredColor) {
    color = revisedAuthoredColor ?? null;
  } else if (baseEffectiveColor === revisedEffectiveColor) {
    color = NO_PRESENTATION_DIFFERENCE;
  } else {
    color = revisedEffectiveColor ?? null;
  }
  if (color !== NO_PRESENTATION_DIFFERENCE) {
    changed = true;
    if (patch !== null) patch[colorProperty] = color;
  }
  return changed;
};

const canonicalInlinePresentationsEqual = (
  base: FolioContentRun | undefined,
  revised: FolioContentRun | undefined,
): boolean => !writeCanonicalInlinePresentationDifferences(base, revised, null);

const changedCanonicalInlinePresentation = (
  base: FolioContentRun | undefined,
  revised: FolioContentRun | undefined,
): FolioContentInlineFormattingPatch | null => {
  const patch: FolioContentInlineFormattingPatch = {};
  return writeCanonicalInlinePresentationDifferences(base, revised, patch) ? patch : null;
};

const sameCanonicalInlinePresentationPatch = (
  left: FolioContentInlineFormattingPatch,
  right: FolioContentInlineFormattingPatch,
): boolean => {
  for (const property of INLINE_PRESENTATION_GRAMMAR.boolean) {
    if (left[property] !== right[property]) return false;
  }
  for (const property of INLINE_PRESENTATION_GRAMMAR.string) {
    if (left[property] !== right[property]) return false;
  }
  for (const property of INLINE_PRESENTATION_GRAMMAR.number) {
    if (left[property] !== right[property]) return false;
  }
  for (const property of INLINE_PRESENTATION_GRAMMAR.color) {
    if (left[property] !== right[property]) return false;
  }
  return true;
};

type InlinePresentationBlock = Pick<FolioContentBlock, "previewRuns" | "text">;

type AlignedRunRangeResult = "budget-exceeded" | "complete" | "stopped" | "unalignable";

type VisitAlignedRunRangesOptions =
  | {
      type: "segments";
      baseBlock: InlinePresentationBlock;
      revisedBlock: InlinePresentationBlock;
      maxSegments: number;
      segments: FolioContentFormatRange[];
    }
  | {
      type: "equivalence";
      baseBlock: InlinePresentationBlock;
      revisedBlock: InlinePresentationBlock;
    };

/**
 * Walk the shared UTF-16 coordinate space once. Run boundaries may differ;
 * malformed run streams never yield a trusted projection.
 */
const visitAlignedRunRanges = (options: VisitAlignedRunRangesOptions): AlignedRunRangeResult => {
  const { baseBlock, revisedBlock } = options;
  if (baseBlock.text !== revisedBlock.text) {
    return "unalignable";
  }

  const baseRuns = baseBlock.previewRuns ?? [{ text: baseBlock.text }];
  const revisedRuns = revisedBlock.previewRuns ?? [{ text: revisedBlock.text }];
  let baseRunIndex = 0;
  let revisedRunIndex = 0;
  let baseRunOffset = 0;
  let revisedRunOffset = 0;
  let baseBlockOffset = 0;
  let revisedBlockOffset = 0;
  let exceededBudget = false;

  while (true) {
    let baseRunText = baseRuns[baseRunIndex]?.text;
    while (baseRunText !== undefined && baseRunOffset === baseRunText.length) {
      baseRunIndex++;
      baseRunOffset = 0;
      baseRunText = baseRuns[baseRunIndex]?.text;
    }
    let revisedRunText = revisedRuns[revisedRunIndex]?.text;
    while (revisedRunText !== undefined && revisedRunOffset === revisedRunText.length) {
      revisedRunIndex++;
      revisedRunOffset = 0;
      revisedRunText = revisedRuns[revisedRunIndex]?.text;
    }

    if (baseRunIndex === baseRuns.length || revisedRunIndex === revisedRuns.length) {
      const fullyConsumed =
        baseRunIndex === baseRuns.length &&
        revisedRunIndex === revisedRuns.length &&
        baseBlockOffset === baseBlock.text.length &&
        revisedBlockOffset === revisedBlock.text.length;
      if (!fullyConsumed) return "unalignable";
      if (exceededBudget) return "budget-exceeded";
      return "complete";
    }
    if (baseRunText === undefined || revisedRunText === undefined) {
      return "unalignable";
    }
    if (
      (baseRunOffset === 0 && !baseBlock.text.startsWith(baseRunText, baseBlockOffset)) ||
      (revisedRunOffset === 0 && !revisedBlock.text.startsWith(revisedRunText, revisedBlockOffset))
    ) {
      return "unalignable";
    }
    const length = Math.min(
      baseRunText.length - baseRunOffset,
      revisedRunText.length - revisedRunOffset,
    );
    const baseRun = baseRuns[baseRunIndex];
    const revisedRun = revisedRuns[revisedRunIndex];
    if (options.type === "equivalence") {
      if (!canonicalInlinePresentationsEqual(baseRun, revisedRun)) {
        return "stopped";
      }
    } else if (!exceededBudget) {
      const formatting = changedCanonicalInlinePresentation(baseRun, revisedRun);
      if (formatting !== null) {
        const previous = options.segments.at(-1);
        if (
          previous &&
          previous.endOffset === baseBlockOffset &&
          sameCanonicalInlinePresentationPatch(previous.formatting, formatting)
        ) {
          previous.endOffset += length;
        } else if (options.segments.length >= options.maxSegments) {
          exceededBudget = true;
        } else {
          options.segments.push({
            startOffset: baseBlockOffset,
            endOffset: baseBlockOffset + length,
            formatting,
          });
        }
      }
    }

    baseRunOffset += length;
    revisedRunOffset += length;
    baseBlockOffset += length;
    revisedBlockOffset += length;
  }
};

const hasAlignedRunStream = (block: InlinePresentationBlock): boolean => {
  if (block.previewRuns === undefined) return true;
  let offset = 0;
  for (const run of block.previewRuns) {
    if (!block.text.startsWith(run.text, offset)) return false;
    offset += run.text.length;
  }
  return offset === block.text.length;
};

type CanonicalInlinePresentationSegmentsOptions = {
  baseBlock: InlinePresentationBlock;
  revisedBlock: InlinePresentationBlock;
  maxSegments: number;
};

/**
 * Collect coalesced canonical presentation differences in character order.
 * Invalid run streams take precedence over the segment budget, so malformed
 * content cannot change meaning as the caller's remaining budget changes.
 */
export const canonicalInlinePresentationSegments = ({
  baseBlock,
  revisedBlock,
  maxSegments,
}: CanonicalInlinePresentationSegmentsOptions): FolioContentInlineComparisonResult => {
  const segments: FolioContentFormatRange[] = [];
  const result = visitAlignedRunRanges({
    type: "segments",
    baseBlock,
    revisedBlock,
    maxSegments,
    segments,
  });
  switch (result) {
    case "complete":
      return { status: "compared", segments };
    case "budget-exceeded":
      return { status: "budget-exceeded", maximum: maxSegments };
    case "unalignable": {
      // Only refusal needs a full per-side diagnosis; successful walks stay single-pass.
      const baseAligned = hasAlignedRunStream(baseBlock);
      const revisedAligned = hasAlignedRunStream(revisedBlock);
      if (baseAligned === revisedAligned) return { status: "unalignable", side: "both" };
      if (baseAligned) return { status: "unalignable", side: "revised" };
      return { status: "unalignable", side: "base" };
    }
    case "stopped":
      return panic("A segment traversal cannot stop at a presentation difference");
    default: {
      const exhaustive: never = result;
      return exhaustive;
    }
  }
};

/** Effective and authored presentation equality, independent of run splits. */
export const sameCanonicalInlinePresentation = (
  baseBlock: InlinePresentationBlock,
  revisedBlock: InlinePresentationBlock,
): boolean => {
  const result = visitAlignedRunRanges({
    type: "equivalence",
    baseBlock,
    revisedBlock,
  });
  return result === "complete";
};
