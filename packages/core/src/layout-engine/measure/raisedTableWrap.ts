import { measuredLineRangeHeight } from "../lineFlow";
import type { FlowBlock, Layout, Measure, ParagraphMeasure } from "../types";
import { isUndecoratedParagraphMark } from "./measureBlocks";
import { MIN_WRAP_SEGMENT_WIDTH } from "./measureParagraph";
import type { FloatingImageZone } from "./measureParagraph";

export type RaisedTableWrapInputs = ReadonlyMap<
  FlowBlock["id"],
  readonly FloatingImageZone[]
>;

const MAX_PLACEMENT_PASSES = 4;
const MAX_PAGE_OVERLAP_COMPARISONS = 16_384;
const settledParagraphs = new WeakSet<ParagraphMeasure>();

export const hasSettledRaisedTableWrap = (measure: ParagraphMeasure): boolean =>
  settledParagraphs.has(measure);

const hasRaisedTableWrap = (block: FlowBlock): boolean => {
  if (block.kind !== "table" || block.floating?.vertAnchor !== "text") {
    return false;
  }
  const { tblpY, tblpYSpec, topFromText } = block.floating;
  return (tblpYSpec === undefined ? (tblpY ?? 0) : 0) < (topFromText ?? 0);
};

const collectInputs = (
  layout: Layout,
  entries: ReadonlyMap<FlowBlock["id"], { block: FlowBlock; measure: Measure }>,
): RaisedTableWrapInputs | undefined => {
  const inputs = new Map<FlowBlock["id"], FloatingImageZone[]>();
  for (const page of layout.pages) {
    const earlierParagraphs = [];
    let remainingComparisons = MAX_PAGE_OVERLAP_COMPARISONS;
    for (const fragment of page.fragments) {
      if (fragment.kind === "paragraph") {
        const entry = entries.get(fragment.blockId);
        if (
          fragment.height > 0 &&
          entry?.measure.kind === "paragraph" &&
          !isUndecoratedParagraphMark(entry.block)
        ) {
          earlierParagraphs.push({
            paragraph: fragment,
            measure: entry.measure,
          });
        }
        continue;
      }
      if (fragment.kind !== "table") continue;
      const source = entries.get(fragment.blockId)?.block;
      if (source?.kind !== "table" || !hasRaisedTableWrap(source)) continue;
      const floating = source.floating;
      const left = fragment.x - (floating?.leftFromText ?? 12);
      const right =
        fragment.x + fragment.width + (floating?.rightFromText ?? 12);
      const top = fragment.y - (floating?.topFromText ?? 0);
      const bottom =
        fragment.y + fragment.height + (floating?.bottomFromText ?? 0);
      remainingComparisons -= earlierParagraphs.length;
      if (remainingComparisons < 0) return undefined;
      for (const { paragraph, measure } of earlierParagraphs) {
        const paragraphBottom = paragraph.y + paragraph.height;
        const paragraphRight = paragraph.x + paragraph.width;
        if (
          paragraphBottom <= top ||
          paragraph.y >= bottom ||
          paragraphRight <= left ||
          paragraph.x >= right
        ) {
          continue;
        }
        const wrapsOnLeft = left >= paragraph.x + paragraph.width / 2;
        const leftMargin = wrapsOnLeft ? 0 : right - paragraph.x;
        const rightMargin = wrapsOnLeft ? paragraphRight - left : 0;
        if (
          paragraph.width - leftMargin - rightMargin <
          MIN_WRAP_SEGMENT_WIDTH
        ) {
          continue;
        }
        const paragraphStart =
          paragraph.y -
          measuredLineRangeHeight(measure.lines, 0, paragraph.fromLine);
        const zones = inputs.get(paragraph.blockId) ?? [];
        zones.push({
          source: "floatingTable",
          leftMargin,
          rightMargin,
          topY: Math.max(top, paragraph.y) - paragraphStart,
          bottomY: Math.min(bottom, paragraphBottom) - paragraphStart,
        });
        inputs.set(paragraph.blockId, zones);
      }
    }
  }
  return inputs;
};

const sameInputs = (
  before: RaisedTableWrapInputs,
  after: RaisedTableWrapInputs,
) =>
  before.size === after.size &&
  [...before].every(([id, zones]) => {
    const next = after.get(id);
    return (
      next !== undefined &&
      zones.length === next.length &&
      zones.every((zone, index) => {
        const candidate = next[index];
        return (
          candidate !== undefined &&
          zone.topY === candidate.topY &&
          zone.bottomY === candidate.bottomY &&
          zone.leftMargin === candidate.leftMargin &&
          zone.rightMargin === candidate.rightMargin
        );
      })
    );
  });

type SettleRaisedTableWrapOptions = {
  blocks: readonly FlowBlock[];
  measures: Measure[];
  layout: (measures: Measure[]) => Layout;
  measure: (inputs: RaisedTableWrapInputs) => Measure[];
};

/** Find a layout whose earlier wraps agree with its actual floating table placements. */
export function settleRaisedTableWrap({
  blocks,
  measures,
  layout,
  measure,
}: SettleRaisedTableWrapOptions) {
  const original = { measures, layout: layout(measures) };
  if (!blocks.some(hasRaisedTableWrap)) return original;

  const collect = (current: typeof original) => {
    const entries = new Map<
      FlowBlock["id"],
      { block: FlowBlock; measure: Measure }
    >();
    blocks.forEach((block, index) => {
      const blockMeasure = current.measures[index];
      if (blockMeasure !== undefined) {
        entries.set(block.id, { block, measure: blockMeasure });
      }
    });
    return collectInputs(current.layout, entries);
  };

  let inputs = collect(original);
  for (
    let pass = 0;
    pass < MAX_PLACEMENT_PASSES && inputs !== undefined && inputs.size > 0;
    pass += 1
  ) {
    const measured = measure(inputs);
    const current = { measures: measured, layout: layout(measured) };
    const nextInputs = collect(current);
    if (nextInputs === undefined) return original;
    if (sameInputs(inputs, nextInputs)) {
      for (const [index, block] of blocks.entries()) {
        const paragraph = measured[index];
        if (inputs.has(block.id) && paragraph?.kind === "paragraph") {
          settledParagraphs.add(paragraph);
        }
      }
      return current;
    }
    inputs = nextInputs;
  }

  // An exclusion can move its own anchor to a different page. If no stable
  // arrangement is found, keep the original matched layout and measures.
  return original;
}
