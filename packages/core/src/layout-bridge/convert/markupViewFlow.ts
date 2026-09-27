/**
 * Markup view — flow-block addressing stage.
 *
 * A review view other than All Markup lays out the document it reads (see
 * `prosemirror/markupViewProjection.ts`), so `toFlowBlocks` addresses its runs
 * by positions in that projection. Everything downstream of layout (caret,
 * selection, click-to-position, dirty-range remeasure, the painted
 * `data-pm-*` anchors) speaks editor positions, so this stage maps every
 * anchor back to the authored document the editor holds.
 *
 * A text run is one PM position per character. Where the view dropped text
 * inside what is now one run (an accepted deletion between two runs with the
 * same formatting merges them), the run is split at the drop, so each piece
 * keeps that invariant in editor positions: a character the page paints
 * addresses the editor character it shows. Splitting a run into pieces with
 * identical formatting does not change what the measurer sees.
 *
 * Blocks keep the positions of their own nodes; content the view hides lies
 * between two anchors, which caret placement already resolves to the nearest
 * painted boundary.
 */

import type { StepMap } from "prosemirror-transform";

import { indexedPositionMap } from "../../internal/indexedPositionMap";
import type {
  FlowBlock,
  ParagraphBlock,
  Run,
  SdtGroup,
  TableBlock,
  TableCell,
  TableRow,
  TextRun,
} from "../../layout-engine/types";

type PositionRange = { from: number; to: number };

export type MarkupViewFlowOptions = {
  /** Authored (editor) positions to the positions the blocks were built with. */
  positionMap: StepMap;
  /** Ranges, in the blocks' positions, whose paragraphs show a change bar. */
  changeBarRanges: readonly PositionRange[];
};

type ToSource = (position: number, assoc: 1 | -1) => number;

type RemapContext = {
  toSource: ToSource;
  /** Sorted view positions where the source is not contiguous. */
  cuts: readonly number[];
  changeBarRanges: readonly PositionRange[];
};

/**
 * Re-address blocks built from a view projection in editor positions, and mark
 * the paragraphs the view flags with a change bar.
 */
export function remapMarkupViewBlocks(
  blocks: readonly FlowBlock[],
  { positionMap, changeBarRanges }: MarkupViewFlowOptions,
): FlowBlock[] {
  const inverse = positionMap.invert();
  const cutSet = new Set<number>();
  inverse.forEach((oldStart, oldEnd) => {
    cutSet.add(oldStart);
    cutSet.add(oldEnd);
  });
  const context: RemapContext = {
    toSource: indexedPositionMap(inverse),
    cuts: [...cutSet].sort((left, right) => left - right),
    changeBarRanges,
  };
  return blocks.map((block) => remapBlock(block, context));
}

const touchesChangeBar = (block: ParagraphBlock, ranges: readonly PositionRange[]): boolean => {
  const { pmStart, pmEnd } = block;
  if (pmStart === undefined || pmEnd === undefined) {
    return false;
  }
  return ranges.some((range) => range.from < pmEnd && range.to > pmStart);
};

type AnchoredRange = { pmStart?: number; pmEnd?: number };

/** A node's own positions: its start after any hidden content before it, its end before any after. */
const remapNodeRange = (range: AnchoredRange, toSource: ToSource): AnchoredRange => {
  if (range.pmStart === undefined || range.pmEnd === undefined) {
    return {};
  }
  const pmStart = toSource(range.pmStart, 1);
  return { pmStart, pmEnd: Math.max(pmStart, toSource(range.pmEnd, -1)) };
};

const remapSdtGroups = (
  groups: readonly SdtGroup[] | undefined,
  toSource: ToSource,
): { sdtGroups?: SdtGroup[] } =>
  groups === undefined
    ? {}
    : { sdtGroups: groups.map((group) => ({ ...group, pmPos: toSource(group.pmPos, 1) })) };

function remapBlock(block: FlowBlock, context: RemapContext): FlowBlock {
  switch (block.kind) {
    case "paragraph":
      return remapParagraph(block, context);
    case "table":
      return remapTable(block, context);
    case "image":
    case "pageBreak":
    case "columnBreak":
      return { ...block, ...remapNodeRange(block, context.toSource) };
    case "textBox":
      return {
        ...block,
        ...remapNodeRange(block, context.toSource),
        content: block.content.map((contentBlock) =>
          contentBlock.kind === "table"
            ? remapTable(contentBlock, context)
            : remapParagraph(contentBlock, context),
        ),
      };
    case "sectionBreak":
      return block;
    default: {
      const exhaustive: never = block;
      return exhaustive;
    }
  }
}

function remapTable(block: TableBlock, context: RemapContext): TableBlock {
  return {
    ...block,
    ...remapNodeRange(block, context.toSource),
    ...remapSdtGroups(block.sdtGroups, context.toSource),
    rows: block.rows.map(
      (row): TableRow => ({
        ...row,
        cells: row.cells.map(
          (cell): TableCell => ({
            ...cell,
            blocks: cell.blocks.map((cellBlock) => remapBlock(cellBlock, context)),
          }),
        ),
      }),
    ),
  };
}

function remapParagraph(block: ParagraphBlock, context: RemapContext): ParagraphBlock {
  const runs: Run[] = [];
  for (const run of block.runs) {
    pushRemappedRun(run, context, runs);
  }
  return {
    ...block,
    ...remapNodeRange(block, context.toSource),
    ...remapSdtGroups(block.sdtGroups, context.toSource),
    ...(touchesChangeBar(block, context.changeBarRanges)
      ? { reviewIndicator: "change-bar" as const }
      : {}),
    runs,
  };
}

/** Index of the first cut strictly after `position`. */
const firstCutAfter = (cuts: readonly number[], position: number): number => {
  let low = 0;
  let high = cuts.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    const cut = cuts.at(middle);
    if (cut !== undefined && cut <= position) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return low;
};

/** Whether each character of the run occupies exactly one position. */
const addressesPerCharacter = (run: TextRun): run is TextRun & { pmStart: number; pmEnd: number } =>
  run.pmStart !== undefined &&
  run.pmEnd !== undefined &&
  run.templatePreview === undefined &&
  run.text.length === run.pmEnd - run.pmStart;

function pushRemappedRun(run: Run, context: RemapContext, out: Run[]): void {
  if (run.kind !== "text" || !addressesPerCharacter(run)) {
    out.push({ ...run, ...remapNodeRange(run, context.toSource) });
    return;
  }
  const { pmStart, pmEnd } = run;
  let pieceStart = pmStart;
  for (let index = firstCutAfter(context.cuts, pmStart); ; index++) {
    const cut = context.cuts.at(index);
    const pieceEnd = cut === undefined || cut >= pmEnd ? pmEnd : cut;
    out.push({
      ...run,
      text: run.text.slice(pieceStart - pmStart, pieceEnd - pmStart),
      ...remapNodeRange({ pmStart: pieceStart, pmEnd: pieceEnd }, context.toSource),
    });
    if (pieceEnd === pmEnd) {
      return;
    }
    pieceStart = pieceEnd;
  }
}
