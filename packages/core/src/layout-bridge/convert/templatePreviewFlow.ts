/**
 * Template Fill Preview — flow-block substitution stage.
 *
 * Rewrites the FlowBlock stream the layout engine consumes so each matched
 * `{{marker}}` range lays out as if its text were the typed preview value:
 * line wrap, pagination, and following text reflow naturally instead of the
 * marker's original width persisting as dead space (the previous overlay
 * approach painted covers over the marker rects, which could not reflow).
 *
 * The substitution is strictly view-side: the ProseMirror document is never
 * touched, so the save path (which converts from the PM doc) is unaffected.
 * The value run keeps the marker's PM range ([from, to)) while carrying the
 * value text, so click-to-position and selection mapping still resolve into
 * the marker range; the host run's formatting is carried onto the value so
 * it renders like the surrounding text. `templatePreview` on the run tells
 * the painter to add the preview classes (`highlighted` paints the accent
 * chip as a layout-aware inline highlight).
 *
 * A conditional block the host reported as not applying is hidden the only way
 * a flow stage can hide one: every top-level block the span swallows whole
 * leaves the stream, so the pages paginate as if it were not in the document
 * and the following paragraph moves up. The dropped block's PM positions then
 * address no block at all, which every consumer already handles — the
 * incremental measure path bails to a full remeasure on the block-count change,
 * the page index and the rect projection find no fragment and report none.
 *
 * Untouched blocks/runs are returned by reference so the painter's
 * fingerprinting and the incremental measure path see them unchanged.
 */

import type {
  FlowBlock,
  ParagraphBlock,
  Run,
  TableBlock,
  TableCell,
  TableRow,
  TextBoxBlock,
  TextRun,
} from "../../layout-engine/types";
import type {
  TemplatePreviewHiddenRange,
  TemplatePreviewValue,
} from "../../prosemirror/plugins/templatePreviewValues";
import {
  templatePreviewHidesWholeBlock,
  templatePreviewValueFingerprint,
} from "../../prosemirror/plugins/templatePreviewValues";

/** One marker→value substitution, in PM doc positions. */
export type TemplatePreviewFlowEntry = {
  /** Inclusive PM doc position of the marker start. */
  from: number;
  /** Exclusive PM doc position of the marker end. */
  to: number;
  /** The typed value displayed in place of the marker. */
  value: TemplatePreviewValue;
};

/** What the preview substitutes and what it hides, as the flow stage sees it. */
export type TemplatePreviewFlowState = {
  entries: readonly TemplatePreviewFlowEntry[];
  /** Conditional spans a `false` condition hides, in PM doc positions. */
  hidden: readonly TemplatePreviewHiddenRange[];
};

export type TemplatePreviewFlowOptions = TemplatePreviewFlowState & {
  /** `highlighted` marks substituted runs for the accent-chip CSS. */
  mode: "highlighted" | "plain";
};

/**
 * A top-level block a hidden span swallows whole, which therefore leaves the
 * flow. A block the span only partly covers stays: its text is authored content
 * outside the conditional, and dropping the block would take that text with it.
 *
 * A block carrying no PM positions stays too. That is every section break,
 * which is the safe answer rather than an accident: a section break sets the
 * page geometry that follows it, so hiding text must not take it out of the
 * flow. `SectionBreakBlock` is the one `FlowBlock` variant without the fields,
 * hence the `in` test, matching `findDirtyBlockIndexes`.
 */
const isHiddenWholeBlock = (
  block: FlowBlock,
  hidden: readonly TemplatePreviewHiddenRange[],
): boolean => {
  const pmStart = "pmStart" in block ? block.pmStart : undefined;
  const pmEnd = "pmEnd" in block ? block.pmEnd : undefined;
  if (pmStart === undefined || pmEnd === undefined) {
    return false;
  }
  return hidden.some((range) => templatePreviewHidesWholeBlock(range, { pmStart, pmEnd }));
};

/**
 * Replace each entry's marker range with its preview value across the given
 * flow blocks (recursing into table cells and text boxes), and drop the blocks
 * a hidden conditional span swallows whole. Returns the input array unchanged
 * when there is nothing to substitute and nothing to hide.
 */
export function applyTemplatePreviewToBlocks(
  blocks: FlowBlock[],
  { entries, hidden, mode }: TemplatePreviewFlowOptions,
): FlowBlock[] {
  if (entries.length === 0 && hidden.length === 0) {
    return blocks;
  }
  const sorted = [...entries].sort((a, b) => a.from - b.from);
  let changed = false;
  const next: FlowBlock[] = [];
  for (const block of blocks) {
    if (isHiddenWholeBlock(block, hidden)) {
      changed = true;
      continue;
    }
    const transformed = transformBlock(block, sorted, mode);
    changed ||= transformed !== block;
    next.push(transformed);
  }
  return changed ? next : blocks;
}

function transformBlock(
  block: FlowBlock,
  entries: TemplatePreviewFlowEntry[],
  mode: TemplatePreviewFlowOptions["mode"],
): FlowBlock {
  if (block.kind === "paragraph") {
    return transformParagraph(block, entries, mode);
  }
  if (block.kind === "table") {
    return transformTable(block, entries, mode);
  }
  if (block.kind === "textBox") {
    return transformTextBox(block, entries, mode);
  }
  return block;
}

function transformTable(
  block: TableBlock,
  entries: TemplatePreviewFlowEntry[],
  mode: TemplatePreviewFlowOptions["mode"],
): TableBlock {
  let changed = false;
  const rows: TableRow[] = [];
  for (const row of block.rows) {
    let rowChanged = false;
    const cells: TableCell[] = [];
    for (const cell of row.cells) {
      let cellChanged = false;
      const cellBlocks: FlowBlock[] = [];
      for (const cellBlock of cell.blocks) {
        const transformed = transformBlock(cellBlock, entries, mode);
        cellChanged ||= transformed !== cellBlock;
        cellBlocks.push(transformed);
      }
      cells.push(cellChanged ? { ...cell, blocks: cellBlocks } : cell);
      rowChanged ||= cellChanged;
    }
    rows.push(rowChanged ? { ...row, cells } : row);
    changed ||= rowChanged;
  }
  return changed ? { ...block, rows } : block;
}

function transformTextBox(
  block: TextBoxBlock,
  entries: TemplatePreviewFlowEntry[],
  mode: TemplatePreviewFlowOptions["mode"],
): TextBoxBlock {
  let changed = false;
  const content: TextBoxBlock["content"] = [];
  for (const contentBlock of block.content) {
    const transformed =
      contentBlock.kind === "table"
        ? transformTable(contentBlock, entries, mode)
        : transformParagraph(contentBlock, entries, mode);
    changed ||= transformed !== contentBlock;
    content.push(transformed);
  }
  return changed ? { ...block, content } : block;
}

function transformParagraph(
  block: ParagraphBlock,
  entries: TemplatePreviewFlowEntry[],
  mode: TemplatePreviewFlowOptions["mode"],
): ParagraphBlock {
  // Cheap reject: markers never cross block boundaries, so a paragraph whose
  // PM span misses every entry passes through by reference.
  const blockFrom = block.pmStart;
  const blockTo = block.pmEnd;
  if (blockFrom !== undefined && blockTo !== undefined) {
    const touches = entries.some((entry) => entry.from < blockTo && entry.to > blockFrom);
    if (!touches) {
      return block;
    }
  }

  const runs: Run[] = [];
  let changed = false;
  for (const run of block.runs) {
    if (transformRun(run, entries, mode, runs)) {
      changed = true;
    }
  }
  return changed ? { ...block, runs } : block;
}

/**
 * Push the transformed projection of `run` onto `out`. Returns true when the
 * run was changed (sliced, replaced, or dropped); unchanged runs are pushed
 * by reference.
 */
function transformRun(
  run: Run,
  entries: TemplatePreviewFlowEntry[],
  mode: TemplatePreviewFlowOptions["mode"],
  out: Run[],
): boolean {
  const pmStart = run.pmStart;
  const pmEnd = run.pmEnd;
  if (pmStart === undefined || pmEnd === undefined) {
    out.push(run);
    return false;
  }

  const overlapping = entries.filter((entry) => entry.from < pmEnd && entry.to > pmStart);
  if (overlapping.length === 0) {
    out.push(run);
    return false;
  }

  if (run.kind !== "text") {
    // Non-text inline nodes (tab, hard break, …) occupy one PM position, so
    // an overlap means the marker swallowed them whole — drop with the rest
    // of the marker text.
    return true;
  }

  let cursor = pmStart;
  for (const entry of overlapping) {
    if (entry.from > cursor) {
      out.push(sliceTextRun(run, pmStart, cursor, entry.from));
    }
    // The run hosting the marker start contributes its formatting to the
    // value run; runs the marker merely continues through are dropped (the
    // value was already emitted at the marker start).
    if (entry.from >= pmStart) {
      out.push(...buildValueRuns(run, entry, mode));
    }
    cursor = Math.min(entry.to, pmEnd);
  }
  if (cursor < pmEnd) {
    out.push(sliceTextRun(run, pmStart, cursor, pmEnd));
  }
  return true;
}

/** Slice `run` to the PM range [from, to); positions map 1:1 onto chars. */
function sliceTextRun(run: TextRun, base: number, from: number, to: number): TextRun {
  return {
    ...run,
    text: run.text.slice(from - base, to - base),
    pmStart: from,
    pmEnd: to,
  };
}

/**
 * The value run(s) replacing a marker. A plain value is one run carrying the
 * host run's formatting; a rich value emits one run per span, each layering
 * its own bold/italic over the host formatting (host bold stays bold, a
 * span's flags OR in). Every run keeps the marker's full PM range so
 * click-to-position keeps resolving into the marker.
 */
function buildValueRuns(
  host: TextRun,
  entry: TemplatePreviewFlowEntry,
  mode: TemplatePreviewFlowOptions["mode"],
): TextRun[] {
  if (typeof entry.value === "string") {
    return [
      {
        ...host,
        text: entry.value,
        pmStart: entry.from,
        pmEnd: entry.to,
        templatePreview: mode,
      },
    ];
  }
  return entry.value.runs.map((span) => {
    const valueRun: TextRun = {
      ...host,
      text: span.text,
      pmStart: entry.from,
      pmEnd: entry.to,
      templatePreview: mode,
    };
    if (span.bold === true) {
      valueRun.bold = true;
    }
    if (span.italic === true) {
      valueRun.italic = true;
    }
    return valueRun;
  });
}

const entryKey = (entry: TemplatePreviewFlowEntry): string =>
  `${entry.from}:${entry.to}:${templatePreviewValueFingerprint(entry.value)}`;

const hiddenKey = (range: TemplatePreviewHiddenRange): string =>
  `${range.from}:${range.to}:${range.expr}`;

type KeyedRange = { from: number; to: number };

/** Widen `bounds` over every range whose key is absent from `known`. */
const growOverUnknown = <T extends KeyedRange>(
  ranges: readonly T[],
  key: (range: T) => string,
  known: ReadonlySet<string>,
  bounds: { from: number; to: number },
): void => {
  for (const range of ranges) {
    if (known.has(key(range))) {
      continue;
    }
    bounds.from = Math.min(bounds.from, range.from);
    bounds.to = Math.max(bounds.to, range.to);
  }
};

/**
 * PM range covering every substitution and every hidden span that differs
 * between two preview states (changed, added, or removed), or `null` when the
 * flow content both produce is identical. Feeds the layout pipeline's
 * dirty-range invalidation so typing a value re-measures only the blocks
 * hosting the affected markers. A hidden span that appears or disappears also
 * changes the block count, which sends the measure path down its full-remeasure
 * branch regardless of how wide this range is.
 */
export function templatePreviewDirtyRange(
  previous: TemplatePreviewFlowState,
  next: TemplatePreviewFlowState,
): { from: number; to: number } | null {
  const bounds = { from: Number.POSITIVE_INFINITY, to: Number.NEGATIVE_INFINITY };
  growOverUnknown(previous.entries, entryKey, new Set(next.entries.map(entryKey)), bounds);
  growOverUnknown(next.entries, entryKey, new Set(previous.entries.map(entryKey)), bounds);
  growOverUnknown(previous.hidden, hiddenKey, new Set(next.hidden.map(hiddenKey)), bounds);
  growOverUnknown(next.hidden, hiddenKey, new Set(previous.hidden.map(hiddenKey)), bounds);
  if (bounds.from === Number.POSITIVE_INFINITY) {
    return null;
  }
  return bounds;
}
