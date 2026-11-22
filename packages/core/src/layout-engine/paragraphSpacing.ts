import { noteTableOrigin } from "./measure/tableMeasureMemo";
import { continuesNumberedSequence } from "./paragraphSequence";
import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
} from "./types";

type EffectiveParagraphSpacing = Readonly<{
  before: number;
  after: number;
}>;

type CollapseParagraphSpacingOptions = {
  before: number;
  after: number;
};

/** Collapse adjacent paragraph spacing to the larger authored side. */
export const collapseParagraphSpacing = ({
  before,
  after,
}: CollapseParagraphSpacingOptions): number => Math.max(before, after);

/** Whether a paragraph has no visible run content. */
export function isEmptyParagraph(block: ParagraphBlock): boolean {
  if (block.runs.length === 0) {
    return true;
  }
  if (block.runs.length !== 1) {
    return false;
  }
  const run = block.runs.at(0);
  return run?.kind === "text" && run.text === "";
}

/** Whether a visually empty paragraph carries authored formatting semantics. */
export const isAuthoredEmptyParagraph = (block: ParagraphBlock): boolean =>
  isEmptyParagraph(block) &&
  (block.attrs?.styleId !== undefined ||
    block.attrs?.hasDirectParagraphFormatting === true ||
    block.attrs?.hasDirectParagraphMarkFormatting === true);

/**
 * Resolve leading paragraph spacing after applying the empty-paragraph
 * inherited-spacing collapse rule.
 */
export const getParagraphSpacingBefore = (block: ParagraphBlock): number =>
  resolveEffectiveParagraphSpacing(block).before;

/**
 * Resolve trailing paragraph spacing after applying the empty-paragraph
 * inherited-spacing collapse rule.
 */
export const getParagraphSpacingAfter = (block: ParagraphBlock): number =>
  resolveEffectiveParagraphSpacing(block).after;

/** Resolve empty-paragraph spacing without modifying the authored block. */
export const resolveEffectiveParagraphSpacing = (
  block: ParagraphBlock,
): EffectiveParagraphSpacing => {
  const before = block.attrs?.spacing?.before ?? 0;
  const after = block.attrs?.spacing?.after ?? 0;
  if (!isEmptyParagraph(block) || isAuthoredEmptyParagraph(block)) {
    return { before, after };
  }
  return {
    before: block.attrs?.spacingExplicit?.before ? before : 0,
    after: block.attrs?.spacingExplicit?.after ? after : 0,
  };
};

type CellEdgeNeighbors = {
  previous: FlowBlock | undefined;
  next: FlowBlock | undefined;
};

/**
 * Derive sequence-level spacing suppression with structural sharing.
 *
 * Contextual and automatic list spacing affect layout values, not authored
 * OOXML formatting. Only changed branches are cloned; the input tree remains
 * untouched and repeated resolution is idempotent.
 */
export const resolveEffectiveParagraphSpacingTree = (
  blocks: FlowBlock[],
  cellEdges?: CellEdgeNeighbors,
): FlowBlock[] => {
  let changed = false;
  const resolved = blocks.map((block, index) => {
    const next = resolveBlockSpacing({ blocks, index, cellEdges });
    if (next !== block) {
      changed = true;
    }
    return next;
  });
  return changed ? resolved : blocks;
};

export const paragraphsShareStyle = (
  previous: ParagraphBlock,
  current: ParagraphBlock,
): boolean => previous.attrs?.styleId === current.attrs?.styleId;

type SequenceSpacingContext = {
  blocks: readonly FlowBlock[];
  index: number;
  cellEdges?: CellEdgeNeighbors | undefined;
};

const resolveBlockSpacing = (context: SequenceSpacingContext): FlowBlock => {
  const { blocks, index } = context;
  const block = blocks[index]!; // SAFETY: called from blocks.map with the same index.
  if (block.kind === "paragraph") {
    return resolveParagraphSequenceSpacing(block, context);
  }
  if (block.kind === "table") {
    return resolveTableSpacing(block);
  }
  if (block.kind === "textBox") {
    return resolveTextBoxSpacing(block);
  }
  return block;
};

const resolveParagraphSequenceSpacing = (
  block: ParagraphBlock,
  { blocks, index, cellEdges }: SequenceSpacingContext,
): ParagraphBlock => {
  const spacing = block.attrs?.spacing;
  if (!spacing) {
    return block;
  }

  const previous = blocks[index - 1];
  const next = blocks[index + 1];
  // Contextual spacing crosses a cell edge. Numbering still compares only
  // paragraphs in this cell so the list keeps its outer margins.
  const contextualPrevious = index === 0 ? cellEdges?.previous : previous;
  const contextualNext = index === blocks.length - 1 ? cellEdges?.next : next;
  const suppressBefore =
    (contextualPrevious?.kind === "paragraph" &&
      block.attrs?.contextualSpacing === true &&
      paragraphsShareStyle(contextualPrevious, block)) ||
    (block.attrs?.automaticSpacing?.before === true &&
      continuesNumberedSequence(previous, block));
  const suppressAfter =
    (contextualNext?.kind === "paragraph" &&
      block.attrs?.contextualSpacing === true &&
      paragraphsShareStyle(block, contextualNext)) ||
    (next !== undefined &&
      block.attrs?.automaticSpacing?.after === true &&
      continuesNumberedSequence(block, next));

  const beforeChanged = suppressBefore && spacing.before !== 0;
  const afterChanged = suppressAfter && spacing.after !== 0;
  if (!beforeChanged && !afterChanged) {
    return block;
  }
  const effectiveSpacing = { ...spacing };
  if (beforeChanged) {
    effectiveSpacing.before = 0;
  }
  if (afterChanged) {
    effectiveSpacing.after = 0;
  }
  return {
    ...block,
    attrs: {
      ...block.attrs,
      spacing: effectiveSpacing,
    },
  };
};

const resolveTextBoxContentSpacing = (
  blocks: TextBoxBlock["content"],
): TextBoxBlock["content"] => {
  let changed = false;
  const resolved = blocks.map((block, index) => {
    const next =
      block.kind === "table"
        ? resolveTableSpacing(block)
        : resolveParagraphSequenceSpacing(block, { blocks, index });
    if (next !== block) {
      changed = true;
    }
    return next;
  });
  return changed ? resolved : blocks;
};

/** Derive the effective paragraph spacing of every cell story in a table. */
export const resolveTableSpacing = (block: TableBlock): TableBlock => {
  let changed = false;
  const rows = block.rows.map((row) => {
    let rowChanged = false;
    const cells = row.cells.map((cell, index) => {
      const blocks = resolveEffectiveParagraphSpacingTree(cell.blocks, {
        previous: row.cells[index - 1]?.blocks.at(-1),
        next: row.cells[index + 1]?.blocks.at(0),
      });
      if (blocks === cell.blocks) {
        return cell;
      }
      rowChanged = true;
      return { ...cell, blocks };
    });
    if (!rowChanged) {
      return row;
    }
    changed = true;
    return { ...row, cells };
  });
  if (!changed) return block;
  const resolved = { ...block, rows };
  // Derived from the table alone, so it measures as the table does.
  noteTableOrigin(resolved, block);
  return resolved;
};

const resolveTextBoxSpacing = (block: TextBoxBlock): TextBoxBlock => {
  const content = resolveTextBoxContentSpacing(block.content);
  if (content === block.content) {
    return block;
  }
  return { ...block, content };
};
