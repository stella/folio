import { continuesNumberedSequence } from "./paragraphSequence";
import type { FlowBlock, ParagraphBlock, TableBlock, TextBoxBlock } from "./types";

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
  if (!isEmptyParagraph(block) || preservesInheritedEmptySpacing(block)) {
    return { before, after };
  }
  return {
    before: block.attrs?.spacingExplicit?.before ? before : 0,
    after: block.attrs?.spacingExplicit?.after ? after : 0,
  };
};

/**
 * Derive sequence-level spacing suppression with structural sharing.
 *
 * Contextual and automatic list spacing affect layout values, not authored
 * OOXML formatting. Only changed branches are cloned; the input tree remains
 * untouched and repeated resolution is idempotent.
 */
export const resolveEffectiveParagraphSpacingTree = (blocks: FlowBlock[]): FlowBlock[] => {
  let changed = false;
  const resolved = blocks.map((block, index) => {
    const next = resolveBlockSpacing(blocks, index);
    if (next !== block) {
      changed = true;
    }
    return next;
  });
  return changed ? resolved : blocks;
};

export const paragraphsShareStyle = (previous: ParagraphBlock, current: ParagraphBlock): boolean =>
  previous.attrs?.styleId === current.attrs?.styleId;

const preservesInheritedEmptySpacing = (block: ParagraphBlock): boolean =>
  Boolean(
    block.attrs?.styleId ||
    block.attrs?.hasDirectParagraphFormatting ||
    block.attrs?.hasDirectParagraphMarkFormatting,
  );

const resolveBlockSpacing = (blocks: readonly FlowBlock[], index: number): FlowBlock => {
  const block = blocks[index]!; // SAFETY: called from blocks.map with the same index.
  if (block.kind === "paragraph") {
    return resolveParagraphSequenceSpacing(blocks, index, block);
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
  blocks: readonly FlowBlock[],
  index: number,
  block: ParagraphBlock,
): ParagraphBlock => {
  const spacing = block.attrs?.spacing;
  if (!spacing) {
    return block;
  }

  const previous = blocks[index - 1];
  const next = blocks[index + 1];
  const suppressBefore =
    previous?.kind === "paragraph" &&
    ((block.attrs?.contextualSpacing === true && paragraphsShareStyle(previous, block)) ||
      (block.attrs?.automaticSpacing?.before === true &&
        continuesNumberedSequence(previous, block)));
  const suppressAfter =
    next?.kind === "paragraph" &&
    ((block.attrs?.contextualSpacing === true && paragraphsShareStyle(block, next)) ||
      (block.attrs?.automaticSpacing?.after === true && continuesNumberedSequence(block, next)));

  const before = suppressBefore ? 0 : spacing.before;
  const after = suppressAfter ? 0 : spacing.after;
  if (before === spacing.before && after === spacing.after) {
    return block;
  }
  return {
    ...block,
    attrs: {
      ...block.attrs,
      spacing: { ...spacing, before, after },
    },
  };
};

const resolveTableSpacing = (block: TableBlock): TableBlock => {
  let changed = false;
  const rows = block.rows.map((row) => {
    let rowChanged = false;
    const cells = row.cells.map((cell) => {
      const blocks = resolveEffectiveParagraphSpacingTree(cell.blocks);
      if (blocks.every((candidate, index) => candidate === cell.blocks[index])) {
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
  return changed ? { ...block, rows } : block;
};

const resolveTextBoxSpacing = (block: TextBoxBlock): TextBoxBlock => {
  const content = resolveEffectiveParagraphSpacingTree(block.content);
  if (content.every((candidate, index) => candidate === block.content[index])) {
    return block;
  }
  return { ...block, content };
};
