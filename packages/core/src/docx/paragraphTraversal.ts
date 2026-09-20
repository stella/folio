import { panic } from "better-result";

import type {
  BlockContent,
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  Run,
  TableRow,
  TrackedRunContent,
} from "../types/document";

export type DocxParagraphSurfaces = {
  documentBody: DocumentBody;
  headers?: Map<string, HeaderFooter> | undefined;
  footers?: Map<string, HeaderFooter> | undefined;
  footnotes?: readonly Footnote[] | undefined;
  endnotes?: readonly Endnote[] | undefined;
};

/** Visit every run directly owned by a paragraph's inline-content tree. */
export const visitParagraphRuns = (paragraph: Paragraph, visit: (run: Run) => void): void => {
  type ParagraphRunTreeNode = ParagraphContent | TrackedRunContent | InlineSdt["content"][number];

  const visitParagraphContent = (content: ParagraphRunTreeNode): void => {
    switch (content.type) {
      case "run":
        visit(content);
        return;
      case "hyperlink":
        for (const child of content.children) {
          visitParagraphContent(child);
        }
        return;
      case "simpleField":
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
      case "inlineSdt":
      // Transparent: the runs inside a bidirectional wrapper are the
      // paragraph's runs, laid out differently.
      case "inlineWrapper":
        for (const child of content.content) {
          visitParagraphContent(child);
        }
        return;
      case "complexField":
        for (const run of content.fieldCode) {
          visit(run);
        }
        for (const run of content.fieldResult) {
          visit(run);
        }
        return;
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "mathEquation":
      case "preservedInline":
        return;
      default: {
        const unsupported: never = content;
        return unsupported;
      }
    }
  };

  for (const content of paragraph.content) {
    visitParagraphContent(content);
  }
};

/**
 * One position in an inline-content array: the array itself, the index, and
 * what sits there. A normaliser that drops or rewrites a marker needs the
 * array and the index, not only the value.
 */
export type InlineContentSlot = {
  content: ParagraphContent[];
  index: number;
  item: ParagraphContent;
};

/**
 * Visit every inline-content position a paragraph owns, in document order,
 * descending through every wrapper that is transparent to a range marker.
 *
 * The model validator walks the whole inline tree; a normaliser that walks
 * only `paragraph.content` sees a different document from the one the
 * validator judges, and a marker inside `w:ins`, `w:hyperlink`, `w:sdt`,
 * `w:bdo` or `w:dir` then reaches the validator unnormalised. Both sides read
 * the tree through this one traversal so they cannot disagree again.
 *
 * A complex field's runs are skipped: `fieldCode` and `fieldResult` hold runs
 * only, and a run is not a marker position.
 */
export const visitInlineContentSlots = (
  paragraph: Paragraph,
  visit: (slot: InlineContentSlot) => void,
): void => {
  const visitContent = (content: ParagraphContent[]): void => {
    for (const [index, item] of content.entries()) {
      visit({ content, index, item });
      switch (item.type) {
        case "hyperlink":
          visitContent(item.children);
          break;
        case "simpleField":
        case "inlineSdt":
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
        case "inlineWrapper":
          visitContent(item.content);
          break;
        case "run":
        case "complexField":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "commentReference":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
        case "mathEquation":
        case "preservedInline":
          break;
        default: {
          const unsupported: never = item;
          panic(`Unsupported paragraph content: ${JSON.stringify(unsupported)}`);
        }
      }
    }
  };

  visitContent(paragraph.content);
};

/**
 * Positions marked for removal, per inline-content array.
 *
 * Removal shifts every later index in that array, so a normaliser records the
 * positions while it reads and drops them once, after it has finished reading.
 */
export class InlineContentRemovals {
  readonly #byContent = new Map<ParagraphContent[], Set<number>>();

  mark({ content, index }: Pick<InlineContentSlot, "content" | "index">): void {
    const indexes = this.#byContent.get(content);
    if (indexes) {
      indexes.add(index);
      return;
    }
    this.#byContent.set(content, new Set([index]));
  }

  /** Applies every marked removal and answers how many items were dropped. */
  apply(): number {
    let removed = 0;
    for (const [content, indexes] of this.#byContent) {
      if (indexes.size === 0) {
        continue;
      }
      const next = content.filter((_, index) => !indexes.has(index));
      removed += content.length - next.length;
      content.length = 0;
      content.push(...next);
    }
    return removed;
  }
}

/** What the walk does below a record it has just visited. */
export const BLOCK_TREE_DESCENT = {
  descend: "descend",
  prune: "prune",
} as const;

export type BlockTreeDescent = (typeof BLOCK_TREE_DESCENT)[keyof typeof BLOCK_TREE_DESCENT];

/**
 * Every record a block tree owns in document order: a paragraph and a table
 * row. These are the two the editor has a record for, and the two an attribute
 * remainder can ride on.
 *
 * A block container is not always reached through a block. `w:txbxContent`
 * hangs off a shape inside a run, so a walk that recurses on block children
 * alone enters the body, a cell and an `w:sdt` and never enters a text box:
 * every record inside one is then invisible to whatever the walk is deciding.
 */
export const visitBlockTreeRecords = (
  blocks: readonly BlockContent[],
  visit: (record: Paragraph | TableRow) => BlockTreeDescent,
): void => {
  const visitParagraph = (paragraph: Paragraph): void => {
    if (visit(paragraph) === BLOCK_TREE_DESCENT.prune) {
      return;
    }
    visitParagraphRuns(paragraph, (run) => {
      for (const content of run.content) {
        if (content.type === "shape" && content.shape.textBody) {
          visitBlocks(content.shape.textBody.content);
        }
      }
    });
  };

  const visitBlocks = (nested: readonly BlockContent[]): void => {
    for (const block of nested) {
      switch (block.type) {
        case "paragraph":
          visitParagraph(block);
          break;
        case "table":
          for (const row of block.rows) {
            if (visit(row) === BLOCK_TREE_DESCENT.prune) {
              continue;
            }
            for (const cell of row.cells) {
              visitBlocks(cell.content);
            }
          }
          break;
        case "blockSdt":
          visitBlocks(block.content);
          break;
        // Opaque markup: folio models nothing inside it, so it owns no record.
        case "preservedBlock":
          break;
        default: {
          const unsupported: never = block;
          panic(`Unsupported block content: ${JSON.stringify(unsupported)}`);
        }
      }
    }
  };

  visitBlocks(blocks);
};

export const visitDocxParagraphs = (
  { documentBody, headers, footers, footnotes, endnotes }: DocxParagraphSurfaces,
  visit: (paragraph: Paragraph) => void,
): void => {
  const seenParagraphs = new WeakSet<Paragraph>();

  const visitParagraph = (paragraph: Paragraph): void => {
    if (seenParagraphs.has(paragraph)) {
      return;
    }
    seenParagraphs.add(paragraph);
    visit(paragraph);
  };

  // A surface can hand back a paragraph another surface already gave: the
  // sections partition the body's own blocks. Pruning at the repeat keeps the
  // callback exactly once per paragraph and the subtree walked exactly once.
  const visitBlocks = (blocks: readonly BlockContent[]): void => {
    visitBlockTreeRecords(blocks, (record) => {
      if (record.type !== "paragraph") {
        return BLOCK_TREE_DESCENT.descend;
      }
      if (seenParagraphs.has(record)) {
        return BLOCK_TREE_DESCENT.prune;
      }
      visitParagraph(record);
      return BLOCK_TREE_DESCENT.descend;
    });
  };

  visitBlocks(documentBody.content);
  for (const section of documentBody.sections ?? []) {
    visitBlocks(section.content);
  }
  for (const header of headers?.values() ?? []) {
    visitBlocks(header.content);
  }
  for (const footer of footers?.values() ?? []) {
    visitBlocks(footer.content);
  }
  for (const footnote of footnotes ?? []) {
    visitBlocks(footnote.content);
  }
  for (const endnote of endnotes ?? []) {
    visitBlocks(endnote.content);
  }
  for (const comment of documentBody.comments ?? []) {
    visitBlocks(comment.content);
  }
};
