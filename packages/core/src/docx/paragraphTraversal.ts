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
  Table,
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
    visitParagraphRuns(paragraph, visitRun);
  };

  const visitRun = (run: Run): void => {
    for (const content of run.content) {
      if (content.type !== "shape" || !content.shape.textBody) {
        continue;
      }
      for (const block of content.shape.textBody.content) {
        visitBlock(block);
      }
    }
  };

  const visitTable = (table: Table): void => {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        visitBlocks(cell.content);
      }
    }
  };

  const visitBlock = (block: BlockContent): void => {
    if (block.type === "paragraph") {
      visitParagraph(block);
      return;
    }
    if (block.type === "table") {
      visitTable(block);
      return;
    }
    if (block.type !== "blockSdt") {
      return;
    }
    visitBlocks(block.content);
  };

  const visitBlocks = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      visitBlock(block);
    }
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
    for (const paragraph of comment.content) {
      visitParagraph(paragraph);
    }
  }
};
