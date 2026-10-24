import { panic } from "better-result";

import type {
  BlockContent,
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  Paragraph,
  ParagraphContent,
  Run,
  Table,
  TableRow,
} from "../types/document";

export type DocxParagraphSurfaces = {
  documentBody: DocumentBody;
  headers?: Map<string, HeaderFooter> | undefined;
  footers?: Map<string, HeaderFooter> | undefined;
  footnotes?: readonly Footnote[] | undefined;
  endnotes?: readonly Endnote[] | undefined;
};

const visitDirectParagraphRuns = (
  content: ParagraphContent,
  visit: (run: Run) => void,
): void => {
  switch (content.type) {
    case "run":
      visit(content);
      return;
    case "complexField":
      for (const run of content.fieldCode) {
        visit(run);
      }
      for (const run of content.fieldResult) {
        visit(run);
      }
      return;
    default:
      return;
  }
};

/** Visit every run directly owned by a paragraph's inline-content tree. */
export const visitParagraphRuns = (
  paragraph: Paragraph,
  visit: (run: Run) => void,
): void => {
  visitInlineContentSlots(paragraph, ({ item }) => {
    visitDirectParagraphRuns(item, visit);
  });
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
          panic(
            `Unsupported paragraph content: ${JSON.stringify(unsupported)}`,
          );
        }
      }
    }
  };

  visitContent(paragraph.content);
};

/** What the walk does below a record it has just visited. */
export const BLOCK_TREE_DESCENT = {
  descend: "descend",
  prune: "prune",
} as const;

export type BlockTreeDescent =
  (typeof BLOCK_TREE_DESCENT)[keyof typeof BLOCK_TREE_DESCENT];

type VisitBlockChildrenOptions = {
  block: BlockContent;
  visitRecord: (record: Paragraph | Table | TableRow) => BlockTreeDescent;
  visitBlocks: (blocks: BlockContent[]) => void;
};

const visitRunBlockChildren = (
  run: Run,
  visitBlocks: (blocks: BlockContent[]) => void,
): void => {
  for (const content of run.content) {
    if (content.type === "shape" && content.shape.textBody) {
      visitBlocks(content.shape.textBody.content);
    }
  }
};

const visitBlockChildren = ({
  block,
  visitRecord,
  visitBlocks,
}: VisitBlockChildrenOptions): void => {
  switch (block.type) {
    case "paragraph":
      if (visitRecord(block) === BLOCK_TREE_DESCENT.prune) {
        return;
      }
      visitParagraphRuns(block, (run) => {
        visitRunBlockChildren(run, visitBlocks);
      });
      return;
    case "table":
      if (visitRecord(block) === BLOCK_TREE_DESCENT.prune) {
        return;
      }
      for (const row of block.rows) {
        if (visitRecord(row) === BLOCK_TREE_DESCENT.prune) {
          continue;
        }
        for (const cell of row.cells) {
          visitBlocks(cell.content);
        }
      }
      return;
    case "blockSdt":
    case "blockCustomXml":
      visitBlocks(block.content);
      return;
    case "preservedBlock":
    case "bookmarkStart":
    case "bookmarkEnd":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return;
    default: {
      const unsupported: never = block;
      panic(`Unsupported block content: ${JSON.stringify(unsupported)}`);
    }
  }
};

/**
 * Every record a block tree owns in document order: a paragraph, table, and
 * table row. These are the records that carry authored state by identity.
 *
 * A block container is not always reached through a block. `w:txbxContent`
 * hangs off a shape inside a run, so a walk that recurses on block children
 * alone enters the body, a cell and an `w:sdt` and never enters a text box:
 * every record inside one is then invisible to whatever the walk is deciding.
 */
export const visitBlockTreeRecords = (
  blocks: readonly BlockContent[],
  visit: (record: Paragraph | Table | TableRow) => BlockTreeDescent,
): void => {
  const visitBlocks = (nested: readonly BlockContent[]): void => {
    for (const block of nested) {
      visitBlockChildren({ block, visitRecord: visit, visitBlocks });
    }
  };

  visitBlocks(blocks);
};

const visitDocxBlockSurfaces = (
  {
    documentBody,
    headers,
    footers,
    footnotes,
    endnotes,
  }: DocxParagraphSurfaces,
  visitBlocks: (
    blocks: BlockContent[],
    surface: "canonical" | "section",
  ) => void,
): void => {
  visitBlocks(documentBody.content, "canonical");
  for (const section of documentBody.sections ?? []) {
    visitBlocks(section.content, "section");
  }
  for (const header of headers?.values() ?? []) {
    visitBlocks(header.content, "canonical");
  }
  for (const footer of footers?.values() ?? []) {
    visitBlocks(footer.content, "canonical");
  }
  for (const footnote of footnotes ?? []) {
    visitBlocks(footnote.content, "canonical");
  }
  for (const endnote of endnotes ?? []) {
    visitBlocks(endnote.content, "canonical");
  }
  for (const comment of documentBody.comments ?? []) {
    visitBlocks(comment.content, "canonical");
  }
};

export const visitDocxParagraphs = (
  surfaces: DocxParagraphSurfaces,
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

  visitDocxBlockSurfaces(surfaces, visitBlocks);
};

export type DocxContentSlot =
  | {
      type: "block";
      content: BlockContent[];
      index: number;
      item: BlockContent;
    }
  | ({ type: "inline" } & InlineContentSlot);

export const visitDocxContentSlots = (
  surfaces: DocxParagraphSurfaces,
  visit: (slot: DocxContentSlot) => void,
): void => {
  const previousSurfaceBlocks = new WeakSet<BlockContent>();

  visitDocxBlockSurfaces(surfaces, (blocks, surface) => {
    const surfaceBlocks = new Set<BlockContent>();
    const visitBlocks = (content: BlockContent[]): void => {
      for (const [index, item] of content.entries()) {
        if (surface === "section" && previousSurfaceBlocks.has(item)) {
          continue;
        }
        surfaceBlocks.add(item);
        visit({ type: "block", content, index, item });
        visitBlockChildren({
          block: item,
          visitBlocks,
          visitRecord: (record) => {
            if (record.type === "paragraph") {
              visitInlineContentSlots(record, (slot) => {
                visit({ type: "inline", ...slot });
                visitDirectParagraphRuns(slot.item, (run) => {
                  visitRunBlockChildren(run, visitBlocks);
                });
              });
              return BLOCK_TREE_DESCENT.prune;
            }
            return BLOCK_TREE_DESCENT.descend;
          },
        });
      }
    };
    visitBlocks(blocks);
    for (const block of surfaceBlocks) {
      previousSurfaceBlocks.add(block);
    }
  });
};
