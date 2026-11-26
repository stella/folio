/** The blocks a clipboard paste places, and the paragraphs they hold. */
import type {
  BlockContent,
  BlockCustomXml,
  BlockSdt,
  Paragraph,
  PreservedBlock,
  Table,
} from "../model/document";

/**
 * A block a clipboard paste places: a paragraph, a table, a block content
 * control or custom XML element holding blocks of its own, or block markup
 * kept as it was read.
 */
export type PastedBlock =
  Paragraph | Table | BlockSdt | BlockCustomXml | PreservedBlock;

/**
 * `block` with every paragraph it holds, at any depth, as `paragraph` gives
 * it, visited in document order: a table's cells and a container's blocks
 * are descended into, and every other block is kept.
 */
export const mapBlockParagraphs = (
  block: BlockContent,
  paragraph: (each: Paragraph) => Paragraph,
): BlockContent => {
  switch (block.type) {
    case "paragraph":
      return paragraph(block);
    case "table":
      return {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            content: cell.content.map((inner) =>
              mapBlockParagraphs(inner, paragraph),
            ),
          })),
        })),
      };
    case "blockSdt":
    case "blockCustomXml":
      return {
        ...block,
        content: block.content.map((inner) =>
          mapBlockParagraphs(inner, paragraph),
        ),
      };
    default:
      return block;
  }
};

/** {@link mapBlockParagraphs} of a pasted block, which stays the kind it was. */
export const mapPastedBlockParagraphs = (
  block: PastedBlock,
  paragraph: (each: Paragraph) => Paragraph,
): PastedBlock => {
  switch (block.type) {
    case "paragraph":
      return paragraph(block);
    case "table":
      return {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            content: cell.content.map((inner) =>
              mapBlockParagraphs(inner, paragraph),
            ),
          })),
        })),
      };
    case "blockSdt":
    case "blockCustomXml":
      return {
        ...block,
        content: block.content.map((inner) =>
          mapBlockParagraphs(inner, paragraph),
        ),
      };
    case "preservedBlock":
      return block;
    default: {
      const unreachable: never = block;
      return unreachable;
    }
  }
};

/** Whether `block` is one a paste places. */
export const isPastedBlock = (block: BlockContent): block is PastedBlock =>
  block.type === "paragraph" ||
  block.type === "table" ||
  block.type === "blockSdt" ||
  block.type === "blockCustomXml" ||
  block.type === "preservedBlock";

/** Every paragraph `block` holds, at any depth, in document order. */
export const blockParagraphsOf = (block: BlockContent): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  mapBlockParagraphs(block, (paragraph) => {
    paragraphs.push(paragraph);
    return paragraph;
  });
  return paragraphs;
};
