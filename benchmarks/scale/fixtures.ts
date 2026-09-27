/** Synthetic, size-parametric documents for scale measurements. */

import { createEmptyDocument } from "@stll/folio-core";
import type {
  Comment,
  Document,
  Insertion,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "@stll/folio-core/types/document";

const paraId = (index: number): string => (0x4000_0000 + index).toString(16).toUpperCase();

const paragraph = (text: string, index: number): Paragraph => ({
  type: "paragraph",
  paraId: paraId(index),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const cell = (text: string, index: number): TableCell => ({
  type: "tableCell",
  content: [paragraph(text, index)],
});

/** A document with `count` distinct body paragraphs. */
export const paragraphDocument = (count: number): Document => {
  const document = createEmptyDocument();
  document.package.document.content = Array.from({ length: count }, (_, index) =>
    paragraph(`Paragraph ${index}: distinct scale fixture text.`, index + 1),
  );
  return document;
};

/** A two-column table with `rowCount` data rows and a header row. */
export const tableDocument = (rowCount: number): Document => {
  const document = createEmptyDocument();
  const rows: TableRow[] = Array.from({ length: rowCount + 1 }, (_, rowIndex) => ({
    type: "tableRow",
    cells: [
      cell(rowIndex === 0 ? "Column A" : `Row ${rowIndex - 1} A`, rowIndex * 2 + 1),
      cell(rowIndex === 0 ? "Column B" : `Row ${rowIndex - 1} B`, rowIndex * 2 + 2),
    ],
  }));
  const table: Table = { type: "table", columnWidths: [4_320, 4_320], rows };
  document.package.document.content = [table];
  return document;
};

/** A document with one anchored comment per paragraph. */
export const commentDocument = (count: number): Document => {
  const document = paragraphDocument(count);
  const comments: Comment[] = Array.from({ length: count }, (_, index) => ({
    id: index + 1,
    author: "Scale fixture",
    content: [paragraph(`Comment ${index + 1}.`, count + index + 1)],
  }));
  for (const [index, block] of document.package.document.content.entries()) {
    if (block.type !== "paragraph") {
      throw new Error(`Expected paragraph at index ${index}`);
    }
    const commentId = index + 1;
    block.content = [
      { type: "commentRangeStart", id: commentId },
      ...block.content,
      { type: "commentRangeEnd", id: commentId },
      { type: "commentReference", id: commentId },
    ];
  }
  document.package.document.comments = comments;
  return document;
};

/** A document with one distinct tracked insertion in each paragraph. */
export const trackedChangeDocument = (count: number): Document => {
  const document = paragraphDocument(count);
  for (const [index, block] of document.package.document.content.entries()) {
    if (block.type !== "paragraph") {
      throw new Error(`Expected paragraph at index ${index}`);
    }
    const insertion: Insertion = {
      type: "insertion",
      info: { id: index + 1, author: "Scale fixture" },
      content: [{ type: "run", content: [{ type: "text", text: `Inserted ${index + 1}.` }] }],
    };
    block.content.push(insertion);
  }
  return document;
};
