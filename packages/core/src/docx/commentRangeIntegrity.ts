import type {
  BlockContent,
  Document,
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  Paragraph,
  ParagraphContent,
  Table,
  TableCell,
  TableRow,
} from "../types/document";

type CommentMarker = Extract<
  ParagraphContent,
  { type: "commentRangeStart" } | { type: "commentRangeEnd" } | { type: "commentReference" }
>;
type TableCellBlock = TableCell["content"][number];

export const withoutOrphanCommentRanges = (doc: Document): Document => {
  const validCommentIds = new Set(
    (doc.package.document.comments ?? []).map((comment) => comment.id),
  );

  const document = withoutOrphanDocumentBodyMarkers(doc.package.document, validCommentIds);
  const headers = withoutOrphanHeaderFooterMapMarkers(doc.package.headers, validCommentIds);
  const footers = withoutOrphanHeaderFooterMapMarkers(doc.package.footers, validCommentIds);
  const footnotes = withoutOrphanNoteMarkers(doc.package.footnotes, validCommentIds);
  const endnotes = withoutOrphanNoteMarkers(doc.package.endnotes, validCommentIds);

  if (!document && !headers && !footers && !footnotes && !endnotes) {
    return doc;
  }

  return {
    ...doc,
    package: {
      ...doc.package,
      ...(document ? { document } : {}),
      ...(headers ? { headers } : {}),
      ...(footers ? { footers } : {}),
      ...(footnotes ? { footnotes } : {}),
      ...(endnotes ? { endnotes } : {}),
    },
  };
};

const withoutOrphanDocumentBodyMarkers = (
  document: DocumentBody,
  validCommentIds: ReadonlySet<number>,
): DocumentBody | null => {
  const content = withoutOrphanBlockMarkers(document.content, validCommentIds);
  if (!content) {
    return null;
  }
  return { ...document, content };
};

const withoutOrphanHeaderFooterMapMarkers = (
  map: Map<string, HeaderFooter> | undefined,
  validCommentIds: ReadonlySet<number>,
): Map<string, HeaderFooter> | null => {
  if (!map) {
    return null;
  }

  let changed = false;
  const next = new Map<string, HeaderFooter>();
  for (const [rId, headerFooter] of map) {
    const content = withoutOrphanBlockMarkers(headerFooter.content, validCommentIds);
    if (!content) {
      next.set(rId, headerFooter);
      continue;
    }
    changed = true;
    next.set(rId, { ...headerFooter, content });
  }

  return changed ? next : null;
};

const withoutOrphanNoteMarkers = <TNote extends Footnote | Endnote>(
  notes: TNote[] | undefined,
  validCommentIds: ReadonlySet<number>,
): TNote[] | null => {
  if (!notes) {
    return null;
  }

  let changed = false;
  const next: TNote[] = [];
  for (const note of notes) {
    const content = withoutOrphanBlockMarkers(note.content, validCommentIds);
    if (!content) {
      next.push(note);
      continue;
    }
    changed = true;
    next.push({ ...note, content });
  }

  return changed ? next : null;
};

const withoutOrphanBlockMarkers = (
  blocks: BlockContent[],
  validCommentIds: ReadonlySet<number>,
): BlockContent[] | null => {
  let changed = false;
  const next: BlockContent[] = [];
  for (const block of blocks) {
    const nextBlock = withoutOrphanBlockMarker(block, validCommentIds);
    if (nextBlock !== block) {
      changed = true;
    }
    next.push(nextBlock);
  }

  return changed ? next : null;
};

const withoutOrphanBlockMarker = (
  block: BlockContent,
  validCommentIds: ReadonlySet<number>,
): BlockContent => {
  if (block.type === "paragraph") {
    return withoutOrphanParagraphMarkers(block, validCommentIds);
  }

  if (block.type === "table") {
    return withoutOrphanTableMarkers(block, validCommentIds);
  }

  const content = withoutOrphanBlockMarkers(block.content, validCommentIds);
  if (!content) {
    return block;
  }
  return { ...block, content };
};

const withoutOrphanTableMarkers = (
  table: Table,
  validCommentIds: ReadonlySet<number>,
): Table => {
  let changed = false;
  const rows: TableRow[] = [];
  for (const row of table.rows) {
    const nextRow = withoutOrphanTableRowMarkers(row, validCommentIds);
    if (nextRow !== row) {
      changed = true;
    }
    rows.push(nextRow);
  }

  if (!changed) {
    return table;
  }
  return { ...table, rows };
};

const withoutOrphanTableRowMarkers = (
  row: TableRow,
  validCommentIds: ReadonlySet<number>,
): TableRow => {
  let changed = false;
  const cells: TableCell[] = [];
  for (const cell of row.cells) {
    const nextCell = withoutOrphanTableCellMarkers(cell, validCommentIds);
    if (nextCell !== cell) {
      changed = true;
    }
    cells.push(nextCell);
  }

  if (!changed) {
    return row;
  }
  return { ...row, cells };
};

const withoutOrphanTableCellMarkers = (
  cell: TableCell,
  validCommentIds: ReadonlySet<number>,
): TableCell => {
  const content = withoutOrphanTableCellBlockMarkers(cell.content, validCommentIds);
  if (!content) {
    return cell;
  }
  return { ...cell, content };
};

const withoutOrphanTableCellBlockMarkers = (
  blocks: TableCellBlock[],
  validCommentIds: ReadonlySet<number>,
): TableCellBlock[] | null => {
  let changed = false;
  const next: TableCellBlock[] = [];
  for (const block of blocks) {
    const nextBlock = withoutOrphanTableCellBlockMarker(block, validCommentIds);
    if (nextBlock !== block) {
      changed = true;
    }
    next.push(nextBlock);
  }

  return changed ? next : null;
};

const withoutOrphanTableCellBlockMarker = (
  block: TableCellBlock,
  validCommentIds: ReadonlySet<number>,
): TableCellBlock => {
  if (block.type === "paragraph") {
    return withoutOrphanParagraphMarkers(block, validCommentIds);
  }

  return withoutOrphanTableMarkers(block, validCommentIds);
};

const withoutOrphanParagraphMarkers = (
  paragraph: Paragraph,
  validCommentIds: ReadonlySet<number>,
): Paragraph => {
  let changed = false;
  const content: ParagraphContent[] = [];
  for (const item of paragraph.content) {
    if (isCommentMarker(item) && !validCommentIds.has(item.id)) {
      changed = true;
      continue;
    }
    content.push(item);
  }

  if (!changed) {
    return paragraph;
  }
  return { ...paragraph, content };
};

const isCommentMarker = (content: ParagraphContent): content is CommentMarker =>
  content.type === "commentRangeStart" ||
  content.type === "commentRangeEnd" ||
  content.type === "commentReference";
