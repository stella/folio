import { setDocumentBodyComments } from "@stll/docx-core/model";
import { survivingComments, withBodyContent } from "@stll/docx-core/ops";

import type {
  BlockContent,
  Comment,
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
import { cloneParagraphWithPropertySource } from "./paragraphPropertySource";
import {
  withDirectRunsRewritten,
  withInlineContentRewritten,
  withRunTextBoxBlocksRewritten,
} from "./paragraphTraversal";

type CommentMarker = Extract<
  ParagraphContent,
  | { type: "commentRangeStart" }
  | { type: "commentRangeEnd" }
  | { type: "commentReference" }
>;
type TableCellBlock = TableCell["content"][number];

/**
 * A range marker under any prefix, or none: this is the guard a splice is
 * refused by, and a guard that only knows `w:` passes every patch of a part
 * that spells WordprocessingML another way.
 */
const COMMENT_RANGE_MARKER_PATTERN =
  /<(?:[\w.-]+:)?commentRange(Start|End)(?:\s[^>]*?)?\s(?:[\w.-]+:)?id=["'](\d+)["']/gu;

/**
 * Comment ids whose range is not a start followed by an end in this story's
 * XML. Either half alone is invalid OOXML and anchors nothing: Word shows no
 * comment for a range that never opens.
 */
export const unbalancedCommentRangeIds = (xml: string): Set<string> => {
  const firstStart = new Map<string, number>();
  const lastEnd = new Map<string, number>();
  for (const match of xml.matchAll(COMMENT_RANGE_MARKER_PATTERN)) {
    const [, half, id] = match;
    if (id === undefined) {
      continue;
    }
    if (half === "Start") {
      if (!firstStart.has(id)) {
        firstStart.set(id, match.index);
      }
      continue;
    }
    lastEnd.set(id, match.index);
  }

  const unbalanced = new Set<string>();
  for (const id of new Set([...firstStart.keys(), ...lastEnd.keys()])) {
    const start = firstStart.get(id);
    const end = lastEnd.get(id);
    if (start === undefined || end === undefined || end < start) {
      unbalanced.add(id);
    }
  }
  return unbalanced;
};

/**
 * Whether rewriting `originalXml` into `patchedXml` left a comment range with
 * only one half. A save patches the paragraphs an edit touched and keeps the
 * rest of the part byte-for-byte, so an edit that moves a range's only start
 * out of a patched paragraph and into an untouched one writes an end with no
 * start. Ranges the source already had unbalanced are not this save's doing
 * and stay tolerated.
 */
export const patchBreaksCommentRangeBalance = (
  originalXml: string,
  patchedXml: string,
): boolean => {
  const introduced = unbalancedCommentRangeIds(patchedXml);
  if (introduced.size === 0) {
    return false;
  }
  const preexisting = unbalancedCommentRangeIds(originalXml);
  return [...introduced].some((id) => !preexisting.has(id));
};

export const withoutOrphanCommentRanges = (doc: Document): Document => {
  const validCommentIds = new Set(
    (doc.package.document.comments ?? []).map((comment) => comment.id),
  );

  const document = withoutOrphanDocumentBodyMarkers(
    doc.package.document,
    validCommentIds,
  );
  const headers = withoutOrphanHeaderFooterMapMarkers(
    doc.package.headers,
    validCommentIds,
  );
  const footers = withoutOrphanHeaderFooterMapMarkers(
    doc.package.footers,
    validCommentIds,
  );
  const footnotes = withoutOrphanNoteMarkers(
    doc.package.footnotes,
    validCommentIds,
  );
  const endnotes = withoutOrphanNoteMarkers(
    doc.package.endnotes,
    validCommentIds,
  );

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

/** A save projection without exhausted definitions or any of their carriers. */
export const withoutExhaustedComments = (doc: Document): Document => {
  const comments = survivingComments(doc);
  if (comments.length === (doc.package.document.comments ?? []).length)
    return doc;
  const documentBody = { ...doc.package.document };
  setDocumentBodyComments(documentBody, comments);
  return withoutOrphanCommentRanges({
    ...doc,
    package: { ...doc.package, document: documentBody },
  });
};

const withoutOrphanDocumentBodyMarkers = (
  document: DocumentBody,
  validCommentIds: ReadonlySet<number>,
): DocumentBody | null => {
  const content = withoutOrphanBlockMarkers(document.content, validCommentIds);
  const comments = withoutOrphanCommentBodyMarkers(
    document.comments,
    validCommentIds,
  );
  if (!content && !comments) {
    return null;
  }
  // The section view holds the same blocks as `content`: derive it again so
  // it does not keep the markers this pass removed from the body.
  const body = content ? withBodyContent(document, content) : document;
  return comments ? { ...body, comments } : body;
};

// Comment bodies are paragraph stories of their own; a dangling
// commentRangeStart / commentReference inside one round-trips into
// comments.xml through the paragraph serializer just like body content.
const withoutOrphanCommentBodyMarkers = (
  comments: Comment[] | undefined,
  validCommentIds: ReadonlySet<number>,
): Comment[] | null => {
  if (!comments) {
    return null;
  }

  let changed = false;
  const next: Comment[] = [];
  for (const comment of comments) {
    const content = withoutOrphanParagraphListMarkers(
      comment.content,
      validCommentIds,
    );
    if (!content) {
      next.push(comment);
      continue;
    }
    changed = true;
    next.push({ ...comment, content });
  }

  return changed ? next : null;
};

const withoutOrphanParagraphListMarkers = (
  paragraphs: Paragraph[],
  validCommentIds: ReadonlySet<number>,
): Paragraph[] | null => {
  let changed = false;
  const next: Paragraph[] = [];
  for (const paragraph of paragraphs) {
    const nextParagraph = withoutOrphanParagraphMarkers(
      paragraph,
      validCommentIds,
    );
    if (nextParagraph !== paragraph) {
      changed = true;
    }
    next.push(nextParagraph);
  }

  return changed ? next : null;
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
    const content = withoutOrphanBlockMarkers(
      headerFooter.content,
      validCommentIds,
    );
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
    if (nextBlock !== null) next.push(nextBlock);
  }

  return changed ? next : null;
};

/**
 * A block with its orphaned comment markers gone, or `null` when the block is
 * itself one: a comment range marker between blocks for a comment the save no
 * longer holds.
 */
const withoutOrphanBlockMarker = (
  block: BlockContent,
  validCommentIds: ReadonlySet<number>,
): BlockContent | null => {
  switch (block.type) {
    case "paragraph":
      return withoutOrphanParagraphMarkers(block, validCommentIds);
    case "table":
      return withoutOrphanTableMarkers(block, validCommentIds);
    case "blockSdt":
    case "blockCustomXml": {
      const content = withoutOrphanBlockMarkers(block.content, validCommentIds);
      return content ? { ...block, content } : block;
    }
    case "commentRangeStart":
    case "commentRangeEnd":
      return validCommentIds.has(block.id) ? block : null;
    case "preservedBlock":
    case "bookmarkStart":
    case "bookmarkEnd":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return block;
    default: {
      const unhandled: never = block;
      return unhandled;
    }
  }
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
  const content = withoutOrphanTableCellBlockMarkers(
    cell.content,
    validCommentIds,
  );
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
    const nextBlock = withoutOrphanBlockMarker(block, validCommentIds);
    if (nextBlock !== block) {
      changed = true;
    }
    if (nextBlock !== null) next.push(nextBlock);
  }

  return changed ? next : null;
};

const withoutOrphanParagraphMarkers = (
  paragraph: Paragraph,
  validCommentIds: ReadonlySet<number>,
): Paragraph => {
  const content = withoutOrphanInlineMarkers(
    paragraph.content,
    validCommentIds,
  );
  if (content === paragraph.content) {
    return paragraph;
  }
  return cloneParagraphWithPropertySource(paragraph, { content });
};

/**
 * Inline content without the markers of comments the save no longer holds, at
 * every depth a marker can sit: inside a content control, hyperlink, field,
 * tracked or moved run wrapper, bidi or tag wrapper, and inside the text box
 * of a shape in a run. The walk descends through the same owner the model
 * validator's view is read through, so a marker the validator would find is
 * never left behind. The same array comes back when nothing is removed.
 */
const withoutOrphanInlineMarkers = <Item extends ParagraphContent>(
  content: Item[],
  validCommentIds: ReadonlySet<number>,
): Item[] => {
  let changed = false;
  const next: Item[] = [];
  for (const item of content) {
    if (isCommentMarker(item) && !validCommentIds.has(item.id)) {
      changed = true;
      continue;
    }
    const nested = withInlineContentRewritten(item, (children) =>
      withoutOrphanInlineMarkers(children, validCommentIds),
    );
    const rewritten = withDirectRunsRewritten(nested, (run) =>
      withRunTextBoxBlocksRewritten(
        run,
        (blocks) =>
          withoutOrphanBlockMarkers(blocks, validCommentIds) ?? blocks,
      ),
    );
    if (rewritten !== item) {
      changed = true;
    }
    next.push(rewritten);
  }
  return changed ? next : content;
};

const isCommentMarker = (content: ParagraphContent): content is CommentMarker =>
  content.type === "commentRangeStart" ||
  content.type === "commentRangeEnd" ||
  content.type === "commentReference";
