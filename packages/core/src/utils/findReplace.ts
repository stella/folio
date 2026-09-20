import { panic } from "better-result";

import type {
  BlockContent,
  Hyperlink,
  Paragraph,
  ParagraphContent,
  Run,
  Table,
  TableCell,
  TableRow,
} from "../types/content";
import type { Document } from "../types/document";

export type FindMatch = {
  paragraphIndex: number;
  contentIndex: number;
  startOffset: number;
  endOffset: number;
  text: string;
};

export type FindOptions = {
  matchCase: boolean;
  matchWholeWord: boolean;
  useRegex?: boolean;
};

export type FindResult = {
  matches: FindMatch[];
  totalCount: number;
  currentIndex: number;
};

export const createDefaultFindOptions = (): FindOptions => ({
  matchCase: false,
  matchWholeWord: false,
  useRegex: false,
});

export const escapeRegexString = (value: string): string =>
  value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");

export const createSearchPattern = (searchText: string, options: FindOptions): RegExp | null => {
  if (!searchText) {
    return null;
  }

  try {
    const source = options.useRegex ? searchText : escapeRegexString(searchText);
    const pattern = options.matchWholeWord ? `\\b${source}\\b` : source;
    return new RegExp(pattern, options.matchCase ? "gu" : "giu");
  } catch {
    return null;
  }
};

export const findAllMatches = (
  content: string,
  searchText: string,
  options: FindOptions,
): Array<{ start: number; end: number }> => {
  if (!content || !searchText) {
    return [];
  }
  const searchFor = options.matchCase ? searchText : searchText.toLowerCase();
  const source = escapeRegexString(searchFor);
  const pattern = new RegExp(
    options.matchWholeWord ? `\\b${source}\\b` : source,
    options.matchCase ? "gu" : "giu",
  );

  const matches: Array<{ start: number; end: number }> = [];
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(content)) !== null) {
    matches.push({ start: match.index, end: match.index + match[0].length });
    if (match[0].length === 0) {
      pattern.lastIndex++;
    }
  }
  return matches;
};

export const findInDocument = (
  document: Document | null | undefined,
  searchText: string,
  options: FindOptions,
): FindMatch[] => {
  if (!document || !searchText) {
    return [];
  }
  const body = document.package.document;
  if (!isRecord(body) || !Array.isArray(body.content)) {
    return [];
  }

  const matches: FindMatch[] = [];
  forEachParagraph(body.content, (paragraph, paragraphIndex) => {
    matches.push(...findInParagraph(paragraph, searchText, options, paragraphIndex));
  });
  return matches;
};

export const findInParagraph = (
  paragraph: Paragraph,
  searchText: string,
  options: FindOptions,
  paragraphIndex: number,
): FindMatch[] => {
  const projection = getParagraphSearchProjection(paragraph);
  if (!projection.text) {
    return [];
  }

  return findAllMatches(projection.text, searchText, options)
    .filter(
      ({ start, end }) =>
        !projection.pageBreakOffsets.some((offset) => offset > start && offset < end),
    )
    .map(({ start, end }) => ({
      paragraphIndex,
      contentIndex: findContentIndexAtOffset(paragraph, start),
      startOffset: start,
      endOffset: end,
      text: projection.text.slice(start, end),
    }));
};

type ParagraphVisitor = (paragraph: Paragraph, paragraphIndex: number) => void;

const forEachParagraph = (blocks: readonly BlockContent[], visit: ParagraphVisitor): void => {
  let paragraphIndex = 0;
  const walkBlocks = (items: readonly BlockContent[]): void => {
    for (const block of items) {
      if (isParagraph(block)) {
        visit(block, paragraphIndex);
        paragraphIndex++;
        continue;
      }
      if (isTable(block)) {
        walkTable(block);
        continue;
      }
      if (isBlockSdt(block)) {
        walkBlocks(block.content);
      }
    }
  };

  const walkTable = (table: Table): void => {
    for (const row of table.rows) {
      if (!isTableRow(row)) {
        continue;
      }
      for (const cell of row.cells) {
        if (isTableCell(cell)) {
          walkBlocks(cell.content);
        }
      }
    }
  };

  walkBlocks(blocks);
};

type SearchProjection = {
  text: string;
  pageBreakOffsets: number[];
};

const joinSearchProjections = (parts: readonly SearchProjection[]): SearchProjection => {
  let text = "";
  const pageBreakOffsets: number[] = [];
  for (const part of parts) {
    const baseOffset = text.length;
    text += part.text;
    for (const offset of part.pageBreakOffsets) {
      pageBreakOffsets.push(baseOffset + offset);
    }
  }
  return { text, pageBreakOffsets };
};

const getRunSearchProjection = (run: Run): SearchProjection => {
  let text = "";
  const pageBreakOffsets: number[] = [];
  for (const item of run.content) {
    if (item.type === "text") {
      text += item.text;
    } else if (item.type === "tab") {
      text += "\t";
    } else if (
      item.type === "break" &&
      (item.breakType === undefined || item.breakType === "textWrapping")
    ) {
      text += "\n";
    } else if (item.type === "break" && item.breakType === "page") {
      pageBreakOffsets.push(text.length);
    }
  }
  return { text, pageBreakOffsets };
};

const getHyperlinkSearchProjection = (hyperlink: Hyperlink): SearchProjection =>
  joinSearchProjections(
    hyperlink.children.flatMap((child) =>
      child.type === "run" ? [getRunSearchProjection(child)] : [],
    ),
  );

/** Nothing on the line: a marker, or markup with no text under it. */
const noSearchText = (): SearchProjection => ({ text: "", pageBreakOffsets: [] });

/**
 * The text one paragraph content item puts on the line, for searching.
 *
 * A `switch` with a `never` default rather than an `if`-chain ending in the
 * empty projection: a chain answers "no text" for a member nobody considered,
 * and a member that holds runs and answers "no text" is text the reader can
 * see and find cannot. That is what the revision wrappers and the transparent
 * inline wrappers were doing — a search for a phrase inside a `w:bdo`, a smart
 * tag or an insertion found nothing.
 *
 * The offsets this produces are resolved back to an editor position by
 * `resolveFindMatchRange`, against the text `getSearchableParagraphText`
 * builds from the ProseMirror tree. So the two have to agree character for
 * character, and this one mirrors that one: every text node counts, including
 * the deleted and moved-away text the editor still shows struck through. A
 * member skipped here shifts every later offset and the replacement lands on
 * the wrong characters.
 */
const getParagraphContentSearchProjection = (content: ParagraphContent): SearchProjection => {
  switch (content.type) {
    case "run":
      return getRunSearchProjection(content);
    case "hyperlink":
      return getHyperlinkSearchProjection(content);
    // Every member of these is paragraph content, so the same projection reads
    // it: a narrowing per type would be a mirror of the union that drifts the
    // next time the union grows.
    case "inlineSdt":
    case "simpleField":
    case "inlineWrapper":
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
      return joinSearchProjections(content.content.map(getParagraphContentSearchProjection));
    case "complexField":
      return joinSearchProjections(content.fieldResult.map(getRunSearchProjection));
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
      return noSearchText();
    default: {
      const unprojected: never = content;
      panic(`Unsupported paragraph content: ${JSON.stringify(unprojected)}`);
    }
  }
};

const getParagraphSearchProjection = (paragraph: Paragraph): SearchProjection =>
  joinSearchProjections(paragraph.content.map(getParagraphContentSearchProjection));

const findContentIndexAtOffset = (paragraph: Paragraph, offset: number): number => {
  let currentOffset = 0;
  for (let contentIndex = 0; contentIndex < paragraph.content.length; contentIndex++) {
    const item = paragraph.content[contentIndex];
    if (!item) {
      continue;
    }
    const itemLength = getParagraphContentSearchProjection(item).text.length;
    if (currentOffset + itemLength > offset) {
      return contentIndex;
    }
    currentOffset += itemLength;
  }
  return Math.max(0, paragraph.content.length - 1);
};

const isRecord = (value: unknown): value is Record<PropertyKey, unknown> =>
  typeof value === "object" && value !== null;

const isParagraph = (value: unknown): value is Paragraph =>
  isRecord(value) && value["type"] === "paragraph" && Array.isArray(value["content"]);

const isTable = (value: unknown): value is Table =>
  isRecord(value) && value["type"] === "table" && Array.isArray(value["rows"]);

const isTableRow = (value: unknown): value is TableRow =>
  isRecord(value) && Array.isArray(value["cells"]);

const isTableCell = (value: unknown): value is TableCell =>
  isRecord(value) && Array.isArray(value["content"]);

const isBlockSdt = (value: unknown): value is Extract<BlockContent, { type: "blockSdt" }> =>
  isRecord(value) && value["type"] === "blockSdt" && Array.isArray(value["content"]);
