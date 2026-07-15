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
  const paragraphText = getParagraphPlainText(paragraph);
  if (!paragraphText) {
    return [];
  }

  return findAllMatches(paragraphText, searchText, options).map(({ start, end }) => ({
    paragraphIndex,
    contentIndex: findContentIndexAtOffset(paragraph, start),
    startOffset: start,
    endOffset: end,
    text: paragraphText.slice(start, end),
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

const getRunText = (run: Run): string => {
  let text = "";
  for (const item of run.content) {
    if (item.type === "text") {
      text += item.text;
    } else if (item.type === "tab") {
      text += "\t";
    } else if (item.type === "break" && item.breakType === "textWrapping") {
      text += "\n";
    }
  }
  return text;
};

const getHyperlinkText = (hyperlink: Hyperlink): string => {
  let text = "";
  for (const child of hyperlink.children) {
    if (child.type === "run") {
      text += getRunText(child);
    }
  }
  return text;
};

const getParagraphContentText = (content: ParagraphContent): string => {
  if (content.type === "run") {
    return getRunText(content);
  }
  if (content.type === "hyperlink") {
    return getHyperlinkText(content);
  }
  if (content.type === "inlineSdt") {
    return content.content.map(getParagraphContentText).join("");
  }
  if (content.type === "simpleField") {
    return content.content
      .map((child) => (child.type === "run" ? getRunText(child) : getHyperlinkText(child)))
      .join("");
  }
  if (content.type === "complexField") {
    return content.fieldResult.map(getRunText).join("");
  }
  return "";
};

const getParagraphPlainText = (paragraph: Paragraph): string =>
  paragraph.content.map(getParagraphContentText).join("");

const findContentIndexAtOffset = (paragraph: Paragraph, offset: number): number => {
  let currentOffset = 0;
  for (let contentIndex = 0; contentIndex < paragraph.content.length; contentIndex++) {
    const item = paragraph.content[contentIndex];
    if (!item) {
      continue;
    }
    const itemLength = getParagraphContentText(item).length;
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
