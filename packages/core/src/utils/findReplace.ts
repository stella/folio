import { panic } from "better-result";
import { documentStories, sameStory, storyBody, type OpStory } from "@stll/docx-core/ops";
import type { DisplayMode } from "../managers/EditorModeManager";
import { isTableCellRetainedInReviewView } from "../prosemirror/tableCellRevisionVisibility";

import type {
  BlockContent,
  Hyperlink,
  Paragraph,
  ParagraphContent,
  Run,
  Shape,
  Table,
  TableCell,
  TableRow,
} from "../types/content";
import type { Document } from "../types/document";
import { decodeOoxmlSymbolCharacter } from "./ooxmlSymbol";

type ParagraphFindMatch = {
  paragraphIndex: number;
  contentIndex: number;
  startOffset: number;
  endOffset: number;
  text: string;
};

export type FindMatch = ParagraphFindMatch &
  (
    | {
        /** Package story; omitted only by paragraph-local search. */
        story?: OpStory;
        textBoxIndex?: never;
        anchorParagraphIndex?: never;
        commentId?: never;
      }
    | {
        story: OpStory;
        /** Zero-based authored shape text-body index within the package story. */
        textBoxIndex: number;
        /** Parent-story paragraph containing the outer shape anchor. */
        anchorParagraphIndex: number;
        commentId?: never;
      }
    | {
        story: "main";
        /** Comment sidebar address; present only for opt-in comment searching. */
        commentId: number;
        textBoxIndex?: never;
        anchorParagraphIndex?: never;
      }
  );

export type FindOptions = {
  matchCase: boolean;
  matchWholeWord: boolean;
  useRegex?: boolean;
  displayMode?: DisplayMode;
  includeComments?: boolean;
};

export type FindResult = {
  matches: FindMatch[];
  totalCount: number;
  currentIndex: number;
};

/** Rendering eligibility supplied by the view; matching remains owned here. */
export type FindSearchContext = {
  renderedStories?: readonly OpStory[];
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
    const pattern = options.matchWholeWord ? `\\b(?:${source})\\b` : source;
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
  const pattern = createSearchPattern(searchText, options);
  if (!pattern) return [];

  const matches: Array<{ start: number; end: number }> = [];
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(content)) !== null) {
    if (match[0].length === 0) {
      const codePoint = content.codePointAt(pattern.lastIndex);
      pattern.lastIndex += codePoint !== undefined && codePoint > 0xffff ? 2 : 1;
      continue;
    }
    matches.push({ start: match.index, end: match.index + match[0].length });
  }
  return matches;
};

export const findInDocument = (
  document: Document | null | undefined,
  searchText: string,
  options: FindOptions,
  context?: FindSearchContext,
): Array<FindMatch & { story: OpStory }> => {
  if (!document || !searchText) {
    return [];
  }
  const body = document.package.document;
  if (!isRecord(body) || !Array.isArray(body.content)) {
    return [];
  }

  const matches: Array<FindMatch & { story: OpStory }> = [];
  // One package part contributes one story, even when rendered on many pages.
  // Main, headers, footers, footnotes, endnotes, then opt-in comments;
  // variants follow section references default/first/even. Shape text follows
  // its anchor paragraph, tables use row/cell order.
  for (const story of orderedSearchStories(document)) {
    if (
      story !== "main" &&
      (story.kind === "header" || story.kind === "footer") &&
      context?.renderedStories &&
      !context.renderedStories.some((rendered) => sameStory(rendered, story))
    )
      continue;
    const blocks = storyBody(document, story).content;
    if (story !== "main" && (story.kind === "footnote" || story.kind === "endnote")) {
      const notes =
        story.kind === "footnote" ? document.package.footnotes : document.package.endnotes;
      const noteType = notes?.find(({ id }) => id === story.id)?.noteType;
      if (noteType && noteType !== "normal") continue;
    }
    let textBoxIndex = 0;
    const searchBlocks = (
      items: readonly BlockContent[],
      shapeAddress?: {
        textBoxIndex: number;
        anchorParagraphIndex: number;
      },
      visible = true,
    ): void => {
      forEachParagraph(
        items,
        (paragraph, paragraphIndex, structureVisible) => {
          const paragraphVisible = visible && structureVisible;
          if (paragraphVisible) {
            for (const match of findInParagraph(paragraph, searchText, options, paragraphIndex)) {
              if (shapeAddress) {
                matches.push({
                  ...match,
                  story,
                  textBoxIndex: shapeAddress.textBoxIndex,
                  anchorParagraphIndex: shapeAddress.anchorParagraphIndex,
                });
              } else {
                matches.push({ ...match, story });
              }
            }
          }
          for (const { shape, visible: shapeVisible } of paragraphShapes(
            paragraph.content,
            options.displayMode,
          )) {
            if (!shape.textBody) continue;
            searchBlocks(
              shape.textBody.content,
              {
                textBoxIndex: textBoxIndex++,
                anchorParagraphIndex: shapeAddress?.anchorParagraphIndex ?? paragraphIndex,
              },
              paragraphVisible && shapeVisible,
            );
          }
        },
        options.displayMode,
      );
    };
    searchBlocks(blocks);
  }
  if (options.includeComments) {
    for (const comment of document.package.document.comments ?? []) {
      forEachParagraph(comment.content, (paragraph, paragraphIndex, visible) => {
        if (!visible) return;
        for (const match of findInParagraph(
          paragraph,
          searchText,
          { ...options, displayMode: "all-markup" },
          paragraphIndex,
        )) {
          matches.push({ ...match, story: "main", commentId: comment.id });
        }
      });
    }
  }
  return matches;
};

const orderedSearchStories = (document: Document): OpStory[] => {
  const stories = documentStories(document);
  const ordered: OpStory[] = ["main"];
  const seen = new Set<string>();
  for (const kind of ["header", "footer"] as const) {
    const add = (rId: string): void => {
      const key = `${kind}:${rId}`;
      const parts = kind === "header" ? document.package.headers : document.package.footers;
      if (seen.has(key) || !parts?.has(rId)) return;
      seen.add(key);
      ordered.push({ kind, rId });
    };
    for (const section of document.package.document.sections ?? []) {
      const refs =
        kind === "header"
          ? section.properties.headerReferences
          : section.properties.footerReferences;
      for (const variant of ["default", "first", "even"] as const) {
        for (const ref of refs ?? []) {
          if (ref.type === variant) add(ref.rId);
        }
      }
    }
    for (const story of stories) {
      if (story !== "main" && story.kind === kind) add(story.rId);
    }
  }
  for (const story of stories) {
    if (story !== "main" && (story.kind === "footnote" || story.kind === "endnote"))
      ordered.push(story);
  }
  return ordered;
};

const isHiddenRevision = (
  type: ParagraphContent["type"],
  mode: DisplayMode | undefined,
): boolean => {
  if (mode === "original") return type === "insertion" || type === "moveTo";
  if (mode === "no-markup" || mode === "simple-markup")
    return type === "deletion" || type === "moveFrom";
  return false;
};

const paragraphShapes = (content: readonly ParagraphContent[], mode: DisplayMode | undefined) => {
  const shapes: Array<{ shape: Shape; visible: boolean }> = [];
  const visit = (items: readonly ParagraphContent[], visible = true): void => {
    for (const item of items) {
      const itemVisible = visible && !isHiddenRevision(item.type, mode);
      if (item.type === "run") {
        for (const child of item.content) {
          if (child.type === "shape") shapes.push({ shape: child.shape, visible: itemVisible });
        }
      } else if (item.type === "hyperlink") {
        visit(
          item.children.filter((child) => child.type === "run" || child.type === "inlineWrapper"),
          itemVisible,
        );
      } else if (item.type === "complexField") {
        visit(item.fieldResult, itemVisible);
      } else if ("content" in item) {
        visit(item.content, itemVisible);
      }
    }
  };
  visit(content);
  return shapes;
};

/** Visible characters retain offsets in the authored paragraph, including hidden revisions. */
const getVisibleParagraphProjection = (paragraph: Paragraph, mode: DisplayMode | undefined) => {
  if (mode === undefined || mode === "all-markup") {
    const projection = joinSearchProjections(
      paragraph.content.map(getParagraphContentSearchProjection),
    );
    return { ...projection, offsets: [] };
  }
  let text = "";
  let authoredOffset = 0;
  const offsets: number[] = [];
  const pageBreakOffsets: number[] = [];
  const visit = (items: readonly ParagraphContent[]): void => {
    for (const item of items) {
      if (isHiddenRevision(item.type, mode)) {
        authoredOffset += getParagraphContentSearchProjection(item).text.length;
        continue;
      }
      if (item.type === "hyperlink") {
        visit(
          item.children.filter((child) => child.type === "run" || child.type === "inlineWrapper"),
        );
        continue;
      }
      if (item.type === "complexField") {
        visit(item.fieldResult);
        continue;
      }
      if (item.type !== "run" && "content" in item) {
        visit(item.content);
        continue;
      }
      const part = getParagraphContentSearchProjection(item);
      for (const offset of part.pageBreakOffsets) pageBreakOffsets.push(text.length + offset);
      for (let index = 0; index < part.text.length; index++) offsets.push(authoredOffset + index);
      text += part.text;
      authoredOffset += part.text.length;
    }
  };
  visit(paragraph.content);
  return { text, offsets, pageBreakOffsets };
};

/** Authored paragraph text displayed by comment UI, using the search owner. */
export const getFindParagraphText = (paragraph: Paragraph): string =>
  getVisibleParagraphProjection(paragraph, undefined).text;

export const findInParagraph = (
  paragraph: Paragraph,
  searchText: string,
  options: FindOptions,
  paragraphIndex: number,
): ParagraphFindMatch[] => {
  const projection = getVisibleParagraphProjection(paragraph, options.displayMode);
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
      contentIndex: findContentIndexAtOffset(paragraph, projection.offsets.at(start) ?? start),
      startOffset: projection.offsets.at(start) ?? start,
      endOffset: (projection.offsets.at(end - 1) ?? end - 1) + 1,
      text: projection.text.slice(start, end),
    }));
};

type ParagraphVisitor = (paragraph: Paragraph, paragraphIndex: number, visible: boolean) => void;

const forEachParagraph = (
  blocks: readonly BlockContent[],
  visit: ParagraphVisitor,
  displayMode?: DisplayMode,
): void => {
  let paragraphIndex = 0;
  const walkBlocks = (items: readonly BlockContent[], visible = true): void => {
    for (const block of items) {
      if (isParagraph(block)) {
        visit(block, paragraphIndex, visible);
        paragraphIndex++;
        continue;
      }
      if (isTable(block)) {
        walkTable(block, visible);
        continue;
      }
      if (isBlockContainer(block)) {
        walkBlocks(block.content, visible);
      }
    }
  };

  const walkTable = (table: Table, visible: boolean): void => {
    for (const row of table.rows) {
      if (!isTableRow(row)) {
        continue;
      }
      for (const cell of row.cells) {
        if (isTableCell(cell)) {
          walkBlocks(
            cell.content,
            visible &&
              isVisibleStructure(row.structuralChange, displayMode) &&
              isVisibleStructure(cell.structuralChange, displayMode),
          );
        }
      }
    }
  };

  walkBlocks(blocks);
};

const isVisibleStructure = (
  change: TableRow["structuralChange"],
  mode: DisplayMode | undefined,
): boolean => {
  if (mode === undefined || mode === "all-markup" || !change) return true;
  let kind: "ins" | "del" | "merge";
  switch (change.type) {
    case "tableRowInsertion":
    case "tableCellInsertion":
      kind = "ins";
      break;
    case "tableRowDeletion":
    case "tableCellDeletion":
      kind = "del";
      break;
    case "tableCellMerge":
      kind = "merge";
      break;
    default: {
      const exhaustive: never = change;
      return exhaustive;
    }
  }
  return isTableCellRetainedInReviewView(kind, mode === "original" ? "original" : "final");
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
    } else if (item.type === "symbol") {
      text += decodeOoxmlSymbolCharacter(item.char) ?? "";
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

// A transparent wrapper inside the link is read through: the text it holds is
// the link's text, and the editor puts it on the line.
const getHyperlinkSearchProjection = (hyperlink: Hyperlink): SearchProjection =>
  joinSearchProjections(
    hyperlink.children.flatMap((child) => {
      if (child.type === "run") {
        return [getRunSearchProjection(child)];
      }
      return child.type === "inlineWrapper" ? [getParagraphContentSearchProjection(child)] : [];
    }),
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

const isBlockContainer = (
  value: unknown,
): value is Extract<BlockContent, { type: "blockSdt" | "blockCustomXml" }> =>
  isRecord(value) &&
  (value["type"] === "blockSdt" || value["type"] === "blockCustomXml") &&
  Array.isArray(value["content"]);
