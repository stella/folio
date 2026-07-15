import type { Node as ProseMirrorNode } from "prosemirror-model";

import { findAllMatches } from "../utils/findReplace";
import type { FindOptions } from "../utils/findReplace";

/**
 * Positional fields of a document search match that the range resolver needs.
 *
 * The React find-replace producer carries a richer match shape (content index,
 * matched text) for its UI; resolution only consumes the paragraph-relative
 * offsets, so this declares exactly that subset and keeps core free of any
 * adapter-side type.
 */
export type FindMatchPosition = {
  /** Index of the paragraph containing the match. */
  paragraphIndex: number;
  /** Character offset of the match start within the paragraph. */
  startOffset: number;
  /** Character offset of the match end within the paragraph. */
  endOffset: number;
};

export type FindMatchRange = {
  from: number;
  to: number;
};

export type ProseMirrorFindMatch = FindMatchPosition &
  FindMatchRange & {
    text: string;
  };

export function resolveFindMatchRange(
  doc: ProseMirrorNode,
  match: FindMatchPosition,
): FindMatchRange | null {
  let resolved: FindMatchRange | null = null;
  forEachSearchParagraph(doc, (paragraph, paragraphPos, paragraphIndex) => {
    if (paragraphIndex !== match.paragraphIndex) {
      return true;
    }
    resolved = resolveTextRangeInParagraph({
      paragraph,
      paragraphPos,
      startOffset: match.startOffset,
      endOffset: match.endOffset,
    });
    return false;
  });

  return resolved;
}

export function findInProseMirrorDocument(
  doc: ProseMirrorNode,
  searchText: string,
  options: FindOptions,
): ProseMirrorFindMatch[] {
  if (!searchText) {
    return [];
  }

  const matches: ProseMirrorFindMatch[] = [];
  forEachSearchParagraph(doc, (paragraph, paragraphPos, paragraphIndex) => {
    const text = getSearchableParagraphText(paragraph);
    for (const { start, end } of findAllMatches(text, searchText, options)) {
      const range = resolveTextRangeInParagraph({
        paragraph,
        paragraphPos,
        startOffset: start,
        endOffset: end,
      });
      if (range) {
        matches.push({
          paragraphIndex,
          startOffset: start,
          endOffset: end,
          text: text.slice(start, end),
          ...range,
        });
      }
    }
    return true;
  });
  return matches;
}

type SearchParagraphVisitor = (
  paragraph: ProseMirrorNode,
  paragraphPos: number,
  paragraphIndex: number,
) => boolean;

function forEachSearchParagraph(doc: ProseMirrorNode, visit: SearchParagraphVisitor): void {
  let paragraphIndex = 0;
  const walkBlocks = (container: ProseMirrorNode, contentStart: number): boolean => {
    let offset = 0;
    for (let childIndex = 0; childIndex < container.childCount; childIndex++) {
      const child = container.child(childIndex);
      const childPos = contentStart + offset;
      if (child.type.name === "paragraph") {
        if (!visit(child, childPos, paragraphIndex)) {
          return false;
        }
        paragraphIndex++;
      } else if (child.type.name === "table" && !walkTable(child, childPos)) {
        return false;
      } else if (child.type.name === "blockSdt" && !walkBlocks(child, childPos + 1)) {
        return false;
      }
      offset += child.nodeSize;
    }
    return true;
  };

  const walkTable = (table: ProseMirrorNode, tablePos: number): boolean => {
    let rowOffset = 0;
    for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
      const row = table.child(rowIndex);
      if (row.type.name !== "tableRow") {
        rowOffset += row.nodeSize;
        continue;
      }
      const rowPos = tablePos + 1 + rowOffset;
      let cellOffset = 0;
      for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
        const cell = row.child(cellIndex);
        if (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader") {
          cellOffset += cell.nodeSize;
          continue;
        }
        const cellPos = rowPos + 1 + cellOffset;
        if (!walkBlocks(cell, cellPos + 1)) {
          return false;
        }
        cellOffset += cell.nodeSize;
      }
      rowOffset += row.nodeSize;
    }
    return true;
  };

  walkBlocks(doc, 0);
}

type ResolveTextRangeInParagraphOptions = {
  paragraph: ProseMirrorNode;
  paragraphPos: number;
  startOffset: number;
  endOffset: number;
};

function resolveTextRangeInParagraph({
  paragraph,
  paragraphPos,
  startOffset,
  endOffset,
}: ResolveTextRangeInParagraphOptions): FindMatchRange | null {
  let textOffset = 0;
  const range = { from: null as number | null, to: null as number | null };

  paragraph.descendants((node, pos) => {
    const tokenLength = getSearchTextTokenLength(node);
    if (tokenLength === 0) {
      return true;
    }

    const textStart = textOffset;
    const textEnd = textStart + tokenLength;
    const nodeStart = paragraphPos + 1 + pos;

    if (range.from === null && startOffset >= textStart && startOffset <= textEnd) {
      range.from = nodeStart + Math.min(startOffset - textStart, node.nodeSize);
    }
    if (range.to === null && endOffset >= textStart && endOffset <= textEnd) {
      range.to = nodeStart + Math.min(endOffset - textStart, node.nodeSize);
    }

    textOffset = textEnd;
    return range.to === null;
  });

  if (range.from === null || range.to === null || range.from >= range.to) {
    return null;
  }

  return { from: range.from, to: range.to };
}

function getSearchTextTokenLength(node: ProseMirrorNode): number {
  if (node.isText) {
    return node.text?.length ?? 0;
  }

  if (node.type.name === "tab" || node.type.name === "hardBreak") {
    return 1;
  }

  return 0;
}

function getSearchableParagraphText(paragraph: ProseMirrorNode): string {
  let text = "";
  paragraph.descendants((node) => {
    if (node.isText) {
      text += node.text ?? "";
      return false;
    }
    if (node.type.name === "tab") {
      text += "\t";
      return false;
    }
    if (node.type.name === "hardBreak") {
      text += "\n";
      return false;
    }
    return true;
  });
  return text;
}
