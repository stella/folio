/**
 * Lightweight text replacement on the document model.
 * Used by find & replace. Operates immutably.
 */

import { panic } from "better-result";

import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import type {
  BlockContent,
  Document,
  DocumentBody,
  Paragraph,
  ParagraphContent,
  Run,
  TextFormatting,
} from "../types/document";
import { getFindParagraphText } from "./findReplace";

type ReplaceRange = {
  start: { paragraphIndex: number; offset: number };
  end: { paragraphIndex: number; offset: number };
};

/** Replace text within a range in the document. Returns a new document. */
export function replaceTextInDocument(
  doc: Document,
  range: ReplaceRange,
  text: string,
  formatting?: TextFormatting,
): Document {
  const newDoc = cloneDocumentWithParagraphPropertySources(doc);
  const body = newDoc.package.document;
  const { start, end } = range;

  if (start.paragraphIndex === end.paragraphIndex) {
    const paragraph = getParagraphByIndex(body, start.paragraphIndex);
    if (!paragraph) {
      return newDoc;
    }

    replaceTextInParagraph({
      paragraph,
      startOffset: start.offset,
      endOffset: end.offset,
      text,
      ...(formatting && { formatting }),
    });
  } else {
    const startParagraph = getParagraphByIndex(body, start.paragraphIndex);
    if (!startParagraph) {
      return newDoc;
    }
    const startText = getFindParagraphText(startParagraph);

    replaceTextInParagraph({
      paragraph: startParagraph,
      startOffset: start.offset,
      endOffset: startText.length,
      text,
      ...(formatting && { formatting }),
    });

    const paragraphsToRemove: number[] = [];
    for (let i = start.paragraphIndex + 1; i <= end.paragraphIndex; i++) {
      paragraphsToRemove.push(getTopLevelParagraphBlockIndex(body, i));
    }
    for (let i = paragraphsToRemove.length - 1; i >= 0; i--) {
      // SAFETY: i >= 0 and i < paragraphsToRemove.length in for loop
      const idx = paragraphsToRemove[i]!;
      if (idx !== -1) {
        body.content.splice(idx, 1);
      }
    }
  }

  return newDoc;
}

// ---------------------------------------------------------------------------
// Helpers (extracted from the former agent/executor)
// ---------------------------------------------------------------------------

function getTopLevelParagraphBlockIndex(
  body: DocumentBody,
  paragraphIndex: number,
): number {
  let pIdx = 0;
  for (let i = 0; i < body.content.length; i++) {
    // SAFETY: i < body.content.length in for loop
    if (body.content[i]!.type === "paragraph") {
      if (pIdx === paragraphIndex) {
        return i;
      }
      pIdx++;
    }
  }
  return -1;
}

function getParagraphByIndex(
  body: DocumentBody,
  paragraphIndex: number,
): Paragraph | null {
  let currentIndex = 0;

  const walkBlocks = (blocks: readonly BlockContent[]): Paragraph | null => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        if (currentIndex === paragraphIndex) {
          return block;
        }
        currentIndex++;
        continue;
      }

      if (block.type === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            const found = walkBlocks(cell.content);
            if (found !== null) {
              return found;
            }
          }
        }
        continue;
      }

      if (block.type !== "blockSdt" && block.type !== "blockCustomXml") {
        continue;
      }

      const found = walkBlocks(block.content);
      if (found !== null) {
        return found;
      }
    }
    return null;
  };

  return walkBlocks(body.content);
}

type ReplaceParagraphTextOptions = {
  paragraph: Paragraph;
  startOffset: number;
  endOffset: number;
  text: string;
  formatting?: TextFormatting;
};

const itemTextLength = (item: ParagraphContent): number =>
  getFindParagraphText({ type: "paragraph", content: [item] }).length;

/** Walk authored offsets once, preserving wrappers and opaque source units. */
const replaceTextInParagraph = ({
  paragraph,
  startOffset,
  endOffset,
  text,
  formatting,
}: ReplaceParagraphTextOptions): void => {
  let currentOffset = 0;
  let inserted = false;
  const visit = (items: ParagraphContent[]): void => {
    for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
      const item = items.at(itemIndex);
      if (!item) continue;
      switch (item.type) {
        case "run": {
          const before: Run["content"] = [];
          const after: Run["content"] = [];
          let splitAtInsertion = false;
          for (const unit of item.content) {
            const unitStart = currentOffset;
            const unitText = getFindParagraphText({
              type: "paragraph",
              content: [{ type: "run", content: [unit] }],
            });
            const unitLength = unitText.length;
            const unitEnd = unitStart + unitLength;
            currentOffset = unitEnd;
            const output = splitAtInsertion ? after : before;
            if (unit.type === "preservedXml" || unitLength === 0) {
              output.push(unit);
              continue;
            }
            const overlaps = unitEnd > startOffset && unitStart < endOffset;
            const insertionPoint =
              !inserted &&
              startOffset === endOffset &&
              startOffset >= unitStart &&
              startOffset <= unitEnd;
            if (!overlaps && !insertionPoint) {
              output.push(unit);
              continue;
            }
            if (unit.type !== "text" && !overlaps) {
              // A zero-width insertion on a native unit's boundary goes beside the unit.
              if (startOffset === unitEnd) output.push(unit);
              inserted = true;
              if (formatting) splitAtInsertion = true;
              else if (text) output.push({ type: "text", text });
              if (startOffset === unitStart)
                (splitAtInsertion ? after : before).push(unit);
              continue;
            }
            const prefix = unitText.slice(
              0,
              Math.max(0, startOffset - unitStart),
            );
            const suffix = unitText.slice(Math.max(0, endOffset - unitStart));
            if (prefix) output.push({ type: "text", text: prefix });
            if (!inserted) {
              inserted = true;
              if (formatting) splitAtInsertion = true;
              else if (text) output.push({ type: "text", text });
            }
            if (suffix)
              (splitAtInsertion ? after : before).push({
                type: "text",
                text: suffix,
              });
          }
          if (!splitAtInsertion) {
            item.content = before;
            break;
          }
          const replacement: Run[] = [];
          if (before.length > 0) replacement.push({ ...item, content: before });
          if (text)
            replacement.push({
              type: "run",
              content: [{ type: "text", text }],
              ...(formatting && { formatting }),
            });
          if (after.length > 0) replacement.push({ ...item, content: after });
          items.splice(itemIndex, 1, ...replacement);
          itemIndex += replacement.length - 1;
          break;
        }
        case "hyperlink":
          visit(item.children);
          break;
        case "inlineWrapper":
        case "inlineSdt":
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
          visit(item.content);
          break;
        case "simpleField":
        case "complexField":
        case "preservedInline":
        case "mathEquation":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "commentReference":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
          // Fields and visible atoms contribute offsets but remain opaque.
          currentOffset += itemTextLength(item);
          break;
        default: {
          const unsupported: never = item;
          panic("Unsupported replacement paragraph content", {
            content: unsupported,
          });
        }
      }
    }
  };
  visit(paragraph.content);
  if (
    !inserted &&
    startOffset === endOffset &&
    startOffset === currentOffset &&
    text
  ) {
    paragraph.content.push({
      type: "run",
      content: [{ type: "text", text }],
      ...(formatting ? { formatting } : {}),
    });
  }
};
