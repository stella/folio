import { panic } from "better-result";

import type { BlockContent, Document, Paragraph } from "../model/document";
import { knownToMeetContract, meetsContract } from "./contract";
import { replaceStoryBody, storyBody } from "./stories";
import type { OpStory } from "./types";

/** Whether two paragraphs hold the same value under every own field. */
const sameFields = (left: Paragraph, right: Paragraph): boolean => {
  const keys = Object.keys(left);
  return (
    keys.length === Object.keys(right).length &&
    keys.every((key) => Reflect.get(left, key) === Reflect.get(right, key))
  );
};

/** `blocks` with each paragraph `copies` names replaced, counting each one in `placed`. */
const withCopies = (
  blocks: BlockContent[],
  copies: ReadonlyMap<Paragraph, Paragraph>,
  placed: Set<Paragraph>,
): BlockContent[] => {
  let changed = false;
  const next = blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph": {
        const copy = copies.get(block);
        if (copy === undefined) return block;
        changed = true;
        placed.add(block);
        return copy;
      }
      case "table": {
        let rowsChanged = false;
        const rows = block.rows.map((row) => {
          let cellsChanged = false;
          const cells = row.cells.map((cell) => {
            const content = withCopies(cell.content, copies, placed);
            if (content === cell.content) return cell;
            cellsChanged = true;
            return { ...cell, content };
          });
          if (!cellsChanged) return row;
          rowsChanged = true;
          return { ...row, cells };
        });
        if (!rowsChanged) return block;
        changed = true;
        return { ...block, rows };
      }
      case "blockSdt":
      case "blockCustomXml": {
        const content = withCopies(block.content, copies, placed);
        if (content === block.content) return block;
        changed = true;
        return { ...block, content };
      }
      default:
        return block;
    }
  });
  return changed ? next : blocks;
};

type WithParagraphCopiesOptions = {
  document: Document;
  story: OpStory;
  /** Each paragraph of the story, at any depth, and the copy that takes its place. */
  copies: ReadonlyMap<Paragraph, Paragraph>;
};

/**
 * `document` with paragraphs of `story`, at any depth, replaced by copies
 * that hold the same value under every field, in its section views too. A
 * copy may differ only in what is kept for it outside its fields, which the
 * contract never reads, so the result meets the contract when `document` is
 * known to.
 */
export const withParagraphCopies = ({
  document,
  story,
  copies,
}: WithParagraphCopiesOptions): Document => {
  if (copies.size === 0) return document;
  for (const [paragraph, copy] of copies)
    if (!sameFields(paragraph, copy))
      panic(
        "A paragraph copy must hold the fields of the paragraph it replaces.",
      );
  const body = storyBody(document, story);
  const placed = new Set<Paragraph>();
  const content = withCopies(body.content, copies, placed);
  if (placed.size !== copies.size)
    panic("A paragraph copy must replace a paragraph of its story.");
  // Section views hold the body's blocks: they hold the same replacements.
  const replaced = new Map<BlockContent, BlockContent>();
  for (const [index, block] of body.content.entries()) {
    const next = content[index];
    if (next !== undefined && next !== block) replaced.set(block, next);
  }
  const sections = body.sections?.map((section) => {
    if (!section.content.some((block) => replaced.has(block))) return section;
    return {
      ...section,
      content: section.content.map((block) => replaced.get(block) ?? block),
    };
  });
  const result = replaceStoryBody({
    document,
    story,
    body: {
      ...body,
      content,
      ...(sections === undefined ? {} : { sections }),
    },
  });
  if (knownToMeetContract(document)) meetsContract(result);
  return result;
};
