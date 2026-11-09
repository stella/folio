import { panic } from "better-result";

import type { BlockContent, Document, Paragraph } from "../model/document";
import { withBodyContent } from "./blocks";
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

/**
 * `blocks` with each paragraph `copies` names replaced, counting each one in
 * `placed` and every place one took in `places`.
 */
const withCopies = (
  blocks: BlockContent[],
  copies: ReadonlyMap<Paragraph, Paragraph>,
  placed: Set<Paragraph>,
  places: { count: number },
): BlockContent[] => {
  let changed = false;
  const next = blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph": {
        const copy = copies.get(block);
        if (copy === undefined) return block;
        changed = true;
        placed.add(block);
        places.count += 1;
        return copy;
      }
      case "table": {
        let rowsChanged = false;
        const rows = block.rows.map((row) => {
          let cellsChanged = false;
          const cells = row.cells.map((cell) => {
            const content = withCopies(cell.content, copies, placed, places);
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
        const content = withCopies(block.content, copies, placed, places);
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

/**
 * The formatting fields a list's rendering derives from its numbering and the
 * paragraph's style: the instance a style supplies and the indentation the
 * list level gives. The contract reads none of them.
 */
const RENDERING_FORMATTING_FIELDS: ReadonlySet<string> = new Set([
  "numPrFromStyle",
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
]);

/**
 * Whether `rendered` holds `paragraph`'s value under every field the
 * contract can read: every own field but its list rendering, and every
 * formatting field but those {@link RENDERING_FORMATTING_FIELDS} names.
 */
const sameBesidesRendering = (
  paragraph: Paragraph,
  rendered: Paragraph,
): boolean => {
  const keys = new Set([...Object.keys(paragraph), ...Object.keys(rendered)]);
  for (const key of keys) {
    if (key === "listRendering" || key === "formatting") continue;
    if (Reflect.get(paragraph, key) !== Reflect.get(rendered, key))
      return false;
  }
  const before: object = paragraph.formatting ?? {};
  const after: object = rendered.formatting ?? {};
  const fields = new Set([...Object.keys(before), ...Object.keys(after)]);
  for (const field of fields)
    if (
      !RENDERING_FORMATTING_FIELDS.has(field) &&
      Reflect.get(before, field) !== Reflect.get(after, field)
    )
      return false;
  return true;
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
  for (const [paragraph, copy] of copies)
    if (!sameFields(paragraph, copy))
      panic(
        "A paragraph copy must hold the fields of the paragraph it replaces.",
      );
  return (
    replacingParagraphs(document, story, copies, "held") ??
    panic("Paragraph copies replace every place their paragraphs stand.")
  );
};

type WithParagraphRenderingsOptions = {
  document: Document;
  story: OpStory;
  /** Each paragraph of the story, at any depth, and its list rendering normalized. */
  renderings: ReadonlyMap<Paragraph, Paragraph>;
};

/**
 * `document` with paragraphs of `story` replaced by ones whose list rendering
 * was normalized: each may differ from the paragraph it replaces only in its
 * list rendering and the formatting fields that rendering derives, which the
 * contract never reads, so the result meets the contract when `document` is
 * known to. A rendering depends on where its paragraph stands, so a paragraph
 * the story holds at more than one place has no one rendering: then the
 * result is undefined, and the caller replaces the story's blocks by place.
 */
export const withParagraphRenderings = ({
  document,
  story,
  renderings,
}: WithParagraphRenderingsOptions): Document | undefined => {
  for (const [paragraph, rendered] of renderings)
    if (!sameBesidesRendering(paragraph, rendered))
      panic(
        "A normalized list rendering must keep every other field of the paragraph it replaces.",
      );
  // A section view may hold a block equal to the body's rather than the
  // block itself; once the rendering changes only the body's block is in
  // step, so the views are derived from the body anew.
  return replacingParagraphs(document, story, renderings, "derived");
};

/**
 * `document` with the paragraphs `replacements` names replaced, its contract
 * mark kept. Section views take the replacement wherever they held the
 * replaced block ("held"), or are derived from the new body ("derived").
 */
const replacingParagraphs = (
  document: Document,
  story: OpStory,
  replacements: ReadonlyMap<Paragraph, Paragraph>,
  views: "held" | "derived",
): Document | undefined => {
  if (replacements.size === 0) return document;
  const body = storyBody(document, story);
  const placed = new Set<Paragraph>();
  const places = { count: 0 };
  const content = withCopies(body.content, replacements, placed, places);
  if (placed.size !== replacements.size)
    panic("A paragraph replacement must replace a paragraph of its story.");
  if (views === "derived") {
    if (places.count !== placed.size) return undefined;
    const result = replaceStoryBody({
      document,
      story,
      body: { ...body, ...withBodyContent(body, content) },
    });
    if (knownToMeetContract(document)) meetsContract(result);
    return result;
  }
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
