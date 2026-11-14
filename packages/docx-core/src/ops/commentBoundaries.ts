/**
 * Comment range markers that stand between blocks: found, taken out and put
 * back by the comment operations, addressed by a paragraph beside them so an
 * inverse can return each one to its exact place.
 */
import type { BlockContent, Document } from "../model/document";
import {
  blockListAt,
  type ParagraphLocation,
  sameBlockList,
  storyParagraphs,
  updateBlockList,
} from "./blocks";
import { idKey } from "./ids";
import { cloneModel } from "./modelClone";
import { documentStories, findStoryBody, replaceStoryBody } from "./stories";
import type { CommentBoundaryAnchor, OpStory } from "./types";

type BlockList = ParagraphLocation["list"];
type BoundaryNode = CommentBoundaryAnchor["node"];

const isBoundary = (block: BlockContent): block is BoundaryNode =>
  block.type === "commentRangeStart" || block.type === "commentRangeEnd";

/** Each block list of a story body with its steps, outermost first. */
const blockLists = (
  blocks: readonly BlockContent[],
  list: BlockList,
  visit: (blocks: readonly BlockContent[], list: BlockList) => void,
): void => {
  visit(blocks, list);
  for (const [index, block] of blocks.entries()) {
    switch (block.type) {
      case "table":
        for (const [row, tableRow] of block.rows.entries())
          for (const [cell, tableCell] of tableRow.cells.entries())
            blockLists(
              tableCell.content,
              [...list, { kind: "tableCell", block: index, row, cell }],
              visit,
            );
        break;
      case "blockSdt":
        blockLists(
          block.content,
          [...list, { kind: "blockSdt", block: index }],
          visit,
        );
        break;
      case "blockCustomXml":
        blockLists(
          block.content,
          [...list, { kind: "blockCustomXml", block: index }],
          visit,
        );
        break;
      default:
        break;
    }
  }
};

/** Whether a paragraph location lies in block `index` of `list`, or is that block. */
const inBlock = (
  location: ParagraphLocation,
  list: BlockList,
  index: number,
): boolean => {
  if (location.list.length === list.length)
    return sameBlockList(location.list, list) && location.index === index;
  return (
    location.list.length > list.length &&
    sameBlockList(location.list.slice(0, list.length), list) &&
    location.list[list.length]?.block === index
  );
};

type StoryBoundaries = {
  /** The owned markers each paragraph beside them addresses, in document order. */
  addressed: CommentBoundaryAnchor[];
  /** Owned markers no paragraph of their block list can address. */
  unaddressed: number;
};

const storyBoundaries = (
  story: OpStory,
  document: Document,
  owned: (id: number) => boolean,
): StoryBoundaries => {
  const body = findStoryBody(document, story);
  const result: StoryBoundaries = { addressed: [], unaddressed: 0 };
  if (body === undefined) return result;
  const paragraphs = storyParagraphs(body);
  blockLists(body.content, [], (blocks, list) => {
    const holdsParagraph = (index: number) =>
      paragraphs.some((location) => inBlock(location, list, index));
    for (const [index, block] of blocks.entries()) {
      if (!isBoundary(block) || !owned(block.id)) continue;
      const counted = (from: number, to: number) =>
        blocks
          .slice(from, to)
          .filter((other) => !(isBoundary(other) && owned(other.id))).length;
      let neighbour = blocks.findIndex(
        (_, at) => at > index && holdsParagraph(at),
      );
      let side: CommentBoundaryAnchor["side"] = "before";
      if (neighbour < 0) {
        neighbour = blocks.findLastIndex(
          (_, at) => at < index && holdsParagraph(at),
        );
        side = "after";
      }
      const inside = paragraphs.filter((location) =>
        inBlock(location, list, neighbour),
      );
      const paragraph = side === "before" ? inside.at(0) : inside.at(-1);
      if (neighbour < 0 || paragraph?.paragraph.paraId === undefined) {
        result.unaddressed += 1;
        continue;
      }
      result.addressed.push({
        story,
        blockId: paragraph.paragraph.paraId,
        depth: paragraph.list.length - list.length,
        side,
        gap:
          side === "before"
            ? counted(index + 1, neighbour)
            : counted(neighbour + 1, index),
        node: block,
      });
    }
  });
  return result;
};

/**
 * The markers between blocks of the comments `ids` owns, in document order,
 * or undefined when one of them has no paragraph in its block list to
 * address it. Markers inside text boxes are not among them.
 */
export const captureCommentBoundaries = (
  document: Document,
  ids: ReadonlySet<number>,
): CommentBoundaryAnchor[] | undefined => {
  const addressed: CommentBoundaryAnchor[] = [];
  for (const story of documentStories(document)) {
    const found = storyBoundaries(story, document, (id) => ids.has(id));
    if (found.unaddressed > 0) return undefined;
    addressed.push(...found.addressed);
  }
  return addressed;
};

/** The ids of every marker between blocks of a story body, text boxes aside. */
export const commentBoundaryIds = (document: Document): Set<number> => {
  const ids = new Set<number>();
  for (const story of documentStories(document))
    for (const { node } of storyBoundaries(story, document, () => true)
      .addressed)
      ids.add(node.id);
  return ids;
};

const removedFrom = (
  blocks: readonly BlockContent[],
  ids: ReadonlySet<number>,
): BlockContent[] | undefined => {
  let changed = false;
  const kept: BlockContent[] = [];
  for (const block of blocks) {
    if (isBoundary(block) && ids.has(block.id)) {
      changed = true;
      continue;
    }
    const nested = (() => {
      switch (block.type) {
        case "table": {
          let rowsChanged = false;
          const rows = block.rows.map((row) => {
            let cellsChanged = false;
            const cells = row.cells.map((cell) => {
              const content = removedFrom(cell.content, ids);
              if (content === undefined) return cell;
              cellsChanged = true;
              return { ...cell, content };
            });
            if (!cellsChanged) return row;
            rowsChanged = true;
            return { ...row, cells };
          });
          return rowsChanged ? { ...block, rows } : undefined;
        }
        case "blockSdt":
        case "blockCustomXml": {
          const content = removedFrom(block.content, ids);
          return content === undefined ? undefined : { ...block, content };
        }
        default:
          return undefined;
      }
    })();
    if (nested !== undefined) changed = true;
    kept.push(nested ?? block);
  }
  return changed ? kept : undefined;
};

/** The document without the markers between blocks of the comments `ids` owns. */
export const withoutCommentBoundaries = (
  document: Document,
  ids: ReadonlySet<number>,
): Document => {
  let current = document;
  for (const story of documentStories(document)) {
    const body = findStoryBody(current, story);
    if (body === undefined) continue;
    const content = removedFrom(body.content, ids);
    if (content === undefined) continue;
    current = replaceStoryBody({
      document: current,
      story,
      body: { ...body, content },
    });
  }
  return current;
};

/**
 * Put markers back where `boundaries` places them, or undefined when an
 * addressing paragraph or the place beside it is gone. Markers before their
 * block go back in document order and those after it in reverse, so each
 * one's count of blocks between holds as it is placed.
 */
export const restoreCommentBoundaries = (
  document: Document,
  boundaries: readonly CommentBoundaryAnchor[],
): Document | undefined => {
  let current = document;
  const ordered = [
    ...boundaries.filter(({ side }) => side === "before"),
    ...boundaries.filter(({ side }) => side === "after").toReversed(),
  ];
  for (const boundary of ordered) {
    const body = findStoryBody(current, boundary.story);
    const location =
      body &&
      storyParagraphs(body).find(
        ({ paragraph }) =>
          idKey(paragraph.paraId ?? "") === idKey(boundary.blockId),
      );
    if (
      body === undefined ||
      location === undefined ||
      !Number.isInteger(boundary.depth) ||
      boundary.depth < 0 ||
      boundary.depth > location.list.length ||
      !Number.isInteger(boundary.gap) ||
      boundary.gap < 0
    )
      return undefined;
    const list = location.list.slice(0, location.list.length - boundary.depth);
    const neighbour =
      boundary.depth === 0 ? location.index : location.list[list.length]?.block;
    if (neighbour === undefined) return undefined;
    const at =
      boundary.side === "before"
        ? neighbour - boundary.gap
        : neighbour + 1 + boundary.gap;
    if (at < 0 || at > blockListAt(body.content, list).length) return undefined;
    current = replaceStoryBody({
      document: current,
      story: boundary.story,
      body: {
        ...body,
        content: updateBlockList(body.content, list, (blocks) =>
          blocks.toSpliced(at, 0, cloneModel(boundary.node)),
        ),
      },
    });
  }
  return current;
};
