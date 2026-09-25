/**
 * Finding and replacing paragraphs in a story.
 *
 * A story is a tree of block lists: the body, each table cell and each block
 * content control holds one. A paragraph is found by its `w14:paraId`
 * wherever it sits in that tree, and a replacement rebuilds only the records
 * on the path to it, so every other block is `===` to its input.
 */

import { panic } from "better-result";

import type { BlockContent, Document, DocumentBody, Paragraph, Section } from "../model/document";
import { OP_STORIES, type OpStory } from "./types";

/** One step from a block list down to a block list nested in one of its blocks. */
type BlockListStep =
  | { kind: "tableCell"; block: number; row: number; cell: number }
  | { kind: "blockSdt"; block: number };

/** Where a paragraph sits: the block list that holds it, and its index there. */
export type ParagraphLocation = {
  list: readonly BlockListStep[];
  index: number;
  paragraph: Paragraph;
};

const collectParagraphs = (
  blocks: readonly BlockContent[],
  list: readonly BlockListStep[],
  out: ParagraphLocation[],
): void => {
  for (const [index, block] of blocks.entries()) {
    switch (block.type) {
      case "paragraph":
        out.push({ list, index, paragraph: block });
        break;
      case "table":
        for (const [row, tableRow] of block.rows.entries()) {
          for (const [cell, tableCell] of tableRow.cells.entries()) {
            collectParagraphs(
              tableCell.content,
              [...list, { kind: "tableCell", block: index, row, cell }],
              out,
            );
          }
        }
        break;
      case "blockSdt":
        collectParagraphs(block.content, [...list, { kind: "blockSdt", block: index }], out);
        break;
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
};

/** The block-level content of a story. */
export const storyBody = (document: Document, story: OpStory): DocumentBody => {
  switch (story) {
    case OP_STORIES.MAIN:
      return document.package.document;
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};

/** Every paragraph of a story's block tree, in document order. */
export const storyParagraphs = (body: DocumentBody): ParagraphLocation[] => {
  const out: ParagraphLocation[] = [];
  collectParagraphs(body.content, [], out);
  return out;
};

const sameStep = (left: BlockListStep, right: BlockListStep): boolean => {
  switch (left.kind) {
    case "tableCell":
      return (
        right.kind === "tableCell" &&
        left.block === right.block &&
        left.row === right.row &&
        left.cell === right.cell
      );
    case "blockSdt":
      return right.kind === "blockSdt" && left.block === right.block;
    default: {
      const unreachable: never = left;
      return unreachable;
    }
  }
};

/** Whether two locations are in the same block list. */
export const sameBlockList = (
  left: readonly BlockListStep[],
  right: readonly BlockListStep[],
): boolean =>
  left.length === right.length &&
  left.every((step, index) => {
    const other = right[index];
    return other !== undefined && sameStep(step, other);
  });

const updateBlockList = (
  blocks: readonly BlockContent[],
  list: readonly BlockListStep[],
  update: (blocks: readonly BlockContent[]) => BlockContent[],
): BlockContent[] => {
  const [step, ...rest] = list;
  if (step === undefined) {
    return update(blocks);
  }
  const block = blocks[step.block];
  const out = [...blocks];
  switch (step.kind) {
    case "tableCell": {
      const row = block?.type === "table" ? block.rows[step.row] : undefined;
      const cell = row?.cells[step.cell];
      if (block?.type !== "table" || row === undefined || cell === undefined) {
        return panic(`Block list step does not name a table cell at block ${step.block}.`);
      }
      const cells = [...row.cells];
      cells[step.cell] = { ...cell, content: updateBlockList(cell.content, rest, update) };
      const rows = [...block.rows];
      rows[step.row] = { ...row, cells };
      out[step.block] = { ...block, rows };
      return out;
    }
    case "blockSdt": {
      if (block?.type !== "blockSdt") {
        return panic(`Block list step does not name a block content control at ${step.block}.`);
      }
      out[step.block] = { ...block, content: updateBlockList(block.content, rest, update) };
      return out;
    }
    default: {
      const unreachable: never = step;
      return unreachable;
    }
  }
};

/**
 * Carry a body edit into `sections`, the parser's per-section view of the
 * same top-level blocks. The edited blocks are found by identity; when they
 * are not there the view was already out of step with the body and is left
 * as it was.
 */
const syncSections = (
  sections: Section[],
  before: readonly BlockContent[],
  after: readonly BlockContent[],
): Section[] => {
  let prefix = 0;
  while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix]) {
    prefix += 1;
  }
  let suffix = 0;
  while (
    suffix < before.length - prefix &&
    suffix < after.length - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  const removed = before.slice(prefix, before.length - suffix);
  const added = after.slice(prefix, after.length - suffix);
  const first = removed.at(0);
  if (first === undefined) {
    return sections;
  }
  for (const [index, section] of sections.entries()) {
    const start = section.content.indexOf(first);
    if (start === -1) {
      continue;
    }
    const matches = removed.every((block, offset) => section.content[start + offset] === block);
    if (!matches) {
      return sections;
    }
    const content = [...section.content];
    content.splice(start, removed.length, ...added);
    const next = [...sections];
    next[index] = { ...section, content };
    return next;
  }
  return sections;
};

type ReplaceParagraphsOptions = {
  document: Document;
  story: OpStory;
  /** The first paragraph replaced; the rest follow it in the same block list. */
  at: ParagraphLocation;
  count: number;
  replacement: readonly Paragraph[];
};

/** The document with `count` paragraphs from `at` replaced. */
export const replaceParagraphs = ({
  document,
  story,
  at,
  count,
  replacement,
}: ReplaceParagraphsOptions): Document => {
  const body = storyBody(document, story);
  const content = updateBlockList(body.content, at.list, (blocks) => {
    const out = [...blocks];
    out.splice(at.index, count, ...replacement);
    return out;
  });
  const nextBody: DocumentBody = { ...body, content };
  if (body.sections !== undefined) {
    nextBody.sections = syncSections(body.sections, body.content, content);
  }
  switch (story) {
    case OP_STORIES.MAIN:
      return { ...document, package: { ...document.package, document: nextBody } };
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
};
