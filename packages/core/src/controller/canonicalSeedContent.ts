import {
  commentAnchorSequence,
  commentDocumentIssue,
  documentStories,
  findStoryBody,
  OP_STORIES,
  type OpStory,
} from "@stll/docx-core/ops";

import type { BlockContent, Document, Paragraph } from "../types/document";
import { canonicalInlineSourceIssue } from "./canonicalInlineProjection";

const STRUCTURE_ISSUE =
  "Canonical sessions currently require paragraphs, tables and supported inline atoms.";
const REVIEW_CARRIER_ISSUE =
  "Canonical sessions currently require plain paragraphs and supported inline atoms.";

/** Range markers may open or close between blocks; they hold no text. */
const BETWEEN_BLOCK_MARKERS = new Set<BlockContent["type"]>([
  "commentRangeStart",
  "commentRangeEnd",
  "bookmarkStart",
  "bookmarkEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

type BlockIssues = {
  /** The first block the session cannot hold, in document order. */
  structure: string | null;
  /** The first paragraph whose content the session cannot hold. */
  paragraph: string | null;
};

const issuesIn = (block: BlockContent): BlockIssues => {
  const paragraphs: Paragraph[] = [];
  const collect = (blocks: readonly BlockContent[]): string | null => {
    for (const child of blocks) {
      if (child.type === "paragraph") {
        paragraphs.push(child);
        continue;
      }
      if (BETWEEN_BLOCK_MARKERS.has(child.type)) continue;
      // Opaque markup the session keeps as it found it and never edits inside.
      if (child.type === "preservedBlock") continue;
      // A block content control holds its own block list, as a cell does.
      if (child.type === "blockSdt") {
        const issue = collect(child.content);
        if (issue !== null) return issue;
        continue;
      }
      if (child.type !== "table") return STRUCTURE_ISSUE;
      for (const row of child.rows) {
        for (const cell of row.cells) {
          const issue = collect(cell.content);
          if (issue !== null) return issue;
        }
      }
    }
    return null;
  };
  const structure = collect([block]);
  let paragraph: string | null = null;
  for (const candidate of paragraphs) {
    paragraph =
      candidate.reviewCarrier === undefined
        ? canonicalInlineSourceIssue(candidate.content)
        : REVIEW_CARRIER_ISSUE;
    if (paragraph !== null) break;
  }
  return { structure, paragraph };
};

// An operation leaves every block it does not name the same object, and never
// changes a block in place, so each block is judged once. Typing then judges
// the one block it edits instead of every paragraph of the story.
const blockIssues = new WeakMap<BlockContent, BlockIssues>();
const issuesOf = (block: BlockContent): BlockIssues => {
  let issues = blockIssues.get(block);
  if (issues === undefined) {
    issues = issuesIn(block);
    blockIssues.set(block, issues);
  }
  return issues;
};

/**
 * Why the session cannot hold these stories of `document`, or `null`: the
 * first block it cannot hold, else the first paragraph whose content it
 * cannot hold, in each story in turn.
 */
export const unsupportedSeedReason = (
  document: Document,
  stories = documentStories(document),
): string | null => {
  const commentIssue = commentDocumentIssue(document, "all");
  if (commentIssue !== undefined) return commentIssue;
  for (const story of stories) {
    const body = findStoryBody(document, story);
    if (
      story !== OP_STORIES.MAIN &&
      (story.kind === "footnote" || story.kind === "endnote")
    ) {
      const notes =
        story.kind === "footnote"
          ? document.package.footnotes
          : document.package.endnotes;
      const note = notes?.find(({ id }) => id === story.id);
      if (note?.noteType !== undefined && note.noteType !== "normal") continue;
    }
    const content = body?.content;
    if (!content || content.length === 0)
      return "Canonical sessions require a nonempty paragraph story.";
    const issues = content.map(issuesOf);
    const structure = issues.find((issue) => issue.structure !== null);
    if (structure !== undefined) return structure.structure;
    const paragraph = issues.find((issue) => issue.paragraph !== null);
    if (paragraph !== undefined) return paragraph.paragraph;
  }
  return null;
};

/** Whether a story is a separator or continuation note, which the session never holds as a story. */
const skippedNote = (document: Document, story: OpStory): boolean => {
  if (
    story === OP_STORIES.MAIN ||
    (story.kind !== "footnote" && story.kind !== "endnote")
  )
    return false;
  const notes =
    story.kind === "footnote"
      ? document.package.footnotes
      : document.package.endnotes;
  const note = notes?.find(({ id }) => id === story.id);
  return note?.noteType !== undefined && note.noteType !== "normal";
};

/** The blocks of a story an edit replaced, and the ones it put in their place. */
export const changedWindow = (
  before: readonly BlockContent[],
  after: readonly BlockContent[],
): { removed: readonly BlockContent[]; added: readonly BlockContent[] } => {
  let prefix = 0;
  while (
    prefix < before.length &&
    prefix < after.length &&
    before[prefix] === after[prefix]
  )
    prefix += 1;
  let suffix = 0;
  while (
    suffix < before.length - prefix &&
    suffix < after.length - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  )
    suffix += 1;
  return {
    removed: before.slice(prefix, before.length - suffix),
    added: after.slice(prefix, after.length - suffix),
  };
};

const sameSequence = (left: readonly string[], right: readonly string[]) =>
  left.length === right.length &&
  left.every((entry, index) => entry === right[index]);

type UnsupportedEditOptions = {
  /** A document the session holds: one {@link unsupportedSeedReason} accepts. */
  previous: Document;
  document: Document;
  /** Every story whose blocks differ from `previous`'s. */
  stories: readonly OpStory[];
};

/**
 * {@link unsupportedSeedReason} of an edit's result, judging only what the
 * edit changed. Every block the edit left was judged with `previous`, which
 * holds none the session cannot hold, so only the blocks between a story's
 * kept prefix and suffix are judged. The comment verdict reads only the
 * comments and the order, kind and id of their anchors: while the comments
 * and the stories are the ones `previous` had and each story's replaced
 * blocks carried the anchors its new blocks carry, it stays as it was.
 */
export const unsupportedEditReason = ({
  previous,
  document,
  stories,
}: UnsupportedEditOptions): string | null => {
  const windows = stories.map((story) => ({
    story,
    content: findStoryBody(document, story)?.content,
    ...changedWindow(
      findStoryBody(previous, story)?.content ?? [],
      findStoryBody(document, story)?.content ?? [],
    ),
  }));
  const anchorsKept =
    document.package.document.comments === previous.package.document.comments &&
    documentStories(document).length === documentStories(previous).length &&
    windows.every(({ removed, added }) =>
      sameSequence(
        commentAnchorSequence(removed),
        commentAnchorSequence(added),
      ),
    );
  if (!anchorsKept) {
    const commentIssue = commentDocumentIssue(document, "all");
    if (commentIssue !== undefined) return commentIssue;
  }
  for (const { story, content, added } of windows) {
    if (skippedNote(document, story)) continue;
    if (!content || content.length === 0)
      return "Canonical sessions require a nonempty paragraph story.";
    const issues = added.map(issuesOf);
    const structure = issues.find((issue) => issue.structure !== null);
    if (structure !== undefined) return structure.structure;
    const paragraph = issues.find((issue) => issue.paragraph !== null);
    if (paragraph !== undefined) return paragraph.paragraph;
  }
  return null;
};

/**
 * The same verdict from one walk of every story, judging nothing once: the
 * reference the block judgments above must equal.
 */
export const unsupportedSeedReasonWholeDocument = (
  document: Document,
  stories = documentStories(document),
): string | null => {
  const commentIssue = commentDocumentIssue(document, "all");
  if (commentIssue !== undefined) return commentIssue;
  for (const story of stories) {
    const body = findStoryBody(document, story);
    if (
      story !== OP_STORIES.MAIN &&
      (story.kind === "footnote" || story.kind === "endnote")
    ) {
      const notes =
        story.kind === "footnote"
          ? document.package.footnotes
          : document.package.endnotes;
      const note = notes?.find(({ id }) => id === story.id);
      if (note?.noteType !== undefined && note.noteType !== "normal") continue;
    }
    const content = body?.content;
    if (!content || content.length === 0)
      return "Canonical sessions require a nonempty paragraph story.";
    const paragraphs: Paragraph[] = [];
    const collect = (blocks: readonly BlockContent[]): string | null => {
      for (const block of blocks) {
        if (block.type === "paragraph") {
          paragraphs.push(block);
          continue;
        }
        // Range markers may open or close between blocks; they hold no text.
        if (
          block.type === "commentRangeStart" ||
          block.type === "commentRangeEnd" ||
          block.type === "bookmarkStart" ||
          block.type === "bookmarkEnd" ||
          block.type === "moveFromRangeStart" ||
          block.type === "moveFromRangeEnd" ||
          block.type === "moveToRangeStart" ||
          block.type === "moveToRangeEnd"
        )
          continue;
        // Opaque markup the session keeps as it found it and never edits inside.
        if (block.type === "preservedBlock") continue;
        // A block content control holds its own block list, as a cell does.
        if (block.type === "blockSdt") {
          const issue = collect(block.content);
          if (issue !== null) return issue;
          continue;
        }
        if (block.type !== "table")
          return "Canonical sessions currently require paragraphs, tables and supported inline atoms.";
        for (const row of block.rows) {
          for (const cell of row.cells) {
            const issue = collect(cell.content);
            if (issue !== null) return issue;
          }
        }
      }
      return null;
    };
    const blockIssue = collect(content);
    if (blockIssue !== null) return blockIssue;
    for (const paragraph of paragraphs) {
      if (paragraph.reviewCarrier !== undefined)
        return "Canonical sessions currently require plain paragraphs and supported inline atoms.";
      const issue = canonicalInlineSourceIssue(paragraph.content);
      if (issue !== null) return issue;
    }
  }
  return null;
};
