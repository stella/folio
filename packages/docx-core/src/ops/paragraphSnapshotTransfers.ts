/**
 * Which authored paragraph property snapshots an edit carries over: a
 * modified paragraph the edit replaced keeps its original's snapshot while
 * its formatting stays the same.
 */
import type { BlockContent, Document, Paragraph } from "../model/document";
import { storyParagraphs } from "./blocks";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { idKey } from "./ids";
import { changedListBounds } from "./listSplice";
import {
  documentStories,
  findStoryBody,
  type StoryAddress,
  storyBody,
  type TextBoxStory,
} from "./stories";

type SnapshotTransfersOptions = {
  before: Document;
  edit: DocumentEdit;
  /** Text boxes the edit was applied in, read as stories of their own. */
  boxes?: readonly TextBoxStory[];
};

const changedStories = ({
  edit,
  boxes = [],
}: SnapshotTransfersOptions): StoryAddress[] => [
  ...documentStories(edit.document),
  ...boxes,
];

/** A snapshot to carry from `source` to `target`. */
export type SnapshotTransfer = { source: Paragraph; target: Paragraph };

const paragraphsOf = (blocks: readonly BlockContent[]): Paragraph[] =>
  storyParagraphs({ content: [...blocks] }).map(({ paragraph }) => paragraph);

const transfersBetween = (
  previous: readonly Paragraph[],
  next: readonly Paragraph[],
  modified: ReadonlySet<string>,
  transfers: SnapshotTransfer[],
): void => {
  const originals = new Map(
    previous.flatMap((paragraph) =>
      paragraph.paraId !== undefined && modified.has(idKey(paragraph.paraId))
        ? [[idKey(paragraph.paraId), paragraph] as const]
        : [],
    ),
  );
  if (originals.size === 0) return;
  for (const paragraph of next) {
    if (paragraph.paraId === undefined) continue;
    const original = originals.get(idKey(paragraph.paraId));
    if (
      original === undefined ||
      original === paragraph ||
      !structurallyEqual(original.formatting, paragraph.formatting)
    )
      continue;
    transfers.push({ source: original, target: paragraph });
  }
};

/**
 * The transfers an edit makes, read from the top-level blocks of each story
 * it changed that are not the very blocks `before` held at the same place
 * from either end. A paragraph of a kept block is its own original, and the
 * seed contract keeps paragraph ids unique, so no other paragraph is one.
 */
export const paragraphSnapshotTransfers = (
  options: SnapshotTransfersOptions,
): SnapshotTransfer[] => {
  const { before, edit } = options;
  const transfers: SnapshotTransfer[] = [];
  if (edit.touched.modified.length === 0) return transfers;
  const modified = new Set(edit.touched.modified.map(idKey));
  for (const story of changedStories(options)) {
    const previous = findStoryBody(before, story);
    const next = storyBody(edit.document, story);
    if (previous === undefined || previous === next) continue;
    const old = previous.content;
    const now = next.content;
    const { start, previousEnd, end } = changedListBounds(old, now);
    transfersBetween(
      paragraphsOf(old.slice(start, previousEnd)),
      paragraphsOf(now.slice(start, end)),
      modified,
      transfers,
    );
  }
  return transfers;
};

/** The transfers read from every paragraph of each story the edit changed. */
export const paragraphSnapshotTransfersReadingEveryParagraph = (
  options: SnapshotTransfersOptions,
): SnapshotTransfer[] => {
  const { before, edit } = options;
  const transfers: SnapshotTransfer[] = [];
  if (edit.touched.modified.length === 0) return transfers;
  const modified = new Set(edit.touched.modified.map(idKey));
  for (const story of changedStories(options)) {
    const previous = findStoryBody(before, story);
    const next = storyBody(edit.document, story);
    if (previous === undefined || previous === next) continue;
    transfersBetween(
      paragraphsOf(previous.content),
      paragraphsOf(next.content),
      modified,
      transfers,
    );
  }
  return transfers;
};
