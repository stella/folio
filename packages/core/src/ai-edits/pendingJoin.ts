/**
 * How a pending join reads: the paragraphs whose breaks are pending removal
 * run into the paragraph the join leaves once accepted. Both engines place a
 * block beside such a paragraph by this reading, the default one as its batch
 * runs and the canonical session against the document its batch read.
 */

import type { Node as PMNode } from "prosemirror-model";

import { endOfParagraphWithItsTextBoxes } from "../prosemirror/anchoredTextBoxes";
import { paragraphLeftAfter } from "../prosemirror/paragraphPropertyCarry";
import { buildCleanBlockText } from "./clean-text";
import type { NoteReferenceLabels } from "./note-references";

/** A paragraph mark that is a pending deletion, or a relocation's source. */
export const isDeletedPPrMark = (value: unknown): boolean =>
  typeof value === "object" &&
  value !== null &&
  "kind" in value &&
  (value.kind === "del" || value.kind === "moveFrom");

/** A table row's or cell's revision record. */
const hasRevisionMarker = (value: unknown): boolean =>
  typeof value === "object" && value !== null && "revisionId" in value;

/** Whether a table cell is pending deletion. */
const isDeletedCell = (cell: PMNode): boolean => {
  const marker: unknown = cell.attrs["cellMarker"];
  return (
    typeof marker === "object" &&
    marker !== null &&
    "kind" in marker &&
    marker.kind === "del"
  );
};

/**
 * Whether accepting every change removes the paragraph at `blockFrom` whole:
 * its paragraph mark and all of its text pending deletion (`cleanText`, the
 * accepted reading, is empty), or the table row or cell holding it pending
 * deletion. A reader lists such a paragraph as a blank block.
 *
 * Word keeps a paragraph whose mark is deleted joined to the next one: text
 * typed into it stays inserted there, and accepting runs it into the next
 * paragraph (or, in a deleted row, takes it away with the row). An operation
 * that asks for the paragraph to read something cannot get that without
 * rejecting the deletion, which is the reviewer's call, not a side effect.
 */
export const isPendingDeletion = (
  doc: PMNode,
  blockFrom: number,
  paragraph: PMNode,
  cleanText: string,
): boolean => {
  if (cleanText.length === 0 && isDeletedPPrMark(paragraph.attrs["pPrMark"])) {
    return true;
  }
  const $block = doc.resolve(blockFrom);
  for (let depth = $block.depth; depth > 0; depth -= 1) {
    const ancestor = $block.node(depth);
    const name = ancestor.type.name;
    if (name === "tableRow" && hasRevisionMarker(ancestor.attrs["trDel"])) {
      return true;
    }
    if (
      (name === "tableCell" || name === "tableHeader") &&
      isDeletedCell(ancestor)
    ) {
      return true;
    }
  }
  return false;
};

export type PendingJoinBoundsOptions = {
  doc: PMNode;
  paragraphPos: number;
  noteReferences: NoteReferenceLabels;
};

const keepsWords = ({
  doc,
  paragraphPos,
  noteReferences,
}: PendingJoinBoundsOptions): boolean => {
  const paragraph = doc.nodeAt(paragraphPos);
  if (paragraph?.type.name !== "paragraph") return false;
  const cleanText = buildCleanBlockText(paragraph, paragraphPos, {
    fieldResults: "text",
    noteReferences,
  }).text;
  return (
    cleanText.length > 0 &&
    !isPendingDeletion(doc, paragraphPos, paragraph, cleanText)
  );
};

/**
 * Whether accepting runs the paragraph into the one its pending join leaves:
 * it keeps words, or it is empty. One whose words are all pending deletion
 * goes whole instead.
 */
export const joinsIn = (options: PendingJoinBoundsOptions): boolean =>
  options.doc.nodeAt(options.paragraphPos)?.content.size === 0 ||
  keepsWords(options);

/**
 * Whether a pending join runs earlier words on to `paragraphPos`, the start of
 * a paragraph or a place between blocks: the paragraphs before it, back to one
 * that keeps words, all have their breaks pending removal.
 */
export const joinRunsWordsInto = ({
  doc,
  paragraphPos,
  noteReferences,
}: PendingJoinBoundsOptions): boolean => {
  const $place = doc.resolve(paragraphPos);
  let position = paragraphPos;
  for (let index = $place.index() - 1; index >= 0; index--) {
    const previous = $place.parent.child(index);
    position -= previous.nodeSize;
    if (
      previous.type.name !== "paragraph" ||
      !isDeletedPPrMark(previous.attrs["pPrMark"])
    ) {
      return false;
    }
    if (keepsWords({ doc, paragraphPos: position, noteReferences })) {
      return true;
    }
  }
  return false;
};

/** The first and the last paragraph of a pending join, by node position. */
export type PendingJoinParagraphs = { first: number; last: number };

/**
 * The paragraphs a pending join makes one of with the paragraph at
 * `paragraphPos`, read as accepting every change leaves them: a block put
 * before it goes before `first`, and one put after it goes after `last`. A join runs the words
 * of each paragraph whose break is pending removal into the paragraph it
 * leaves, so a block put between them would take those words once accepted,
 * as the same edits applied directly never do. Where a paragraph runs into
 * one the join leaves that keeps words of its own, a block after any
 * paragraph of the join goes after the paragraph it leaves, and a block
 * before that one goes before the first paragraph the join takes in, empty
 * ones included. A join into a paragraph with no words left (an emptied
 * final carrier, or one going whole) runs nothing together, nor does a
 * paragraph going whole with no join running through it: blocks keep their
 * places.
 */
export const pendingJoinParagraphs = (
  options: PendingJoinBoundsOptions,
): PendingJoinParagraphs => {
  const { doc, paragraphPos, noteReferences } = options;
  const paragraph = doc.nodeAt(paragraphPos);
  let after = paragraphPos;
  if (
    paragraph?.type.name === "paragraph" &&
    isDeletedPPrMark(paragraph.attrs["pPrMark"]) &&
    (joinsIn(options) || joinRunsWordsInto(options))
  ) {
    const survivor = paragraphLeftAfter({ doc, paragraphPos });
    if (
      survivor !== null &&
      keepsWords({ doc, paragraphPos: survivor, noteReferences })
    ) {
      after = survivor;
    }
  }
  let before = paragraphPos;
  if (paragraph?.type.name === "paragraph" && keepsWords(options)) {
    let start = paragraphPos;
    for (;;) {
      const previous = doc.resolve(start).nodeBefore;
      if (
        previous?.type !== paragraph.type ||
        !isDeletedPPrMark(previous.attrs["pPrMark"])
      ) {
        break;
      }
      start -= previous.nodeSize;
      if (joinsIn({ doc, paragraphPos: start, noteReferences })) {
        before = start;
      }
    }
  }
  return { first: before, last: after };
};

/**
 * Where a block goes before and after the paragraph at `paragraphPos`, read
 * as the joined paragraph a pending join makes of it (`pendingJoinParagraphs`):
 * before the first paragraph of the join, and after the last one with the
 * text boxes drawn in it.
 */
export const pendingJoinBounds = (
  options: PendingJoinBoundsOptions,
): { before: number; after: number } => {
  const { first, last } = pendingJoinParagraphs(options);
  return {
    before: first,
    after: endOfParagraphWithItsTextBoxes(options.doc, last),
  };
};
