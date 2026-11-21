/**
 * How a pending join reads: the paragraphs whose breaks are pending removal
 * run into the paragraph the join leaves once accepted. Both engines place a
 * block beside such a paragraph by this reading, the default one as its batch
 * runs and the canonical session against the document its batch read.
 */

import type { Node as PMNode } from "prosemirror-model";

import type { FolioDocumentOperation } from "../document-operations";
import { endOfParagraphWithItsTextBoxes } from "../prosemirror/anchoredTextBoxes";
import { paragraphJoinAfter } from "../prosemirror/commands/paragraphBookmarkJoin";
import { isBlockBoundaryNode } from "../prosemirror/keptRangeBoundaries";
import { paragraphLeftAfter } from "../prosemirror/paragraphPropertyCarry";
import { paragraphGoes, runInCellPending } from "../prosemirror/tableRunIn";
import { isZeroWidthAnchor } from "../prosemirror/zeroWidthAnchors";
import { buildCleanBlockText } from "./clean-text";
import type { NoteReferenceLabels } from "./note-references";
import type { FolioAIEditSkippedOperation } from "./types";

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
 * The paragraph whose break, once gone, runs it into `position`, the start of
 * a paragraph or a place between blocks: the paragraph before it, past the
 * typed range boundaries between them, when its break is pending removal.
 * Into a paragraph, the two must join as `paragraphJoinAfter` reads it, the
 * rule `paragraphLeftAfter` follows forward, so a join reads the same in
 * either direction. Null when nothing runs into `position`.
 */
const joinedFromBefore = (doc: PMNode, position: number): number | null => {
  const $place = doc.resolve(position);
  let end = position;
  let index = $place.index() - 1;
  while (index >= 0 && isBlockBoundaryNode($place.parent.child(index))) {
    end -= $place.parent.child(index).nodeSize;
    index -= 1;
  }
  const previous = index < 0 ? null : $place.parent.child(index);
  if (
    previous?.type.name !== "paragraph" ||
    !isDeletedPPrMark(previous.attrs["pPrMark"])
  )
    return null;
  if (
    doc.nodeAt(position)?.type.name === "paragraph" &&
    paragraphJoinAfter(doc, end)?.nextPosition !== position
  )
    return null;
  return end - previous.nodeSize;
};

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
  for (
    let previous = joinedFromBefore(doc, paragraphPos);
    previous !== null;
    previous = joinedFromBefore(doc, previous)
  )
    if (keepsWords({ doc, paragraphPos: previous, noteReferences }))
      return true;
  return false;
};

/**
 * The first paragraph a pending join runs into the table at `tablePos`, its
 * words joining the table's first cell once accepted, or null when no join
 * runs words into it or the cell they would join is itself pending. A block
 * put before the table goes before this paragraph, as one put before any
 * paragraph of a join goes before its first (`pendingJoinParagraphs`): in
 * the gap between the join and the table it would take the joined words.
 */
export const firstParagraphJoiningTable = ({
  doc,
  paragraphPos: tablePos,
  noteReferences,
}: PendingJoinBoundsOptions): number | null => {
  const table = doc.nodeAt(tablePos);
  if (
    table?.type.spec["tableRole"] !== "table" ||
    runInCellPending(table) ||
    !joinRunsWordsInto({ doc, paragraphPos: tablePos, noteReferences })
  )
    return null;
  let first: number | null = null;
  for (
    let previous = joinedFromBefore(doc, tablePos);
    previous !== null;
    previous = joinedFromBefore(doc, previous)
  )
    if (joinsIn({ doc, paragraphPos: previous, noteReferences }))
      first = previous;
  return first;
};

/** The first and the last paragraph of a pending join, by node position. */
export type PendingJoinParagraphs = { first: number; last: number };

/**
 * The paragraphs a pending join makes one of with the paragraph at
 * `paragraphPos`, read as accepting every change leaves them: a block put
 * before it goes before `first`, and one put after it goes after `last`. A
 * join runs the words of each paragraph whose break is pending removal into
 * the paragraph it leaves, so a block put between them would take those
 * words once accepted, as the same edits applied directly never do. Where a
 * paragraph runs into one the join leaves that keeps words of its own, a
 * block after any paragraph of the join goes after the paragraph it leaves,
 * and a block before any of them goes before the first paragraph the join
 * takes in, empty ones included. Both directions cross the typed range
 * boundaries a join crosses (`joinedFromBefore`). A join into a paragraph
 * with no words left (an emptied final carrier, or one going whole) runs
 * nothing together, nor does a paragraph going whole with no join running
 * through it: blocks keep their places.
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
  // Back from the paragraph the join leaves, so every paragraph of the join
  // reads the same first one, whichever of them the block is put beside.
  let before = paragraphPos;
  if (
    paragraph?.type.name === "paragraph" &&
    (after !== paragraphPos || keepsWords(options))
  ) {
    for (
      let previous = joinedFromBefore(doc, after);
      previous !== null;
      previous = joinedFromBefore(doc, previous)
    )
      if (
        previous < before &&
        joinsIn({ doc, paragraphPos: previous, noteReferences })
      )
        before = previous;
  }
  return { first: before, last: after };
};

/** Where a paragraph put after one whose break is pending removal goes. */
export type PendingBreakInsertion =
  /**
   * Right after that paragraph, with its properties: both it and the
   * paragraph its join leaves keep words once accepted, and accepting runs
   * its words into the new paragraph while the other stays its own, as a
   * paragraph typed at its end records it.
   */
  | { type: "afterParagraph" }
  /**
   * After `survivor`, the paragraph the join leaves, with that one's
   * properties, so accepting leaves one paragraph there and the new one
   * after it, as the same edits applied directly do: the paragraph keeps
   * words but the one its join leaves keeps none (a last-paragraph deletion
   * took it over), or the paragraph keeps none and reads as one paragraph
   * with the one its join leaves (`pendingJoinParagraphs`).
   */
  | { type: "afterJoin"; survivor: number };

/**
 * The place of a paragraph put after the paragraph at `paragraphPos`
 * (`PendingBreakInsertion`). A paragraph whose break stays takes it right
 * after itself.
 */
export const insertionAfterPendingBreak = (
  options: PendingJoinBoundsOptions,
): PendingBreakInsertion => {
  const { doc, paragraphPos, noteReferences } = options;
  if (!isDeletedPPrMark(doc.nodeAt(paragraphPos)?.attrs["pPrMark"]))
    return { type: "afterParagraph" };
  if (!keepsWords(options)) {
    const { last } = pendingJoinParagraphs(options);
    return last === paragraphPos
      ? { type: "afterParagraph" }
      : { type: "afterJoin", survivor: last };
  }
  const survivor = paragraphLeftAfter({ doc, paragraphPos });
  return survivor === null ||
    keepsWords({ doc, paragraphPos: survivor, noteReferences })
    ? { type: "afterParagraph" }
    : { type: "afterJoin", survivor };
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

/** A refusal both engines give beside a pending paragraph mark deletion. */
export type PendingRefusal = Omit<FolioAIEditSkippedOperation, "id">;

/** An insertion a pending join would run into the words before it or a table after it. */
export const PENDING_JOIN_INSERTION_REFUSAL: PendingRefusal = {
  reason: "pendingParagraphMarkDeletion",
  message:
    "Accepting the pending paragraph mark deletion before this place would run the paragraph together with the words before it or the table after it. Resolve that change first, or insert elsewhere.",
};

/**
 * A public operation that edits a paragraph's properties waits while the
 * paragraph's mark is pending deletion, in both engines: accepting the
 * deletion gives the joined paragraph the next paragraph's properties, so
 * the edit would not survive it. A merge carries its own result's properties
 * (`mergedParagraphProperties`). The refusal is the rule, not a gap.
 */
export const PENDING_PARAGRAPH_MARK_PROPERTY_REFUSAL: PendingRefusal = {
  reason: "pendingParagraphMarkDeletion",
  message:
    "Resolve the pending paragraph mark deletion before editing its properties, or provide mergedParagraphProperties on the merge operation for atomic intent.",
};

/** Whether every inline the paragraph holds, anchors aside, is pending deletion. */
export const holdsOnlyDeletedContent = (paragraph: PMNode): boolean => {
  let whollyDeleted = true;
  paragraph.descendants((node) => {
    if (!whollyDeleted) {
      return false;
    }
    if (!node.isInline || isZeroWidthAnchor(node)) {
      return true;
    }
    whollyDeleted = node.marks.some((mark) => mark.type.name === "deletion");
    return false;
  });
  return whollyDeleted;
};

type PendingJoinInsertionOptions = PendingJoinBoundsOptions & {
  insertedBlock: "paragraph" | "table";
};
/**
 * Whether the join a pending mark deletion makes stops at a table after the
 * paragraph, so a block placed there would run into it.
 */
export const joinStopsAtTableAfter = (
  options: PendingJoinInsertionOptions,
): boolean => {
  if (!joinRunsWordsInto(options)) return false;
  const $place = options.doc.resolve(options.paragraphPos);
  const before = $place.nodeBefore;
  // A deleted paragraph stood at the end of this join: accepting its
  // removal keeps the incoming paragraph's break, so its words stop here.
  // An inserted paragraph instead extends the join past that stopping mark.
  if (
    options.insertedBlock === "table" &&
    before !== null &&
    paragraphGoes(before, "accept")
  )
    return false;
  for (let index = $place.index(); index < $place.parent.childCount; index++) {
    const next = $place.parent.child(index);
    if (next.type.name === "paragraph") return false;
    if (next.type.spec["tableRole"] === "table") {
      return (
        !runInCellPending(next) ||
        (before !== null && holdsOnlyDeletedContent(before))
      );
    }
  }
  return false;
};

/** A paragraph that goes whole once its pending deletions are accepted. */
export const PENDING_WHOLE_PARAGRAPH_REFUSAL: PendingRefusal = {
  reason: "pendingDeletion",
};

/** What an operation that places no block asks of the paragraph it names. */
export type PendingMarkEdit = "paragraphProperties" | "wholeText" | "inline";

/** What an operation of this type asks of the paragraph it names. */
export const pendingMarkEditOf = (
  type: FolioDocumentOperation["type"],
): PendingMarkEdit => {
  if (type === "setBlockParagraphProperties") return "paragraphProperties";
  if (type === "replaceBlock") return "wholeText";
  return "inline";
};

export type PendingMarkEditOptions = {
  edit: PendingMarkEdit;
  doc: PMNode;
  paragraphPos: number;
  paragraph: PMNode;
  /** The paragraph's accepted reading. */
  cleanText: string;
};

/** The shared property boundary, independent of the paragraph's representation. */
export const pendingParagraphPropertiesRefusal = (
  mark: unknown,
): PendingRefusal | undefined =>
  isDeletedPPrMark(mark) ? PENDING_PARAGRAPH_MARK_PROPERTY_REFUSAL : undefined;

/**
 * The refusal an operation that places no block meets on a paragraph whose
 * mark is pending deletion, by the rules the default engine applies: its
 * properties wait for the change to be resolved, and so does a replacement of
 * a paragraph that goes whole; inline edits apply.
 */
export const pendingMarkEditRefusal = ({
  edit,
  doc,
  paragraphPos,
  paragraph,
  cleanText,
}: PendingMarkEditOptions): PendingRefusal | undefined => {
  if (edit === "paragraphProperties")
    return pendingParagraphPropertiesRefusal(paragraph.attrs["pPrMark"]);
  if (
    edit === "wholeText" &&
    isPendingDeletion(doc, paragraphPos, paragraph, cleanText)
  )
    return PENDING_WHOLE_PARAGRAPH_REFUSAL;
  return undefined;
};
