/**
 * Apply a document operation and record its exact inverse.
 *
 * `applyDocumentOp` is total: it returns the new document or a typed refusal,
 * and a refused operation changes nothing. It is deterministic: no clock, no
 * randomness and no counter outside its inputs, so equal documents and equal
 * operations give equal results in any realm. It is local: every block the
 * operation does not name is the same object in the output as in the input,
 * and so is every package part other than the story it edits.
 *
 * The inverse is captured from the state the operation replaced rather than
 * derived from the operation afterwards. It holds the original records, with
 * every field the model carries (unmodelled attributes, captured markup,
 * tracked changes, range markers), so applying it restores the document
 * structurally, and the restored blocks are the original objects.
 */

import { Result } from "better-result";

import type { Document, Paragraph, ParagraphFormatting } from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import {
  type ParagraphLocation,
  replaceParagraphs,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { structurallyEqual } from "./equality";
import { deleteInContent, insertTextInContent, patchRunsInContent, splitContent } from "./inline";
import { paragraphLogicalText } from "./offsets";
import { applyFormattingPatch } from "./patch";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  type DeleteRangeOp,
  type DocumentOp,
  type InsertTextOp,
  type JoinBlocksOp,
  type OpStory,
  type ReplaceBlocksOp,
  type SetParagraphPropsOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  type TextPosition,
  type TouchedBlocks,
} from "./types";

/** A document an operation produced, the operations that undo it, and what it changed. */
export type AppliedDocumentOp = {
  document: Document;
  /** Apply in order to restore the input document. */
  inverse: readonly DocumentOp[];
  touched: TouchedBlocks;
};

type Applied = Result<AppliedDocumentOp, DocumentOpRefusal>;

const UNTOUCHED: TouchedBlocks = Object.freeze({ modified: [], inserted: [], removed: [] });

const unchanged = (document: Document): Applied =>
  Result.ok({ document, inverse: [], touched: UNTOUCHED });

const refusal = (
  op: DocumentOp,
  reason: DocumentOpRefusalReason,
  message: string,
): DocumentOpRefusal => new DocumentOpRefusal({ message, reason, opType: op.type });

const refuse = (op: DocumentOp, reason: DocumentOpRefusalReason, message: string): Applied =>
  Result.err(refusal(op, reason, message));

/** `ST_LongHexNumber`: exactly eight hex digits. */
const LONG_HEX_NUMBER_PATTERN = /^[0-9A-Fa-f]{8}$/u;
/** `w14:paraId` values are below this bound, and zero means "no id". */
const PARA_ID_EXCLUSIVE_BOUND = 0x80_00_00_00;
/** Characters `insertText` does not carry: each is an inline atom of its own. */
const NON_TEXT_CHARACTER_PATTERN = /[\t\n\r]/u;

const isHighSurrogate = (code: number): boolean => code >= 0xd8_00 && code <= 0xdb_ff;
const isLowSurrogate = (code: number): boolean => code >= 0xdc_00 && code <= 0xdf_ff;

const findParagraph = (
  op: DocumentOp,
  paragraphs: readonly ParagraphLocation[],
  blockId: string,
): Result<ParagraphLocation, DocumentOpRefusal> => {
  const matches = paragraphs.filter(({ paragraph }) => paragraph.paraId === blockId);
  const [match] = matches;
  if (match === undefined) {
    return Result.err(
      refusal(op, DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, `No paragraph is ${blockId}.`),
    );
  }
  if (matches.length > 1) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.AMBIGUOUS_BLOCK_ID,
        `${matches.length} paragraphs are ${blockId}.`,
      ),
    );
  }
  return Result.ok(match);
};

/** Why `offset` is not a position in a paragraph whose logical text is `text`. */
const positionProblem = (text: string, offset: number): DocumentOpRefusalReason | undefined => {
  if (!Number.isInteger(offset) || offset < 0 || offset > text.length) {
    return DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET;
  }
  if (
    offset > 0 &&
    offset < text.length &&
    isHighSurrogate(text.charCodeAt(offset - 1)) &&
    isLowSurrogate(text.charCodeAt(offset))
  ) {
    return DOCUMENT_OP_REFUSAL_REASONS.SPLITS_SURROGATE_PAIR;
  }
  return undefined;
};

type LocatedRange = { story: OpStory; location: ParagraphLocation; from: number; to: number };

const locateRange = (
  document: Document,
  op: DocumentOp,
  from: TextPosition,
  to: TextPosition,
): Result<LocatedRange, DocumentOpRefusal> => {
  if (from.story !== to.story || from.blockId !== to.blockId) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.CROSS_BLOCK_RANGE,
        "A range starts and ends in one paragraph.",
      ),
    );
  }
  const located = findParagraph(op, storyParagraphs(storyBody(document, from.story)), from.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const text = paragraphLogicalText(located.value.paragraph);
  const problem =
    positionProblem(text, from.offset) ??
    positionProblem(text, to.offset) ??
    (from.offset > to.offset ? DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET : undefined);
  if (problem !== undefined) {
    return Result.err(
      refusal(op, problem, `[${from.offset}, ${to.offset}) is not a range of ${from.blockId}.`),
    );
  }
  return Result.ok({
    story: from.story,
    location: located.value,
    from: from.offset,
    to: to.offset,
  });
};

const locatePosition = (
  document: Document,
  op: DocumentOp,
  at: TextPosition,
): Result<ParagraphLocation, DocumentOpRefusal> => {
  const located = findParagraph(op, storyParagraphs(storyBody(document, at.story)), at.blockId);
  if (located.isErr()) {
    return located;
  }
  const problem = positionProblem(paragraphLogicalText(located.value.paragraph), at.offset);
  if (problem !== undefined) {
    return Result.err(refusal(op, problem, `${at.offset} is not a position in ${at.blockId}.`));
  }
  return located;
};

const touchedBetween = (
  before: readonly Paragraph[],
  after: readonly Paragraph[],
): TouchedBlocks => {
  const beforeIds = before.flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId]));
  const afterIds = after.flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId]));
  const afterSet = new Set(afterIds);
  const beforeSet = new Set(beforeIds);
  return {
    modified: beforeIds.filter((id) => afterSet.has(id)),
    inserted: afterIds.filter((id) => !beforeSet.has(id)),
    removed: beforeIds.filter((id) => !afterSet.has(id)),
  };
};

type ReplacedOptions = {
  document: Document;
  story: OpStory;
  at: ParagraphLocation;
  before: readonly Paragraph[];
  after: readonly Paragraph[];
};

/** Put `after` where `before` stands and record the replacement that undoes it. */
const replaced = ({ document, story, at, before, after }: ReplacedOptions): Applied =>
  Result.ok({
    document: replaceParagraphs({ document, story, at, count: before.length, replacement: after }),
    inverse: [{ type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS, story, expected: after, blocks: before }],
    touched: touchedBetween(before, after),
  });

const replacedOne = (
  document: Document,
  story: OpStory,
  at: ParagraphLocation,
  paragraph: Paragraph,
): Applied =>
  paragraph === at.paragraph
    ? unchanged(document)
    : replaced({ document, story, at, before: [at.paragraph], after: [paragraph] });

const withContent = (paragraph: Paragraph, content: Paragraph["content"]): Paragraph =>
  content === paragraph.content ? paragraph : { ...paragraph, content };

const insertText = (document: Document, op: InsertTextOp): Applied => {
  if (
    op.text === "" ||
    NON_TEXT_CHARACTER_PATTERN.test(op.text) ||
    hasIllegalXmlCharacters(op.text)
  ) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT, "The text cannot be inserted.");
  }
  const located = locatePosition(document, op, op.at);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  const content = insertTextInContent(paragraph.content, op.at.offset, op.text, op.runProps);
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION,
      `${op.at.offset} in ${op.at.blockId} is inside tracked-removed content.`,
    );
  }
  return replacedOne(document, op.at.story, located.value, withContent(paragraph, content));
};

const deleteRange = (document: Document, op: DeleteRangeOp): Applied => {
  const located = locateRange(document, op, op.from, op.to);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { story, location, from, to } = located.value;
  if (from === to) {
    return unchanged(document);
  }
  const { paragraph } = location;
  return replacedOne(
    document,
    story,
    location,
    withContent(paragraph, deleteInContent(paragraph.content, from, to)),
  );
};

const setRunProps = (document: Document, op: SetRunPropsOp): Applied => {
  const located = locateRange(document, op, op.from, op.to);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { story, location, from, to } = located.value;
  const { paragraph } = location;
  return replacedOne(
    document,
    story,
    location,
    withContent(paragraph, patchRunsInContent(paragraph.content, from, to, op.patch)),
  );
};

const withParagraphFormatting = (
  paragraph: Paragraph,
  formatting: ParagraphFormatting | undefined,
): Paragraph => {
  const next: Paragraph = { ...paragraph };
  if (formatting === undefined || Object.keys(formatting).length === 0) {
    delete next.formatting;
  } else {
    next.formatting = formatting;
  }
  return next;
};

const setParagraphProps = (document: Document, op: SetParagraphPropsOp): Applied => {
  const located = findParagraph(op, storyParagraphs(storyBody(document, op.story)), op.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  const formatting = applyFormattingPatch(paragraph.formatting, op.patch);
  if (structurallyEqual(formatting ?? {}, paragraph.formatting ?? {})) {
    return unchanged(document);
  }
  return replacedOne(
    document,
    op.story,
    located.value,
    withParagraphFormatting(paragraph, formatting),
  );
};

const newBlockIdProblem = (
  paragraphs: readonly ParagraphLocation[],
  id: string,
): DocumentOpRefusalReason | undefined => {
  if (!LONG_HEX_NUMBER_PATTERN.test(id)) {
    return DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID;
  }
  const value = Number.parseInt(id, 16);
  if (value === 0 || value >= PARA_ID_EXCLUSIVE_BOUND) {
    return DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID;
  }
  const normalized = id.toUpperCase();
  return paragraphs.some(({ paragraph }) => paragraph.paraId?.toUpperCase() === normalized)
    ? DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION
    : undefined;
};

const splitBlock = (document: Document, op: SplitBlockOp): Applied => {
  const paragraphs = storyParagraphs(storyBody(document, op.at.story));
  const idProblem = newBlockIdProblem(paragraphs, op.newBlockId);
  if (idProblem !== undefined) {
    return refuse(op, idProblem, `${op.newBlockId} cannot name a new paragraph.`);
  }
  const located = locatePosition(document, op, op.at);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  const halves = splitContent(paragraph.content, op.at.offset);
  if (halves === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.SPLITS_IDENTIFIED_CONTAINER,
      `${op.at.offset} in ${op.at.blockId} is inside a tracked change or content control.`,
    );
  }
  const first: Paragraph = { ...paragraph, content: halves.left };
  delete first.sectionProperties;
  delete first.pPrMark;
  const second = withParagraphFormatting(
    { type: "paragraph", paraId: op.newBlockId, content: halves.right },
    op.newProps ?? paragraph.formatting,
  );
  if (paragraph.sectionProperties !== undefined) {
    second.sectionProperties = paragraph.sectionProperties;
  }
  if (paragraph.pPrMark !== undefined) {
    second.pPrMark = paragraph.pPrMark;
  }
  return replaced({
    document,
    story: op.at.story,
    at: located.value,
    before: [paragraph],
    after: [first, second],
  });
};

const joinBlocks = (document: Document, op: JoinBlocksOp): Applied => {
  const paragraphs = storyParagraphs(storyBody(document, op.story));
  const first = findParagraph(op, paragraphs, op.blockId);
  if (first.isErr()) {
    return Result.err(first.error);
  }
  const next = findParagraph(op, paragraphs, op.nextBlockId);
  if (next.isErr()) {
    return Result.err(next.error);
  }
  if (
    !sameBlockList(first.value.list, next.value.list) ||
    next.value.index !== first.value.index + 1
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      `${op.nextBlockId} does not directly follow ${op.blockId}.`,
    );
  }
  const leading = first.value.paragraph;
  const trailing = next.value.paragraph;
  if (leading.sectionProperties !== undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      `${op.blockId} ends a section.`,
    );
  }
  const joined: Paragraph = { ...leading, content: [...leading.content, ...trailing.content] };
  delete joined.pPrMark;
  if (trailing.sectionProperties !== undefined) {
    joined.sectionProperties = trailing.sectionProperties;
  }
  if (trailing.pPrMark !== undefined) {
    joined.pPrMark = trailing.pPrMark;
  }
  return replaced({
    document,
    story: op.story,
    at: first.value,
    before: [leading, trailing],
    after: [joined],
  });
};

const replaceBlocks = (document: Document, op: ReplaceBlocksOp): Applied => {
  const [firstExpected] = op.expected;
  if (firstExpected === undefined || op.blocks.length === 0) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "A replacement names the paragraphs it replaces and the ones it puts in their place.",
    );
  }
  const paragraphs = storyParagraphs(storyBody(document, op.story));
  const found: ParagraphLocation[] = [];
  for (const expected of op.expected) {
    if (expected.paraId === undefined) {
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "A replaced paragraph has no id.",
      );
    }
    const located = findParagraph(op, paragraphs, expected.paraId);
    if (located.isErr()) {
      return Result.err(located.error);
    }
    found.push(located.value);
  }
  const [start] = found;
  if (start === undefined) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST, "Nothing to replace.");
  }
  const contiguous = found.every(
    (location, offset) =>
      sameBlockList(location.list, start.list) && location.index === start.index + offset,
  );
  if (!contiguous) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      "The replaced paragraphs are not adjacent.",
    );
  }
  const stale = found.some(
    (location, offset) => !structurallyEqual(location.paragraph, op.expected[offset]),
  );
  if (stale) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STALE, "The paragraphs to replace have changed.");
  }
  const replacedParagraphs = new Set(found.map(({ paragraph }) => paragraph));
  const remainingIds = new Set(
    paragraphs.flatMap(({ paragraph }) =>
      replacedParagraphs.has(paragraph) || paragraph.paraId === undefined ? [] : [paragraph.paraId],
    ),
  );
  const incomingIds = new Set<string>();
  for (const block of op.blocks) {
    if (block.paraId === undefined) {
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "A replacement paragraph has no id.",
      );
    }
    if (remainingIds.has(block.paraId) || incomingIds.has(block.paraId)) {
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        `${block.paraId} is already used in the story.`,
      );
    }
    incomingIds.add(block.paraId);
  }
  return replaced({
    document,
    story: op.story,
    at: start,
    before: found.map(({ paragraph }) => paragraph),
    after: op.blocks,
  });
};

/** Apply one operation. */
export const applyDocumentOp = (
  document: Document,
  op: DocumentOp,
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
      return insertText(document, op);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return deleteRange(document, op);
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return setRunProps(document, op);
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
      return setParagraphProps(document, op);
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return splitBlock(document, op);
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return joinBlocks(document, op);
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
      return replaceBlocks(document, op);
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

type TouchedState = "modified" | "inserted" | "removed";

const TOUCH_TRANSITIONS = {
  modified: { modified: "modified", inserted: "modified", removed: "removed" },
  inserted: { modified: "inserted", inserted: "inserted", removed: undefined },
  removed: { modified: "modified", inserted: "modified", removed: "removed" },
} as const satisfies Record<TouchedState, Record<TouchedState, TouchedState | undefined>>;

/**
 * Apply operations in order, atomically: the first refusal is returned and
 * the document is unchanged. The inverse undoes the whole list.
 */
export const applyDocumentOps = (
  document: Document,
  ops: readonly DocumentOp[],
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  let current = document;
  const inverses: (readonly DocumentOp[])[] = [];
  const touched = new Map<string, TouchedState>();
  for (const op of ops) {
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) {
      return applied;
    }
    current = applied.value.document;
    inverses.push(applied.value.inverse);
    for (const state of ["modified", "inserted", "removed"] as const) {
      for (const id of applied.value.touched[state]) {
        const previous = touched.get(id);
        const nextState = previous === undefined ? state : TOUCH_TRANSITIONS[previous][state];
        if (nextState === undefined) {
          touched.delete(id);
        } else {
          touched.set(id, nextState);
        }
      }
    }
  }
  const ids = (state: TouchedState): string[] =>
    [...touched].flatMap(([id, value]) => (value === state ? [id] : []));
  return Result.ok({
    document: current,
    inverse: inverses.toReversed().flat(),
    touched: { modified: ids("modified"), inserted: ids("inserted"), removed: ids("removed") },
  });
};
