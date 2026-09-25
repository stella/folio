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
 * The inverse is a list of ordinary operations, addressed by position like
 * the one it undoes, captured from the state the operation changed:
 *
 * | operation           | inverse                                                  |
 * | ------------------- | -------------------------------------------------------- |
 * | `insertText`        | `deleteRange` (+ `joinInline` for a run it cut)          |
 * | `insertContent`     | `deleteRange` (+ `joinInline` / `splitInline` repair)    |
 * | `deleteRange`       | `insertContent` with the removed slice                   |
 * | `splitInline`       | `joinInline`                                             |
 * | `joinInline`        | `splitInline`                                            |
 * | `setRunProps`       | `setRunProps` per stretch of prior values, `joinInline`  |
 * | `setParagraphProps` | `setParagraphProps` with the prior values                |
 * | `splitBlock`        | `joinBlocks`                                             |
 * | `joinBlocks`        | `splitBlock` with the second paragraph's fields          |
 * | `replaceBlocks`     | `replaceBlocks`                                          |
 *
 * Deletions carry the slice they remove as `expected`, so an inverse applied
 * to content that has changed since is refused as stale.
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
import { countIds, paragraphIdsIn } from "./ids";
import {
  cutAt,
  deleteBetween,
  emptySetSpelling,
  gapAfterInserted,
  type InsertionRepair,
  insertionInverse,
  insertSliceAt,
  insertTextInContent,
  joinAt,
  joinContent,
  patchedSet,
  patchRunsBetween,
  splitAt,
} from "./inline";
import {
  defaultInsertionGap,
  type Gap,
  isIdentifiedNode,
  spanningRecords,
  textsIn,
  zeroWidthLeavesAt,
} from "./leaves";
import { paragraphLength, paragraphLogicalText } from "./offsets";
import { priorValues } from "./patch";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import {
  DOCUMENT_OP_TYPES,
  type DeleteRangeOp,
  type DocumentOp,
  type InsertContentOp,
  type InsertTextOp,
  type JoinBlocksOp,
  type JoinInlineOp,
  type OpStory,
  type ReplaceBlocksOp,
  type SetParagraphPropsOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  type SplitInlineOp,
  type SplitParagraphFields,
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

/** `w14:paraId` reserves zero for "no id". */
const RESERVED_PARA_ID_PATTERN = /^0{8}$/u;
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

const locate = (
  document: Document,
  op: DocumentOp,
  story: OpStory,
  blockId: string,
): Result<ParagraphLocation, DocumentOpRefusal> =>
  findParagraph(op, storyParagraphs(storyBody(document, story)), blockId);

/** What a position that states no `zeroWidthBefore` assumes. */
type ZeroWidthDefault = "afterAll" | "beforeAll" | "insertion";

/** The gap a position names, or why it names none in this paragraph. */
const resolveGap = (
  paragraph: Paragraph,
  position: TextPosition,
  fallback: ZeroWidthDefault,
): Gap | DocumentOpRefusalReason => {
  const text = paragraphLogicalText(paragraph);
  const { offset, zeroWidthBefore } = position;
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
  const available = zeroWidthLeavesAt(paragraph.content, offset).length;
  if (zeroWidthBefore !== undefined) {
    return Number.isInteger(zeroWidthBefore) && zeroWidthBefore >= 0 && zeroWidthBefore <= available
      ? { offset, zeroWidthBefore }
      : DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET;
  }
  switch (fallback) {
    case "afterAll":
      return { offset, zeroWidthBefore: available };
    case "beforeAll":
      return { offset, zeroWidthBefore: 0 };
    case "insertion":
      return defaultInsertionGap(paragraph.content, offset);
    default: {
      const unreachable: never = fallback;
      return unreachable;
    }
  }
};

const isGap = (value: Gap | DocumentOpRefusalReason): value is Gap => typeof value === "object";

type LocatedRange = {
  story: OpStory;
  location: ParagraphLocation;
  from: Gap;
  to: Gap;
  /** The gaps enclose no leaf. */
  empty: boolean;
};

/**
 * A range of one paragraph. Absent `zeroWidthBefore` leaves the zero-width
 * leaves at both ends outside the range.
 */
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
  const located = locate(document, op, from.story, from.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  const start = resolveGap(paragraph, from, "afterAll");
  const end = resolveGap(paragraph, to, "beforeAll");
  const invalid = (reason: DocumentOpRefusalReason) =>
    Result.err(
      refusal(op, reason, `[${from.offset}, ${to.offset}) is not a range of ${from.blockId}.`),
    );
  if (!isGap(start)) return invalid(start);
  if (!isGap(end)) return invalid(end);
  const statesZeroWidth = from.zeroWidthBefore !== undefined || to.zeroWidthBefore !== undefined;
  if (
    start.offset > end.offset ||
    (start.offset === end.offset && start.zeroWidthBefore > end.zeroWidthBefore && statesZeroWidth)
  ) {
    return invalid(DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET);
  }
  const empty = start.offset === end.offset && start.zeroWidthBefore >= end.zeroWidthBefore;
  return Result.ok({ story: from.story, location: located.value, from: start, to: end, empty });
};

const locatePosition = (
  document: Document,
  op: DocumentOp,
  at: TextPosition,
  fallback: ZeroWidthDefault,
): Result<{ location: ParagraphLocation; gap: Gap }, DocumentOpRefusal> => {
  const located = locate(document, op, at.story, at.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const gap = resolveGap(located.value.paragraph, at, fallback);
  if (!isGap(gap)) {
    return Result.err(refusal(op, gap, `${at.offset} is not a position in ${at.blockId}.`));
  }
  return Result.ok({ location: located.value, gap });
};

const positionAt = (story: OpStory, blockId: string, gap: Gap): TextPosition => ({
  story,
  blockId,
  offset: gap.offset,
  zeroWidthBefore: gap.zeroWidthBefore,
});

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
  inverse: readonly DocumentOp[];
};

/** Put `after` where `before` stands. */
const replaced = ({ document, story, at, before, after, inverse }: ReplacedOptions): Applied =>
  Result.ok({
    document: replaceParagraphs({ document, story, at, count: before.length, replacement: after }),
    inverse,
    touched: touchedBetween(before, after),
  });

type EditedOptions = {
  document: Document;
  story: OpStory;
  at: ParagraphLocation;
  paragraph: Paragraph;
  inverse: readonly DocumentOp[];
};

const edited = ({ document, story, at, paragraph, inverse }: EditedOptions): Applied =>
  paragraph === at.paragraph
    ? unchanged(document)
    : replaced({ document, story, at, before: [at.paragraph], after: [paragraph], inverse });

const withContent = (
  paragraph: Paragraph,
  content: readonly Paragraph["content"][number][],
): Paragraph =>
  content === paragraph.content ? paragraph : { ...paragraph, content: [...content] };

const repairOps = (repair: InsertionRepair, at: TextPosition): DocumentOp[] => {
  switch (repair.kind) {
    case "none":
      return [];
    case "join":
      return [{ type: DOCUMENT_OP_TYPES.JOIN_INLINE, at, depth: repair.depth }];
    case "split":
      return [{ type: DOCUMENT_OP_TYPES.SPLIT_INLINE, at, depth: repair.depth }];
    default: {
      const unreachable: never = repair;
      return unreachable;
    }
  }
};

type InsertedOptions = {
  document: Document;
  op: DocumentOp;
  story: OpStory;
  location: ParagraphLocation;
  content: Paragraph["content"];
  start: Gap;
  end: Gap;
};

/** Commit an insertion, recording the deletion (and repair) that undoes it. */
const inserted = ({
  document,
  op,
  story,
  location,
  content,
  start,
  end,
}: InsertedOptions): Applied => {
  const { paragraph } = location;
  const blockId = paragraph.paraId ?? "";
  const inverse = insertionInverse(paragraph.content, content, start, end);
  if (inverse.touchesIdentified) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.SPLITS_IDENTIFIED_CONTAINER,
      `The insertion in ${blockId} would cut or merge a tracked change or content control.`,
    );
  }
  const from = positionAt(story, blockId, start);
  return edited({
    document,
    story,
    at: location,
    paragraph: withContent(paragraph, content),
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from,
        to: positionAt(story, blockId, end),
        expected: inverse.removed,
      },
      ...repairOps(inverse.repair, from),
    ],
  });
};

const insertText = (document: Document, op: InsertTextOp): Applied => {
  if (
    op.text === "" ||
    NON_TEXT_CHARACTER_PATTERN.test(op.text) ||
    hasIllegalXmlCharacters(op.text)
  ) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT, "The text cannot be inserted.");
  }
  const located = locatePosition(document, op, op.at, "afterAll");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location } = located.value;
  const { offset } = op.at;
  const content = insertTextInContent(location.paragraph.content, offset, op.text, op.runProps);
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION,
      `${offset} in ${op.at.blockId} is inside tracked-removed content.`,
    );
  }
  // Every zero-width leaf at the offset precedes the first inserted character.
  const start = { offset, zeroWidthBefore: zeroWidthLeavesAt(content, offset).length };
  const end = { offset: offset + op.text.length, zeroWidthBefore: 0 };
  return inserted({ document, op, story: op.at.story, location, content, start, end });
};

/** Paragraph ids in `incoming` that the package already has, or that `incoming` repeats. */
const idCollision = (
  existing: ReadonlyMap<string, number>,
  incoming: readonly string[],
): boolean => {
  const seen = new Set<string>();
  return incoming.some((id) => {
    const repeated = seen.has(id) || (existing.get(id) ?? 0) > 0;
    seen.add(id);
    return repeated;
  });
};

const insertContent = (document: Document, op: InsertContentOp): Applied => {
  const { slice } = op;
  if (slice.content.length === 0) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT, "The slice holds nothing.");
  }
  if (
    !Number.isInteger(slice.openStart) ||
    !Number.isInteger(slice.openEnd) ||
    slice.openStart < 0 ||
    slice.openEnd < 0
  ) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "Open depths are counts.");
  }
  if (textsIn(slice.content).some(hasIllegalXmlCharacters)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      "The slice holds text that cannot be written.",
    );
  }
  const incoming = paragraphIdsIn(slice.content);
  if (incoming.length > 0 && idCollision(countIds(paragraphIdsIn(document.package)), incoming)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "The slice holds a paragraph id the package already uses.",
    );
  }
  const located = locatePosition(document, op, op.at, "insertion");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location, gap } = located.value;
  const content = insertSliceAt(location.paragraph.content, gap, slice);
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      `The slice's open ends do not fit ${op.at.offset} in ${op.at.blockId}.`,
    );
  }
  return inserted({
    document,
    op,
    story: op.at.story,
    location,
    content,
    start: gap,
    end: gapAfterInserted(gap, slice.content),
  });
};

const deleteRange = (document: Document, op: DeleteRangeOp): Applied => {
  const located = locateRange(document, op, op.from, op.to);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { story, location, from, to, empty } = located.value;
  const { expected } = op;
  if (empty) {
    return expected === undefined || expected.content.length === 0
      ? unchanged(document)
      : refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STALE, "The range to delete is empty.");
  }
  const { paragraph } = location;
  const { content, removed } = deleteBetween(paragraph.content, from, to);
  if (expected !== undefined && !structurallyEqual(expected, removed)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `The range of ${op.from.blockId} holds other content than expected.`,
    );
  }
  if (removed.content.length === 0) {
    return unchanged(document);
  }
  return edited({
    document,
    story,
    at: location,
    paragraph: withContent(paragraph, content),
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at: positionAt(story, op.from.blockId, from),
        slice: removed,
      },
    ],
  });
};

const isDepth = (depth: number): boolean => Number.isInteger(depth) && depth >= 1;

const splitInline = (document: Document, op: SplitInlineOp): Applied => {
  if (!isDepth(op.depth)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A split cuts one level or more.",
    );
  }
  const located = locatePosition(document, op, op.at, "beforeAll");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location, gap } = located.value;
  const outcome = splitAt(location.paragraph.content, gap, op.depth);
  switch (outcome.kind) {
    case "tooShallow":
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        `Fewer than ${op.depth} records run across ${op.at.offset} in ${op.at.blockId}.`,
      );
    case "identified":
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.SPLITS_IDENTIFIED_CONTAINER,
        `The split in ${op.at.blockId} would cut a tracked change or content control.`,
      );
    case "split":
      return edited({
        document,
        story: op.at.story,
        at: location,
        paragraph: withContent(location.paragraph, outcome.content),
        inverse: [
          {
            type: DOCUMENT_OP_TYPES.JOIN_INLINE,
            at: positionAt(op.at.story, op.at.blockId, gap),
            depth: op.depth,
          },
        ],
      });
    default: {
      const unreachable: never = outcome;
      return unreachable;
    }
  }
};

const joinInline = (document: Document, op: JoinInlineOp): Applied => {
  if (!isDepth(op.depth)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A join merges one level or more.",
    );
  }
  const located = locatePosition(document, op, op.at, "beforeAll");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location, gap } = located.value;
  const content = joinAt(location.paragraph.content, gap, op.depth);
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      `The records meeting at ${op.at.offset} in ${op.at.blockId} cannot be merged.`,
    );
  }
  return edited({
    document,
    story: op.at.story,
    at: location,
    paragraph: withContent(location.paragraph, content),
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.SPLIT_INLINE,
        at: positionAt(op.at.story, op.at.blockId, gap),
        depth: op.depth,
      },
    ],
  });
};

const setRunProps = (document: Document, op: SetRunPropsOp): Applied => {
  const located = locateRange(document, op, op.from, op.to);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { story, location, from, to, empty } = located.value;
  if (empty) {
    return unchanged(document);
  }
  const { paragraph } = location;
  const patched = patchRunsBetween(paragraph.content, from, to, op.patch, op.whenEmpty);
  if (patched === undefined) {
    return unchanged(document);
  }
  const blockId = op.from.blockId;
  const inverse: DocumentOp[] = patched.restoring.map((span) => ({
    type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
    from: positionAt(story, blockId, span.from),
    to: positionAt(story, blockId, span.to),
    patch: span.patch,
    whenEmpty: span.whenEmpty,
  }));
  // Runs cut at either end merge back once their values are restored.
  for (const gap of [from, to]) {
    const cut =
      spanningRecords(paragraph.content, [gap], 0, 1).length -
      spanningRecords(patched.content, [gap], 0, 1).length;
    if (cut > 0) {
      inverse.push({
        type: DOCUMENT_OP_TYPES.JOIN_INLINE,
        at: positionAt(story, blockId, gap),
        depth: cut,
      });
    }
  }
  return edited({
    document,
    story,
    at: location,
    paragraph: withContent(paragraph, patched.content),
    inverse,
  });
};

const withParagraphFormatting = (
  paragraph: Paragraph,
  formatting: ParagraphFormatting | undefined,
): Paragraph => {
  const next: Paragraph = { ...paragraph };
  if (formatting === undefined) {
    delete next.formatting;
  } else {
    next.formatting = formatting;
  }
  return next;
};

const setParagraphProps = (document: Document, op: SetParagraphPropsOp): Applied => {
  const located = locate(document, op, op.story, op.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  const formatting = patchedSet(paragraph.formatting, op.patch, op.whenEmpty);
  if (structurallyEqual(formatting ?? {}, paragraph.formatting ?? {})) {
    return unchanged(document);
  }
  return edited({
    document,
    story: op.story,
    at: located.value,
    paragraph: withParagraphFormatting(paragraph, formatting),
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: op.story,
        blockId: op.blockId,
        patch: priorValues(paragraph.formatting, op.patch),
        whenEmpty: emptySetSpelling(paragraph.formatting),
      },
    ],
  });
};

const isReservedParaId = (id: string): boolean => id === "" || RESERVED_PARA_ID_PATTERN.test(id);

/** The fields of a paragraph a split gives the new half: all but its id, content and mark. */
const splitFieldsOf = (paragraph: Paragraph): SplitParagraphFields => {
  const fields: Partial<Paragraph> = { ...paragraph };
  delete fields.type;
  delete fields.paraId;
  delete fields.content;
  delete fields.sectionProperties;
  delete fields.pPrMark;
  return fields;
};

const splitBlock = (document: Document, op: SplitBlockOp): Applied => {
  if (isReservedParaId(op.newBlockId)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      `${op.newBlockId} is not an id.`,
    );
  }
  const wanted = op.newBlockId.toUpperCase();
  if (paragraphIdsIn(document.package).some((id) => id.toUpperCase() === wanted)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      `${op.newBlockId} is already used in the package.`,
    );
  }
  const located = locatePosition(document, op, op.at, "insertion");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location, gap } = located.value;
  const { paragraph } = location;
  const cut = cutAt(paragraph.content, gap);
  if (cut.through.some(isIdentifiedNode)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.SPLITS_IDENTIFIED_CONTAINER,
      `${op.at.offset} in ${op.at.blockId} is inside a tracked change or content control.`,
    );
  }
  const first: Paragraph = { ...paragraph, content: cut.before };
  delete first.sectionProperties;
  delete first.pPrMark;
  if (op.firstMark !== undefined) {
    first.pPrMark = op.firstMark;
  }
  const fields =
    op.newParagraph ??
    (paragraph.formatting === undefined ? {} : { formatting: paragraph.formatting });
  const second: Paragraph = {
    ...fields,
    type: "paragraph",
    paraId: op.newBlockId,
    content: cut.after,
  };
  if (paragraph.sectionProperties !== undefined) {
    second.sectionProperties = paragraph.sectionProperties;
  }
  if (paragraph.pPrMark !== undefined) {
    second.pPrMark = paragraph.pPrMark;
  }
  return replaced({
    document,
    story: op.at.story,
    at: location,
    before: [paragraph],
    after: [first, second],
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: op.at.story,
        blockId: op.at.blockId,
        nextBlockId: op.newBlockId,
        depth: cut.through.length,
      },
    ],
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
  if (isReservedParaId(op.nextBlockId)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      `${op.nextBlockId} could not name the paragraph again.`,
    );
  }
  const depth = op.depth ?? 0;
  if (!Number.isInteger(depth) || depth < 0) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A join depth is a count.");
  }
  const content = joinContent(leading.content, trailing.content, depth);
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      `The records meeting between ${op.blockId} and ${op.nextBlockId} cannot be merged.`,
    );
  }
  const joined: Paragraph = { ...leading, content };
  delete joined.pPrMark;
  if (trailing.sectionProperties !== undefined) {
    joined.sectionProperties = trailing.sectionProperties;
  }
  if (trailing.pPrMark !== undefined) {
    joined.pPrMark = trailing.pPrMark;
  }
  const length = paragraphLength(leading);
  const split: SplitBlockOp = {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: {
      story: op.story,
      blockId: op.blockId,
      offset: length,
      zeroWidthBefore: zeroWidthLeavesAt(leading.content, length).length,
    },
    newBlockId: op.nextBlockId,
    newParagraph: splitFieldsOf(trailing),
  };
  if (leading.pPrMark !== undefined) {
    split.firstMark = leading.pPrMark;
  }
  return replaced({
    document,
    story: op.story,
    at: first.value,
    before: [leading, trailing],
    after: [joined],
    inverse: [split],
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
  if (op.blocks.some(({ paraId }) => paraId === undefined)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      "A replacement paragraph has no id.",
    );
  }
  const before = found.map(({ paragraph }) => paragraph);
  const remaining = countIds(paragraphIdsIn(document.package));
  for (const id of paragraphIdsIn(before)) {
    remaining.set(id, (remaining.get(id) ?? 1) - 1);
  }
  if (idCollision(remaining, paragraphIdsIn(op.blocks))) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "A replacement paragraph id is already used in the package.",
    );
  }
  return replaced({
    document,
    story: op.story,
    at: start,
    before,
    after: op.blocks,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: op.story,
        expected: op.blocks,
        blocks: before,
      },
    ],
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
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
      return insertContent(document, op);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return deleteRange(document, op);
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
      return splitInline(document, op);
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return joinInline(document, op);
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
