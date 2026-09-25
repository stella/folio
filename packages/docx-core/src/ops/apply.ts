/**
 * Apply a document operation and record its exact inverse.
 *
 * `applyDocumentOp` is total: it returns the new document or a typed refusal,
 * and a refused operation changes nothing. It is deterministic: no clock, no
 * randomness and no counter outside its inputs, so equal documents and equal
 * operations give equal results in any realm, whether or not their records
 * share objects. It is local: every block the operation does not name is the
 * same object in the output as in the input, and so is every package part
 * other than the story it edits. It applies only to a document that meets
 * the seed contract (`contract.ts`), and leaves the contract holding.
 *
 * The inverse is a list of ordinary operations, addressed by position like
 * the one it undoes, captured from the state the operation changed:
 *
 * | operation           | inverse                                                   |
 * | ------------------- | --------------------------------------------------------- |
 * | `insertText`        | `deleteRange` (its `join` merges a run the text cut)      |
 * | `insertContent`     | `deleteRange` (its `join` merges the records the cut left) |
 * | `deleteRange`       | `insertContent` with the removed slice                    |
 * | `splitInline`       | `joinInline`                                              |
 * | `joinInline`        | `splitInline`                                             |
 * | `setRunProps`       | `setRunProps` per stretch of prior values, joining cuts   |
 * | `setParagraphProps` | `setParagraphProps` with the prior values                 |
 * | `splitBlock`        | `joinBlocks`                                              |
 * | `joinBlocks`        | `splitBlock` with the second paragraph's fields           |
 * | `replaceBlocks`     | `replaceBlocks`                                           |
 *
 * Each inverse carries its precondition: a deletion the slice it removes, a
 * patch the values it replaces, a join the fields of the paragraph it merges
 * away. An inverse applied to content that has changed since is refused as
 * stale. An inverse that cuts records apart again names the ids its forward
 * operation retired.
 */

import { Result } from "better-result";

import type { Document, Paragraph, ParagraphFormatting, Run } from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import {
  type ParagraphLocation,
  replaceParagraphs,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { meetsContract, validateOpsDocument } from "./contract";
import { equalForStaleness, structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import {
  collides,
  countIds,
  countKeys,
  identityKeysIn,
  idKey,
  isParaId,
  packageIdentityKeys,
  packageParagraphIds,
  paragraphIdsIn,
} from "./ids";
import {
  concatIds,
  cutAt,
  deleteBetween,
  emptySetSpelling,
  gapAfterInserted,
  insertionInverse,
  insertSliceAt,
  insertTextInContent,
  type Joined,
  joinAt,
  joinContent,
  namesIds,
  patchedSet,
  patchRunsBetween,
  runsBetween,
  splitAt,
} from "./inline";
import {
  childNodes,
  defaultInsertionGap,
  type Gap,
  type InlineNode,
  isEmptyRecord,
  recordsBetween,
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
  type NewIds,
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

/** Characters `insertText` does not carry: each is an inline atom of its own. */
const NON_TEXT_CHARACTER_PATTERN = /[\t\n\r]/u;

const isHighSurrogate = (code: number): boolean => code >= 0xd8_00 && code <= 0xdb_ff;
const isLowSurrogate = (code: number): boolean => code >= 0xdc_00 && code <= 0xdf_ff;

/** The paragraph carrying an id, compared as hex; the contract makes it unique. */
const findParagraph = (
  op: DocumentOp,
  paragraphs: readonly ParagraphLocation[],
  blockId: string,
): Result<ParagraphLocation, DocumentOpRefusal> => {
  const key = idKey(blockId);
  const match = paragraphs.find(({ paragraph }) => idKey(paragraph.paraId ?? "") === key);
  return match === undefined
    ? Result.err(
        refusal(op, DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, `No paragraph is ${blockId}.`),
      )
    : Result.ok(match);
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
  if (from.story !== to.story || idKey(from.blockId) !== idKey(to.blockId)) {
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

/** A position naming its paragraph by the id the paragraph carries. */
const positionAt = (story: OpStory, paragraph: Paragraph, gap: Gap): TextPosition => ({
  story,
  blockId: paragraph.paraId ?? "",
  offset: gap.offset,
  zeroWidthBefore: gap.zeroWidthBefore,
});

const touchedBetween = (
  before: readonly Paragraph[],
  after: readonly Paragraph[],
): TouchedBlocks => {
  const beforeIds = before.flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId]));
  const afterIds = after.flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId]));
  const afterKeys = new Set(afterIds.map(idKey));
  const beforeKeys = new Set(beforeIds.map(idKey));
  return {
    modified: beforeIds.filter((id) => afterKeys.has(idKey(id))),
    inserted: afterIds.filter((id) => !beforeKeys.has(idKey(id))),
    removed: beforeIds.filter((id) => !afterKeys.has(idKey(id))),
  };
};

/** Revision and control ids the package uses outside `before`. */
const identitiesOutside = (document: Document, before: readonly Paragraph[]): Set<string> => {
  const counts = countKeys(packageIdentityKeys(document.package));
  for (const key of identityKeysIn(before)) counts.set(key, (counts.get(key) ?? 1) - 1);
  return new Set([...counts].flatMap(([key, count]) => (count > 0 ? [key] : [])));
};

type Commit = {
  document: Document;
  op: DocumentOp;
  story: OpStory;
  at: ParagraphLocation;
  before: readonly Paragraph[];
  after: readonly Paragraph[];
  newIds: NewIds | undefined;
  /** Records of `after` holding only inserted content. */
  inserted?: ReadonlySet<object>;
};

type Committed = { document: Document; paragraphs: Paragraph[]; touched: TouchedBlocks };

/** `after` with the records the edit would leave sharing an id given the operation's new ids. */
const withFreshIds = ({
  document,
  op,
  before,
  after,
  newIds,
  inserted,
}: Omit<Commit, "story" | "at">): Result<Paragraph[], DocumentOpRefusal> => {
  const freshened = freshenIdentities({
    before,
    after,
    newIds: newIds ?? {},
    usedElsewhere: () => identitiesOutside(document, before),
    ...(inserted === undefined ? {} : { inserted }),
  });
  switch (freshened.kind) {
    case "needsIds":
      return Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
          `The edit cuts records carrying ids and needs ${freshened.missing} more new ids.`,
        ),
      );
    case "invalidId":
      return Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
          `${freshened.id} cannot be a new revision or content-control id here.`,
        ),
      );
    case "fresh":
      return Result.ok(freshened.paragraphs);
    default: {
      const unreachable: never = freshened;
      return unreachable;
    }
  }
};

/** Put `after` where `before` stands, its records given fresh ids where they need them. */
const commit = (options: Commit): Result<Committed, DocumentOpRefusal> => {
  const paragraphs = withFreshIds(options);
  if (paragraphs.isErr()) {
    return Result.err(paragraphs.error);
  }
  const { document, story, at, before } = options;
  return Result.ok({
    document: replaceParagraphs({
      document,
      story,
      at,
      count: before.length,
      replacement: paragraphs.value,
    }),
    paragraphs: paragraphs.value,
    touched: touchedBetween(before, paragraphs.value),
  });
};

const withContent = (
  paragraph: Paragraph,
  content: readonly Paragraph["content"][number][],
): Paragraph =>
  content === paragraph.content ? paragraph : { ...paragraph, content: [...content] };

/** An edit of one paragraph, committed, with the inverse built from the paragraph it became. */
const editOne = (
  commitOptions: Omit<Commit, "before" | "after"> & { paragraph: Paragraph },
  inverseOf: (result: Paragraph) => Result<readonly DocumentOp[], DocumentOpRefusal>,
): Applied => {
  const { paragraph, at, document } = commitOptions;
  if (paragraph === at.paragraph) {
    return unchanged(document);
  }
  const committed = commit({ ...commitOptions, before: [at.paragraph], after: [paragraph] });
  if (committed.isErr()) {
    return Result.err(committed.error);
  }
  const [result] = committed.value.paragraphs;
  if (result === undefined) {
    return Result.ok({ document: committed.value.document, inverse: [], touched: UNTOUCHED });
  }
  const inverse = inverseOf(result);
  if (inverse.isErr()) {
    return Result.err(inverse.error);
  }
  return Result.ok({
    document: committed.value.document,
    inverse: inverse.value,
    touched: committed.value.touched,
  });
};

type InsertedOptions = {
  document: Document;
  op: DocumentOp;
  story: OpStory;
  location: ParagraphLocation;
  content: Paragraph["content"];
  start: Gap;
  end: Gap;
  newIds: NewIds | undefined;
};

/** Commit an insertion, recording the one deletion that undoes it. */
const inserted = ({
  document,
  op,
  story,
  location,
  content,
  start,
  end,
  newIds,
}: InsertedOptions): Applied =>
  editOne(
    {
      document,
      op,
      story,
      at: location,
      paragraph: withContent(location.paragraph, content),
      newIds,
      inserted: recordsBetween(content, start, end),
    },
    (result) => {
      const inverse = insertionInverse(location.paragraph.content, result.content, start, end);
      if (inverse === undefined) {
        return Result.err(
          refusal(
            op,
            DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
            "The insertion would merge records that were separate.",
          ),
        );
      }
      const deletion: DeleteRangeOp = {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: positionAt(story, result, start),
        to: positionAt(story, result, end),
        expected: inverse.removed,
      };
      if (inverse.join > 0) {
        deletion.join = inverse.join;
      }
      return Result.ok([deletion]);
    },
  );

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
  return inserted({
    document,
    op,
    story: op.at.story,
    location,
    content,
    start,
    end,
    newIds: op.newIds,
  });
};

const holdsEmptyRecord = (nodes: readonly InlineNode[]): boolean =>
  nodes.some((node) => isEmptyRecord(node) || holdsEmptyRecord(childNodes(node) ?? []));

const insertContent = (document: Document, op: InsertContentOp): Applied => {
  // The slice's records are new records wherever they came from: copies, so
  // one that repeats a record still in the document is told apart from it.
  const slice = { ...op.slice, content: structuredClone(op.slice.content) };
  if (slice.content.length === 0 || holdsEmptyRecord(slice.content)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
      "The slice holds nothing, or an empty run or text node.",
    );
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
  if (incoming.length > 0 && collides(countIds(packageParagraphIds(document.package)), incoming)) {
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
    newIds: op.newIds,
  });
};

/** A merge's content, or the refusal naming why it cannot be made. */
const joinedContent = (
  op: DocumentOp,
  joined: Joined,
  where: string,
): Result<{ content: Paragraph["content"]; retired: NewIds }, DocumentOpRefusal> => {
  switch (joined.kind) {
    case "joined":
      return Result.ok({ content: joined.content, retired: joined.retired });
    case "notAlike":
      return Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          `The records meeting ${where} cannot be merged.`,
        ),
      );
    case "sharedId":
      return Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.SHARED_ID,
          `The records meeting ${where} carry the same id.`,
        ),
      );
    default: {
      const unreachable: never = joined;
      return unreachable;
    }
  }
};

const isCount = (value: number): boolean => Number.isInteger(value) && value >= 0;

const deleteRange = (document: Document, op: DeleteRangeOp): Applied => {
  const join = op.join ?? 0;
  if (!isCount(join)) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A join depth is a count.");
  }
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
  const { content: remaining, removed } = deleteBetween(paragraph.content, from, to);
  if (expected !== undefined && !equalForStaleness(expected, removed)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `The range of ${op.from.blockId} holds other content than expected.`,
    );
  }
  if (removed.content.length === 0) {
    return unchanged(document);
  }
  let content = remaining;
  let retired: NewIds = {};
  if (join > 0) {
    const joined = joinedContent(op, joinAt(remaining, from, join), `at ${from.offset}`);
    if (joined.isErr()) {
      return Result.err(joined.error);
    }
    ({ content, retired } = joined.value);
  }
  return editOne(
    {
      document,
      op,
      story,
      at: location,
      paragraph: withContent(paragraph, content),
      newIds: undefined,
    },
    (result) => {
      const insertion: InsertContentOp = {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at: positionAt(story, result, from),
        slice: removed,
      };
      if (namesIds(retired)) {
        insertion.newIds = retired;
      }
      return Result.ok([insertion]);
    },
  );
};

const splitInline = (document: Document, op: SplitInlineOp): Applied => {
  if (!isCount(op.depth) || op.depth === 0) {
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
  const content = splitAt(location.paragraph.content, gap, op.depth);
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      `Fewer than ${op.depth} records run across ${op.at.offset} in ${op.at.blockId}.`,
    );
  }
  return editOne(
    {
      document,
      op,
      story: op.at.story,
      at: location,
      paragraph: withContent(location.paragraph, content),
      newIds: op.newIds,
    },
    (result) =>
      Result.ok([
        {
          type: DOCUMENT_OP_TYPES.JOIN_INLINE,
          at: positionAt(op.at.story, result, gap),
          depth: op.depth,
        },
      ]),
  );
};

const joinInline = (document: Document, op: JoinInlineOp): Applied => {
  if (!isCount(op.depth) || op.depth === 0) {
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
  const joined = joinedContent(
    op,
    joinAt(location.paragraph.content, gap, op.depth),
    `at ${op.at.offset} in ${op.at.blockId}`,
  );
  if (joined.isErr()) {
    return Result.err(joined.error);
  }
  const { content, retired } = joined.value;
  return editOne(
    {
      document,
      op,
      story: op.at.story,
      at: location,
      paragraph: withContent(location.paragraph, content),
      newIds: undefined,
    },
    (result) => {
      const split: SplitInlineOp = {
        type: DOCUMENT_OP_TYPES.SPLIT_INLINE,
        at: positionAt(op.at.story, result, gap),
        depth: op.depth,
      };
      if (namesIds(retired)) {
        split.newIds = retired;
      }
      return Result.ok([split]);
    },
  );
};

/** Whether a property set states each key of `expected` as it says (`null`: absent). */
const statesValues = (formatting: object | undefined, expected: object): boolean => {
  const values = new Map(Object.entries(formatting ?? {}));
  return Object.entries(expected).every(
    ([key, value]) => value === undefined || structurallyEqual(values.get(key) ?? null, value),
  );
};

const setRunProps = (document: Document, op: SetRunPropsOp): Applied => {
  const joinStart = op.joinStart ?? 0;
  const joinEnd = op.joinEnd ?? 0;
  if (!isCount(joinStart) || !isCount(joinEnd)) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A join depth is a count.");
  }
  const located = locateRange(document, op, op.from, op.to);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { story, location, from, to, empty } = located.value;
  if (empty) {
    return unchanged(document);
  }
  const { paragraph } = location;
  const { expected } = op;
  if (
    expected !== undefined &&
    !runsBetween(paragraph.content, from, to).every((run: Run) =>
      statesValues(run.formatting, expected),
    )
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `The runs of ${op.from.blockId} state other values than expected.`,
    );
  }
  const patched = patchRunsBetween(paragraph.content, from, to, op.patch, op.whenEmpty);
  const patchedContent = patched?.content ?? paragraph.content;

  // Records the patch cut take their new ids before any join sees them.
  const freshened = withFreshIds({
    document,
    op,
    before: [paragraph],
    after: [withContent(paragraph, patchedContent)],
    newIds: op.newIds,
  });
  if (freshened.isErr()) {
    return Result.err(freshened.error);
  }
  let content = freshened.value[0]?.content ?? patchedContent;
  const retiredAt: { from: NewIds; to: NewIds } = { from: {}, to: {} };
  for (const [gap, depth, side] of [
    [from, joinStart, "from"],
    [to, joinEnd, "to"],
  ] as const) {
    if (depth === 0) continue;
    const joined = joinedContent(op, joinAt(content, gap, depth), `at ${gap.offset}`);
    if (joined.isErr()) {
      return Result.err(joined.error);
    }
    content = joined.value.content;
    retiredAt[side] = joined.value.retired;
  }
  const result = withContent(paragraph, content);
  if (structurallyEqual(result.content, paragraph.content)) {
    return unchanged(document);
  }

  const restoring = patched?.restoring ?? [];
  const inverse: SetRunPropsOp[] = restoring.map((span) => ({
    type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
    from: positionAt(story, result, span.from),
    to: positionAt(story, result, span.to),
    patch: span.patch,
    whenEmpty: span.whenEmpty,
    expected: op.patch,
  }));
  // Runs this patch left cut at either end merge back once their values are restored.
  const cutAtGap = (gap: Gap): number =>
    spanningRecords(paragraph.content, [gap], 0, 1).length -
    spanningRecords(content, [gap], 0, 1).length;
  const first = inverse.at(0);
  const last = inverse.at(-1);
  if (first !== undefined && last !== undefined) {
    const cutFrom = cutAtGap(from);
    const cutTo = cutAtGap(to);
    if (cutFrom > 0) first.joinStart = cutFrom;
    if (cutTo > 0) last.joinEnd = cutTo;
    // The cuts that undo this patch's joins give the retired ids back.
    if (namesIds(retiredAt.from)) first.newIds = concatIds(first.newIds ?? {}, retiredAt.from);
    if (namesIds(retiredAt.to)) last.newIds = concatIds(last.newIds ?? {}, retiredAt.to);
  }
  return Result.ok({
    document: replaceParagraphs({ document, story, at: location, count: 1, replacement: [result] }),
    inverse,
    touched: touchedBetween([paragraph], [result]),
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
  if (op.expected !== undefined && !statesValues(paragraph.formatting, op.expected)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${op.blockId} states other paragraph properties than expected.`,
    );
  }
  const formatting = patchedSet(paragraph.formatting, op.patch, op.whenEmpty);
  if (structurallyEqual(formatting ?? {}, paragraph.formatting ?? {})) {
    return unchanged(document);
  }
  return editOne(
    {
      document,
      op,
      story: op.story,
      at: located.value,
      paragraph: withParagraphFormatting(paragraph, formatting),
      newIds: undefined,
    },
    (result) =>
      Result.ok([
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: op.story,
          blockId: result.paraId ?? op.blockId,
          patch: priorValues(paragraph.formatting, op.patch),
          whenEmpty: emptySetSpelling(paragraph.formatting),
          expected: op.patch,
        },
      ]),
  );
};

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
  if (!isParaId(op.newBlockId)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      `${op.newBlockId} is not a paragraph id.`,
    );
  }
  if (collides(countIds(packageParagraphIds(document.package)), [op.newBlockId])) {
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
  const committed = commit({
    document,
    op,
    story: op.at.story,
    at: location,
    before: [paragraph],
    after: [first, second],
    newIds: op.newIds,
  });
  if (committed.isErr()) {
    return Result.err(committed.error);
  }
  const [, placedSecond] = committed.value.paragraphs;
  return Result.ok({
    document: committed.value.document,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: op.at.story,
        blockId: paragraph.paraId ?? op.at.blockId,
        nextBlockId: op.newBlockId,
        depth: cut.through.length,
        expectedSecond: splitFieldsOf(placedSecond ?? second),
      },
    ],
    touched: committed.value.touched,
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
  if (
    op.expectedSecond !== undefined &&
    !equalForStaleness(
      { type: "paragraph", ...splitFieldsOf(trailing) },
      { type: "paragraph", ...op.expectedSecond },
    )
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${op.nextBlockId} states other fields than expected.`,
    );
  }
  const trailingId = trailing.paraId ?? "";
  // The split that undoes the join creates the second paragraph again, under its id.
  if (!isParaId(trailingId)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      `${trailingId} could not name the paragraph again.`,
    );
  }
  const depth = op.depth ?? 0;
  if (!isCount(depth)) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A join depth is a count.");
  }
  const merged = joinedContent(
    op,
    joinContent(leading.content, trailing.content, depth),
    `between ${op.blockId} and ${op.nextBlockId}`,
  );
  if (merged.isErr()) {
    return Result.err(merged.error);
  }
  const joined: Paragraph = { ...leading, content: merged.value.content };
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
      blockId: leading.paraId ?? op.blockId,
      offset: length,
      zeroWidthBefore: zeroWidthLeavesAt(leading.content, length).length,
    },
    newBlockId: trailingId,
    newParagraph: splitFieldsOf(trailing),
  };
  if (leading.pPrMark !== undefined) {
    split.firstMark = leading.pPrMark;
  }
  if (namesIds(merged.value.retired)) {
    split.newIds = merged.value.retired;
  }
  const committed = commit({
    document,
    op,
    story: op.story,
    at: first.value,
    before: [leading, trailing],
    after: [joined],
    newIds: undefined,
  });
  if (committed.isErr()) {
    return Result.err(committed.error);
  }
  return Result.ok({
    document: committed.value.document,
    inverse: [split],
    touched: committed.value.touched,
  });
};

/**
 * Whether a replacement leaves every section break where it was: the replaced
 * paragraphs lie in one section (only the last may end it) and the
 * replacement ends it the same way. Anything else would change the sections
 * the body is divided into, which is a section operation.
 */
const sameSectionBreaks = (before: readonly Paragraph[], after: readonly Paragraph[]): boolean => {
  const breaksInside = (paragraphs: readonly Paragraph[]) =>
    paragraphs.slice(0, -1).some(({ sectionProperties }) => sectionProperties !== undefined);
  return (
    !breaksInside(before) &&
    !breaksInside(after) &&
    structurallyEqual(before.at(-1)?.sectionProperties, after.at(-1)?.sectionProperties)
  );
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
    (location, offset) => !equalForStaleness(location.paragraph, op.expected[offset]),
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
  if (op.blocks.some(({ content }) => holdsEmptyRecord(content))) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
      "A replacement paragraph holds an empty run or text node.",
    );
  }
  const before = found.map(({ paragraph }) => paragraph);
  // An id only one side has is created by the replacement or by its inverse.
  const kept = new Set(before.map(({ paraId }) => idKey(paraId ?? "")));
  const placed = new Set(op.blocks.map(({ paraId }) => idKey(paraId ?? "")));
  const created = [
    ...op.blocks.filter(({ paraId }) => !kept.has(idKey(paraId ?? ""))),
    ...before.filter(({ paraId }) => !placed.has(idKey(paraId ?? ""))),
  ];
  if (created.some(({ paraId }) => paraId === undefined || !isParaId(paraId))) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      "A paragraph the replacement adds or removes has no usable id.",
    );
  }
  if (!sameSectionBreaks(before, op.blocks)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      "A replacement keeps every section break where it is.",
    );
  }
  const remaining = countIds(packageParagraphIds(document.package));
  for (const id of paragraphIdsIn(before)) {
    remaining.set(idKey(id), (remaining.get(idKey(id)) ?? 1) - 1);
  }
  if (collides(remaining, paragraphIdsIn(op.blocks))) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "A replacement paragraph id is already used in the package.",
    );
  }
  const outside = identitiesOutside(document, before);
  const incomingRecords = identityKeysIn(op.blocks);
  if (
    incomingRecords.some((key) => outside.has(key)) ||
    new Set(incomingRecords).size !== incomingRecords.length
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "A replacement carries a revision or content-control id already used in the package.",
    );
  }
  return Result.ok({
    document: replaceParagraphs({
      document,
      story: op.story,
      at: start,
      count: before.length,
      replacement: op.blocks,
    }),
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: op.story,
        expected: op.blocks,
        blocks: before,
      },
    ],
    touched: touchedBetween(before, op.blocks),
  });
};

const dispatch = (document: Document, op: DocumentOp): Applied => {
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

/** Apply one operation to a document that meets the seed contract. */
export const applyDocumentOp = (
  document: Document,
  op: DocumentOp,
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  const valid = validateOpsDocument(document);
  if (valid.isErr()) {
    return refuse(op, valid.error.reason, valid.error.message);
  }
  const applied = dispatch(document, op);
  if (applied.isOk()) {
    meetsContract(applied.value.document);
  }
  return applied;
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
