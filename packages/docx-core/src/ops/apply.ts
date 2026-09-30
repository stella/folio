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
 * | `splitBlock`        | `joinBlocks`, then the kept half's review fields         |
 * | `joinBlocks`        | `splitBlock` with the retired paragraph's fields, then  |
 * |                     | the survivor's review fields                              |
 * | `insertBlocks`      | an anchored `replaceBlocks`                              |
 * | `replaceBlocks`     | `replaceBlocks`                                           |
 * | `setParagraphReview`| `setParagraphReview`                                      |
 * | `replaceInline`     | `replaceInline`                                           |
 * | `resolveRevision`   | the inverses of the operations it expands to              |
 *
 * A tracked operation (one carrying a `revision` stamp, see `review.ts`) is
 * undone as follows: a tracked insertion by `deleteRange`, a tracked split by
 * `joinBlocks`, a tracked deletion or run patch by `replaceInline`, and a
 * tracked paragraph patch or join by `setParagraphReview`.
 *
 * Each inverse carries its precondition: a deletion the slice it removes, a
 * patch the values it replaces, a join the fields of the paragraph it merges
 * away. An inverse applied to content that has changed since is refused as
 * stale. An inverse that cuts records apart again names the ids its forward
 * operation retired.
 */

import { Result } from "better-result";

import {
  type Document,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphFormatting,
  type Run,
  type RunPropertyChange,
} from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import { deleteBlocks } from "./blockDeletion";
import { insertBlocks } from "./blockInsertion";
import {
  endsItsContainer,
  type ParagraphLocation,
  replaceParagraphs,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { meetsContract, validateOpsDocument } from "./contract";
import { combineEdits, type DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import { freshenIdentities } from "./identity";
import {
  collides,
  countIds,
  countKeys,
  IDENTITY_SPACES,
  identityKeysIn,
  idKey,
  isParaId,
  packageIdentityKeys,
  packageParagraphIds,
  paragraphIdsIn,
  slotKey,
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
import { resolveRevision } from "./resolve";
import { applyRowOp } from "./tableRows";
import { applyTableOp } from "./tables";
import { stampedTableRowRevisionIds } from "./tableTracking";
import {
  namesMarkFormatting,
  paragraphPropertiesOf,
  paragraphPropertyChange,
  reviewFieldsOf,
  sameParagraphProperties,
  stampedRevisionIds,
  stampInfo,
  withMarkFormatting,
  withReviewFields,
  WRAP_KINDS,
  type WrapKind,
  wrapTracked,
} from "./review";
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
  type ParagraphReviewFields,
  type ReplaceBlocksOp,
  type ReplaceInlineOp,
  type RevisionStamp,
  type SetParagraphPropsOp,
  type SetParagraphReviewOp,
  type SetRunPropsOp,
  type SplitBlockOp,
  SPLIT_HALVES,
  type SplitHalf,
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
  /**
   * The revision ids of the tracked changes the operation recorded, in
   * document order: its stamp's id and the new ids its other records took.
   * Empty for a direct operation, and for a tracked one that merged into a
   * change carrying the same stamp.
   */
  revisions: readonly number[];
};

/** What one operation did, before the revisions it recorded are read off it. */
type Applied = Result<DocumentEdit, DocumentOpRefusal>;

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

/**
 * Whether a tracked operation's stamp can name the first record it creates:
 * a revision id in range that no record in the package carries.
 */
const stampRefusal = (
  document: Document,
  op: DocumentOp,
  stamp: RevisionStamp,
): DocumentOpRefusal | undefined => {
  // A revision id reserves no value: every integer in range names a record.
  const revisionId = stamp.id;
  if (!Number.isInteger(revisionId) || revisionId < 0 || revisionId > MAX_REVISION_ID) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      `${revisionId} cannot be a revision id.`,
    );
  }
  const key = slotKey({ space: IDENTITY_SPACES.REVISION, id: stamp.id });
  if (packageIdentityKeys(document.package).includes(key)) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      `Revision id ${stamp.id} is already used in the package.`,
    );
  }
  return undefined;
};

/** The replacement that gives a paragraph back the content it had. */
const contentRestoring = (
  story: OpStory,
  result: Paragraph,
  original: Paragraph,
): ReplaceInlineOp => ({
  type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
  story,
  blockId: result.paraId ?? original.paraId ?? "",
  expected: result.content,
  content: original.content,
});

/** The review fields' setting that gives a paragraph back the ones it had. */
const reviewRestoring = (
  story: OpStory,
  result: Paragraph,
  original: Paragraph,
): SetParagraphReviewOp => ({
  type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
  story,
  blockId: result.paraId ?? original.paraId ?? "",
  expected: reviewFieldsOf(result),
  review: reviewFieldsOf(original),
});

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
  revision: RevisionStamp | undefined;
};

type Wrapped = Result<Paragraph["content"] | undefined, DocumentOpRefusal>;

/**
 * Content with the leaves between two gaps recorded as a tracked change;
 * `undefined` when there is nothing to record.
 */
const wrappedContent = (
  op: DocumentOp,
  options: {
    items: Paragraph["content"];
    from: Gap;
    to: Gap;
    kind: WrapKind;
    stamp: RevisionStamp;
  },
): Wrapped => {
  const wrapped = wrapTracked(options);
  switch (wrapped.kind) {
    case "wrapped":
      return Result.ok(wrapped.content);
    case "unchanged":
      return Result.ok(undefined);
    case "refused":
      return Result.err(
        refusal(
          op,
          wrapped.reason,
          wrapped.reason === DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE
            ? "A tracked change cannot hold a comment boundary or reference."
            : "The tracked edit falls inside tracked-removed content.",
        ),
      );
    default: {
      const unreachable: never = wrapped;
      return unreachable;
    }
  }
};

/**
 * Commit an insertion, recording the one deletion that undoes it. A tracked
 * insertion is wrapped first; the deletion removes the wrapper with it.
 */
const inserted = (options: InsertedOptions): Applied => {
  const { document, op, story, location, start, end, newIds, revision } = options;
  let { content } = options;
  if (revision !== undefined) {
    // The wrap cuts the records again; an insertion the direct one refuses stays refused.
    if (insertionInverse(location.paragraph.content, content, start, end) === undefined) {
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "The insertion would merge records that were separate.",
      );
    }
    const wrapped = wrappedContent(op, {
      items: content,
      from: start,
      to: end,
      kind: WRAP_KINDS.INSERTION,
      stamp: revision,
    });
    if (wrapped.isErr()) {
      return Result.err(wrapped.error);
    }
    content = wrapped.value ?? content;
  }
  return editOne(
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
  return inserted({
    document,
    op,
    story: op.at.story,
    location,
    content,
    start,
    end,
    newIds: op.newIds,
    revision: op.revision,
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
      "The slice holds nothing, or an empty run, text node, or revision wrapper.",
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
    revision: op.revision,
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
  if (op.revision !== undefined && join !== 0) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A tracked deletion joins nothing.",
    );
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
  if (op.revision !== undefined) {
    const wrapped = wrappedContent(op, {
      items: paragraph.content,
      from,
      to,
      kind: WRAP_KINDS.DELETION,
      stamp: op.revision,
    });
    if (wrapped.isErr()) {
      return Result.err(wrapped.error);
    }
    if (wrapped.value === undefined) {
      return unchanged(document);
    }
    return editOne(
      {
        document,
        op,
        story,
        at: location,
        paragraph: withContent(paragraph, wrapped.value),
        newIds: op.newIds,
      },
      (result) => Result.ok([contentRestoring(story, result, paragraph)]),
    );
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
  const { revision } = op;
  if (revision !== undefined && (joinStart !== 0 || joinEnd !== 0)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A tracked patch joins nothing.",
    );
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
  if (revision !== undefined) {
    // A run already carrying a property change keeps it, and the formatting it started from.
    const recordChange = (patchedRun: Run, previous: Run): Run => {
      if ((previous.propertyChanges?.length ?? 0) > 0) {
        return patchedRun;
      }
      const change: RunPropertyChange = { type: "runPropertyChange", info: stampInfo(revision) };
      if (previous.formatting !== undefined) {
        change.previousFormatting = previous.formatting;
      }
      return { ...patchedRun, propertyChanges: [change] };
    };
    const tracked = patchRunsBetween(
      paragraph.content,
      from,
      to,
      op.patch,
      op.whenEmpty,
      recordChange,
    );
    if (tracked === undefined) {
      return unchanged(document);
    }
    return editOne(
      {
        document,
        op,
        story,
        at: location,
        paragraph: withContent(paragraph, tracked.content),
        newIds: op.newIds,
      },
      (result) => Result.ok([contentRestoring(story, result, paragraph)]),
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
  const { revision } = op;
  if (revision !== undefined && namesMarkFormatting(op.patch)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A paragraph property change does not record the paragraph mark's run properties.",
    );
  }
  const formatting = patchedSet(paragraph.formatting, op.patch, op.whenEmpty);
  if (structurallyEqual(formatting ?? {}, paragraph.formatting ?? {})) {
    return unchanged(document);
  }
  if (revision !== undefined) {
    const next = withParagraphFormatting(paragraph, formatting);
    // A paragraph already carrying a property change keeps it, and the formatting it started from.
    if ((paragraph.propertyChanges?.length ?? 0) === 0) {
      next.propertyChanges = [paragraphPropertyChange(stampInfo(revision), paragraph.formatting)];
    }
    return editOne(
      {
        document,
        op,
        story: op.story,
        at: located.value,
        paragraph: next,
        newIds: undefined,
      },
      (result) => Result.ok([reviewRestoring(op.story, result, paragraph)]),
    );
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

/**
 * The paragraph properties two paragraphs joined have: the first's paragraph
 * properties with the second's mark run properties, since the mark that stays
 * is the second's. A first paragraph that holds no content brings nothing:
 * the second keeps its own.
 */
const joinedFormatting = (first: Paragraph, second: Paragraph) =>
  paragraphLength(first) === 0
    ? second.formatting
    : withMarkFormatting(paragraphPropertiesOf(first.formatting), second.formatting);

/** The review fields of two paragraphs once joined: the first's properties, the second's mark. */
const joinedReview = (first: Paragraph, second: Paragraph): ParagraphReviewFields => {
  const review: ParagraphReviewFields = {};
  const formatting = joinedFormatting(first, second);
  if (formatting !== undefined) review.formatting = formatting;
  if (second.propertyChanges !== undefined) review.propertyChanges = second.propertyChanges;
  if (second.pPrMark !== undefined) review.pPrMark = second.pPrMark;
  return review;
};

/** A review-field setting from `from` to `to`; none when they are the same. */
const reviewSetting = (
  story: OpStory,
  blockId: string,
  from: ParagraphReviewFields,
  to: ParagraphReviewFields,
): SetParagraphReviewOp[] =>
  structurallyEqual(from, to)
    ? []
    : [
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          story,
          blockId,
          expected: from,
          review: to,
        },
      ];

const isSplitHalf = (value: unknown): value is SplitHalf =>
  value === SPLIT_HALVES.FIRST || value === SPLIT_HALVES.SECOND;

type TrackSplitOptions = {
  op: SplitBlockOp;
  stamp: RevisionStamp;
  paragraph: Paragraph;
  /** The halves, recorded as a tracked split in place. */
  first: Paragraph;
  /** The half carrying the new id. */
  made: Paragraph;
};

/**
 * Record a split as tracked: the new mark is an insertion, and the new
 * paragraph records a property change from the source's properties when it
 * states other ones and carries no change already. Returns the refusal when
 * the operation states a mark or property change of its own.
 */
const trackSplit = ({
  op,
  stamp,
  paragraph,
  first,
  made,
}: TrackSplitOptions): DocumentOpRefusal | undefined => {
  if (op.firstMark !== undefined || (op.newParagraph?.propertyChanges?.length ?? 0) > 0) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      "A tracked split records its own mark and property change.",
    );
  }
  first.pPrMark = { kind: "ins", info: stampInfo(stamp) };
  if (
    !sameParagraphProperties(made.formatting, paragraph.formatting) &&
    (made.propertyChanges?.length ?? 0) === 0
  ) {
    made.propertyChanges = [paragraphPropertyChange(stampInfo(stamp), paragraph.formatting)];
  }
  return undefined;
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
  if (op.newHalf !== undefined && !isSplitHalf(op.newHalf)) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A half is first or second.");
  }
  const located = locatePosition(document, op, op.at, "insertion");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location, gap } = located.value;
  const { paragraph } = location;
  const newHalf =
    op.newHalf ??
    (gap.offset === paragraphLength(paragraph) ? SPLIT_HALVES.SECOND : SPLIT_HALVES.FIRST);
  if (newHalf === SPLIT_HALVES.SECOND && (op.newParagraph?.propertyChanges?.length ?? 0) > 0) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "The second half takes the paragraph's property changes with its mark.",
    );
  }
  const cut = cutAt(paragraph.content, gap);
  const fields =
    op.newParagraph ??
    (paragraph.formatting === undefined ? {} : { formatting: paragraph.formatting });
  const made: Paragraph = { ...fields, type: "paragraph", paraId: op.newBlockId, content: [] };
  // The paragraph's own fields stay with the half keeping its id; the mark, the
  // section break and the pending property changes end the second half.
  let first: Paragraph;
  let second: Paragraph;
  if (newHalf === SPLIT_HALVES.FIRST) {
    first = { ...made, content: cut.before };
    second = { ...paragraph, content: cut.after };
  } else {
    first = { ...paragraph, content: cut.before };
    delete first.sectionProperties;
    delete first.pPrMark;
    delete first.propertyChanges;
    second = { ...made, content: cut.after };
    if (paragraph.propertyChanges !== undefined) second.propertyChanges = paragraph.propertyChanges;
    if (paragraph.sectionProperties !== undefined) {
      second.sectionProperties = paragraph.sectionProperties;
    }
    if (paragraph.pPrMark !== undefined) second.pPrMark = paragraph.pPrMark;
  }
  if (op.firstMark !== undefined) {
    first.pPrMark = op.firstMark;
  }
  if (op.revision !== undefined) {
    const madeHalf = newHalf === SPLIT_HALVES.FIRST ? first : second;
    const tracked = trackSplit({ op, stamp: op.revision, paragraph, first, made: madeHalf });
    if (tracked !== undefined) {
      return Result.err(tracked);
    }
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
  const [placedFirst = first, placedSecond = second] = committed.value.paragraphs;
  const kept = newHalf === SPLIT_HALVES.FIRST ? placedSecond : placedFirst;
  const madePlaced = newHalf === SPLIT_HALVES.FIRST ? placedFirst : placedSecond;
  const keptId = paragraph.paraId ?? op.at.blockId;
  return Result.ok({
    document: committed.value.document,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: op.at.story,
        blockId: placedFirst.paraId ?? "",
        nextBlockId: placedSecond.paraId ?? "",
        depth: cut.through.length,
        survivor: newHalf === SPLIT_HALVES.FIRST ? SPLIT_HALVES.SECOND : SPLIT_HALVES.FIRST,
        expectedRetired: splitFieldsOf(madePlaced),
        expectedSurvivor: reviewFieldsOf(kept),
      },
      ...reviewSetting(
        op.at.story,
        kept.paraId ?? keptId,
        joinedReview(placedFirst, placedSecond),
        reviewFieldsOf(paragraph),
      ),
    ],
    touched: committed.value.touched,
  });
};

type TrackJoinOptions = {
  document: Document;
  op: JoinBlocksOp;
  stamp: RevisionStamp;
  at: ParagraphLocation;
  leading: Paragraph;
  trailing: Paragraph;
};

/**
 * Record a join as tracked: the first paragraph's mark becomes a deletion,
 * and the second takes the first's paragraph properties as a property
 * change, so accepting leaves what the direct join leaves: the second
 * paragraph, with the first's properties.
 */
const trackJoin = ({ document, op, stamp, at, leading, trailing }: TrackJoinOptions): Applied => {
  const depth = op.depth ?? 0;
  if (!isCount(depth)) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A join depth is a count.");
  }
  if ((op.survivor ?? SPLIT_HALVES.SECOND) !== SPLIT_HALVES.SECOND) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Resolving a tracked join leaves the second paragraph.",
    );
  }
  if (leading.pPrMark !== undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      `The mark of ${op.blockId} already carries a tracked change.`,
    );
  }
  const first: Paragraph = { ...leading, pPrMark: { kind: "del", info: stampInfo(stamp) } };
  let second = trailing;
  const formatting = joinedFormatting(leading, trailing);
  if (!sameParagraphProperties(formatting, trailing.formatting)) {
    second = withParagraphFormatting(trailing, formatting);
    // A paragraph already carrying a property change keeps it, and the formatting it started from.
    if ((trailing.propertyChanges?.length ?? 0) === 0) {
      second.propertyChanges = [paragraphPropertyChange(stampInfo(stamp), trailing.formatting)];
    }
  }
  const committed = commit({
    document,
    op,
    story: op.story,
    at,
    before: [leading, trailing],
    after: [first, second],
    newIds: op.newIds,
  });
  if (committed.isErr()) {
    return Result.err(committed.error);
  }
  const [placedFirst = first, placedSecond = second] = committed.value.paragraphs;
  const inverse: SetParagraphReviewOp[] = [reviewRestoring(op.story, placedFirst, leading)];
  if (second !== trailing) {
    inverse.push(reviewRestoring(op.story, placedSecond, trailing));
  }
  return Result.ok({
    document: committed.value.document,
    inverse,
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
  const survivorSide = op.survivor ?? SPLIT_HALVES.SECOND;
  if (!isSplitHalf(survivorSide)) {
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A half is first or second.");
  }
  const survivor = survivorSide === SPLIT_HALVES.SECOND ? trailing : leading;
  const retired = survivorSide === SPLIT_HALVES.SECOND ? leading : trailing;
  if (
    op.expectedRetired !== undefined &&
    !equalForStaleness(
      { type: "paragraph", ...splitFieldsOf(retired) },
      { type: "paragraph", ...op.expectedRetired },
    )
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${retired.paraId ?? ""} states other fields than expected.`,
    );
  }
  if (
    op.expectedSurvivor !== undefined &&
    !equalForStaleness(reviewFieldsOf(survivor), op.expectedSurvivor)
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${survivor.paraId ?? ""} states other review fields than expected.`,
    );
  }
  if (op.revision !== undefined) {
    return trackJoin({
      document,
      op,
      stamp: op.revision,
      at: first.value,
      leading,
      trailing,
    });
  }
  const retiredId = retired.paraId ?? "";
  // The split that undoes the join creates the retired paragraph again, under its id.
  if (!isParaId(retiredId)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      `${retiredId} could not name the paragraph again.`,
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
  // The survivor's identity; the first's properties; the second's mark, section
  // break and pending property changes.
  const joined = withReviewFields(
    { ...survivor, content: merged.value.content },
    joinedReview(leading, trailing),
  );
  delete joined.sectionProperties;
  if (trailing.sectionProperties !== undefined) {
    joined.sectionProperties = trailing.sectionProperties;
  }
  const length = paragraphLength(leading);
  const newHalf = survivorSide === SPLIT_HALVES.SECOND ? SPLIT_HALVES.FIRST : SPLIT_HALVES.SECOND;
  const retiredFields = splitFieldsOf(retired);
  if (newHalf === SPLIT_HALVES.SECOND) {
    // The second half takes the property changes back with its mark.
    delete retiredFields.propertyChanges;
  }
  const split: SplitBlockOp = {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: {
      story: op.story,
      blockId: survivor.paraId ?? "",
      offset: length,
      zeroWidthBefore: zeroWidthLeavesAt(leading.content, length).length,
    },
    newBlockId: retiredId,
    newHalf,
    newParagraph: retiredFields,
  };
  if (leading.pPrMark !== undefined) {
    split.firstMark = leading.pPrMark;
  }
  if (namesIds(merged.value.retired)) {
    split.newIds = merged.value.retired;
  }
  // What the split leaves on the survivor's half, before its own fields are given back.
  const splitReview: ParagraphReviewFields = reviewFieldsOf(joined);
  if (survivorSide === SPLIT_HALVES.FIRST) {
    delete splitReview.propertyChanges;
    delete splitReview.pPrMark;
    if (leading.pPrMark !== undefined) splitReview.pPrMark = leading.pPrMark;
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
    inverse: [
      split,
      ...reviewSetting(op.story, survivor.paraId ?? "", splitReview, reviewFieldsOf(survivor)),
    ],
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
      "A replacement paragraph holds an empty run, text node, or revision wrapper.",
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

/** Revision and control ids a paragraph's own review fields carry, as slot keys. */
const reviewKeysOf = (review: ParagraphReviewFields): string[] =>
  identityKeysIn({ propertyChanges: review.propertyChanges, pPrMark: review.pPrMark });

const setParagraphReview = (document: Document, op: SetParagraphReviewOp): Applied => {
  const located = locate(document, op, op.story, op.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  if (!equalForStaleness(reviewFieldsOf(paragraph), op.expected)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${op.blockId} states other review fields than expected.`,
    );
  }
  const review = structuredClone(op.review);
  const next = withReviewFields(paragraph, review);
  if (structurallyEqual(next, paragraph)) {
    return unchanged(document);
  }
  if (
    review.pPrMark !== undefined &&
    !structurallyEqual(review.pPrMark, paragraph.pPrMark) &&
    endsItsContainer(storyBody(document, op.story), located.value) &&
    !(
      located.value.list.some((step) => step.kind === "tableCell") &&
      (review.pPrMark.kind === "ins" || review.pPrMark.kind === "del")
    )
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
      `${op.blockId} ends its container: there is no next paragraph its mark could join.`,
    );
  }
  const taken = identitiesOutside(document, [paragraph]);
  for (const key of identityKeysIn(paragraph.content)) taken.add(key);
  const incoming = reviewKeysOf(review);
  if (incoming.some((key) => taken.has(key)) || new Set(incoming).size !== incoming.length) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "The review fields carry a revision id already used in the package.",
    );
  }
  return Result.ok({
    document: replaceParagraphs({
      document,
      story: op.story,
      at: located.value,
      count: 1,
      replacement: [next],
    }),
    inverse: [reviewRestoring(op.story, next, paragraph)],
    touched: touchedBetween([paragraph], [next]),
  });
};

const replaceInline = (document: Document, op: ReplaceInlineOp): Applied => {
  const located = locate(document, op, op.story, op.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  if (!equalForStaleness(paragraph.content, op.expected)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      `${op.blockId} holds other content than expected.`,
    );
  }
  const content = structuredClone(op.content);
  if (holdsEmptyRecord(content)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
      "The content holds an empty run, text node, or revision wrapper.",
    );
  }
  if (textsIn(content).some(hasIllegalXmlCharacters)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      "The content holds text that cannot be written.",
    );
  }
  if (structurallyEqual(content, paragraph.content)) {
    return unchanged(document);
  }
  const remaining = countIds(packageParagraphIds(document.package));
  for (const id of paragraphIdsIn(paragraph.content)) {
    remaining.set(idKey(id), (remaining.get(idKey(id)) ?? 1) - 1);
  }
  if (collides(remaining, paragraphIdsIn(content))) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "The content holds a paragraph id the package already uses.",
    );
  }
  const taken = identitiesOutside(document, [paragraph]);
  for (const key of reviewKeysOf(reviewFieldsOf(paragraph))) taken.add(key);
  const incoming = identityKeysIn(content);
  if (incoming.some((key) => taken.has(key)) || new Set(incoming).size !== incoming.length) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "The content carries a revision or content-control id already used in the package.",
    );
  }
  const next = withContent(paragraph, content);
  return Result.ok({
    document: replaceParagraphs({
      document,
      story: op.story,
      at: located.value,
      count: 1,
      replacement: [next],
    }),
    inverse: [contentRestoring(op.story, next, paragraph)],
    touched: touchedBetween([paragraph], [next]),
  });
};

const dispatch = (document: Document, op: DocumentOp): Applied => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
      return deleteBlocks({ document, op, applyOps: applyDocumentOps });
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return insertBlocks({ document, op, applyOps: applyDocumentOps });
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
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
      return setParagraphReview(document, op);
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return replaceInline(document, op);
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
      return resolveRevision(document, op, applyDocumentOps);
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
      return applyTableOp(document, op);
    case DOCUMENT_OP_TYPES.INSERT_ROW:
    case DOCUMENT_OP_TYPES.DELETE_ROW:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
      return applyRowOp(document, op);
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

/** The stamp of a tracked operation; `undefined` for a direct one. */
export const stampOf = (op: DocumentOp): RevisionStamp | undefined => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_ROW:
    case DOCUMENT_OP_TYPES.DELETE_ROW:
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
      return op.revision;
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
      return undefined;
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

/** The revision ids a tracked operation's records took. */
const recordedRevisions = (
  before: Document,
  edit: DocumentEdit,
  stamp: RevisionStamp,
): number[] => {
  const touched = new Set([...edit.touched.modified, ...edit.touched.inserted].map(idKey));
  const paragraphs = storyParagraphs(edit.document.package.document)
    .map(({ paragraph }) => paragraph)
    .filter(({ paraId }) => paraId !== undefined && touched.has(idKey(paraId)));
  const known = new Set(packageIdentityKeys(before.package));
  const revisions = stampedRevisionIds(paragraphs, stamp, known);
  // Every physical row record is reported, including later rows of a table.
  revisions.unshift(
    ...stampedTableRowRevisionIds(edit.document.package.document.content, stamp, known),
  );
  return revisions;
};

/** Apply one operation to a document that meets the seed contract. */
export const applyDocumentOp = (
  document: Document,
  op: DocumentOp,
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  const valid = validateOpsDocument(document);
  if (valid.isErr()) {
    return Result.err(refusal(op, valid.error.reason, valid.error.message));
  }
  const stamp = stampOf(op);
  const badStamp = stamp === undefined ? undefined : stampRefusal(document, op, stamp);
  if (badStamp !== undefined) {
    return Result.err(badStamp);
  }
  const applied = dispatch(document, op);
  if (applied.isErr()) {
    return Result.err(applied.error);
  }
  meetsContract(applied.value.document);
  return Result.ok({
    ...applied.value,
    revisions: stamp === undefined ? [] : recordedRevisions(document, applied.value, stamp),
  });
};

/**
 * Apply operations in order, atomically: the first refusal is returned and
 * the document is unchanged. The inverse undoes the whole list, and
 * `revisions` lists each operation's in turn.
 */
export const applyDocumentOps = (
  document: Document,
  ops: readonly DocumentOp[],
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  let current = document;
  const edits: DocumentEdit[] = [];
  const revisions: number[] = [];
  for (const op of ops) {
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) {
      return applied;
    }
    current = applied.value.document;
    edits.push(applied.value);
    revisions.push(...applied.value.revisions);
  }
  return Result.ok({ ...combineEdits(document, edits), revisions });
};
