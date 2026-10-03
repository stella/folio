import { cloneModel } from "./modelClone";
import { noteContentWithAutomaticMark } from "./noteMarks";
import { applyStoryLifecycle, storyLifecycleEdit } from "./storyLifecycle";
import { documentStories, findStoryBody, sameStory } from "./stories";
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
 * | `splitBlock`        | `joinBlocks`, then authored content and review fields    |
 * | `joinBlocks`        | `splitBlock` with the retired paragraph's fields, then  |
 * |                     | authored content and the survivor's review fields         |
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

import { panic, Result } from "better-result";

import {
  type Document,
  type Insertion,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
  type ParagraphFormatting,
  type Run,
  type RunPropertyChange,
} from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import { deleteBlocks } from "./blockDeletion";
import { insertBlocks } from "./blockInsertion";
import {
  endsItsContainer,
  captureSectionView,
  type ParagraphLocation,
  replaceParagraphs,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { captureDocumentOp, restoreDocumentOp } from "./wire";
import {
  hasMultipleParagraphPropertyChanges,
  meetsContract,
  validateOpsDocument,
} from "./contract";
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
  reservedIdentityKeysIn,
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
  joinParagraphSeam,
  namesIds,
  patchedSet,
  patchRunsBetween,
  type RunDecoration,
  runsBetween,
  splitAt,
} from "./inline";
import {
  asParagraphContent,
  leafSpans,
  childNodes,
  defaultInsertionGap,
  type Gap,
  type InlineNode,
  isEmptyRecord,
  recordsBetween,
  rebuildNode,
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
import { applyNumberingSectionOp } from "./numberingSections";
import { applyTableOp } from "./tables";
import { stampedTableRowRevisionIds } from "./tableTracking";
import {
  namesMarkFormatting,
  paragraphPropertiesOf,
  foldedParagraphPropertyChange,
  paragraphPropertyChange,
  reviewFieldsOf,
  sameMarkFormatting,
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
  PROPERTY_REVIEW_POLICIES,
  DOCUMENT_OP_SCHEMA_VERSION,
  SECTION_BOUNDARY_POLICIES,
  type AddNoteOp,
  type RemoveNoteOp,
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
   * Newly allocated tracked revision ids, in document order: its stamp's
   * id and the new ids its other records took. Existing ids retained by a
   * folded or continued review are excluded; a direct operation reports none.
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
): DocumentOpRefusal => new DocumentOpRefusal({ message, reason, opType: op?.type });

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

type CutHasResolutionProvenanceOptions = {
  content: readonly ParagraphContent[];
  gaps: readonly Gap[];
};

// Repartitioned source-slot references differ between the two cut pieces.
// A compact seam merge cannot restore the original reference object exactly.
const cutHasResolutionProvenance = ({
  content,
  gaps,
}: CutHasResolutionProvenanceOptions): boolean =>
  gaps.some((gap) =>
    spanningRecords(content, [gap], 0, 1).some(
      (record) =>
        (record.type === "insertion" ||
          record.type === "deletion" ||
          record.type === "moveFrom" ||
          record.type === "moveTo") &&
        record.resolutionJoins !== undefined,
    ),
  );

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
  if (!sameStory(from.story, to.story) || idKey(from.blockId) !== idKey(to.blockId)) {
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
    reserved: new Set(reservedIdentityKeysIn(document.package)),
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
  const replaced = replaceParagraphs({
    document,
    story,
    at,
    count: before.length,
    replacement: paragraphs.value,
    ...("sectionView" in options.op && options.op.sectionView !== undefined
      ? { restoreSections: options.op.sectionView.restore }
      : {}),
  });
  if (replaced.isErr())
    return Result.err(refusal(options.op, replaced.error.reason, replaced.error.message));
  return Result.ok({
    document: replaced.value,
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
    resolutionJoin?: number;
    acceptance?: NonNullable<Insertion["resolutionJoins"]>["acceptance"];
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
    const inverse = insertionInverse(location.paragraph.content, content, start, end);
    if (inverse === undefined) {
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
      resolutionJoin: inverse.join + spanningRecords(content, [start, end], 0, 2).length,
      ...(op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT && op.seamPolicy !== undefined
        ? { acceptance: op.seamPolicy }
        : {}),
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
      if (cutHasResolutionProvenance({ content: location.paragraph.content, gaps: [start] }))
        return Result.ok([contentRestoring(story, result, location.paragraph)]);
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
  const slice = { ...op.slice, content: cloneModel(op.slice.content) };
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
      resolutionJoin: spanningRecords(paragraph.content, [from, to], 0, 2).length,
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
      if (cutHasResolutionProvenance({ content: paragraph.content, gaps: [from, to] }))
        return Result.ok([contentRestoring(story, result, paragraph)]);
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
    (result) => {
      if (cutHasResolutionProvenance({ content: location.paragraph.content, gaps: [gap] }))
        return Result.ok([contentRestoring(op.at.story, result, location.paragraph)]);
      return Result.ok([
        {
          type: DOCUMENT_OP_TYPES.JOIN_INLINE,
          at: positionAt(op.at.story, result, gap),
          depth: op.depth,
        },
      ]);
    },
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
      if (cutHasResolutionProvenance({ content: result.content, gaps: [gap] }))
        return Result.ok([contentRestoring(op.at.story, result, location.paragraph)]);
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

/** Compare the stated values and own presence; null expects an omitted key. */
const statesValues = (formatting: object | undefined, expected: object): boolean => {
  return Object.entries(expected).every(([key, value]) => {
    const owns = formatting !== undefined && Object.hasOwn(formatting, key);
    if (value === null) return !owns;
    if (formatting === undefined || !owns) return false;
    return structurallyEqual(Reflect.get(formatting, key), value);
  });
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
    // Editor actions append separately rejectable records; low-level patches
    // retain the pending review's original formatting unless append is explicit.
    const changes = new Map<Run, Parameters<RunDecoration>[0]>();
    const recordChange: RunDecoration = (options) => {
      const { patched: patchedRun, previous } = options;
      if (
        (previous.propertyChanges?.length ?? 0) > 0 &&
        op.propertyReview !== PROPERTY_REVIEW_POLICIES.APPEND
      ) {
        return patchedRun;
      }
      changes.set(patchedRun, options);
      return patchedRun;
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
    // Source records cut by the patch consume their identities before this
    // action introduces any new property-change records.
    const freshened = withFreshIds({
      document,
      op,
      before: [paragraph],
      after: [withContent(paragraph, tracked.content)],
      newIds: op.newIds,
    });
    if (freshened.isErr()) return Result.err(freshened.error);
    const appendChanges = (
      source: readonly InlineNode[],
      fresh: readonly InlineNode[],
    ): InlineNode[] => {
      if (source.length !== fresh.length) panic("Freshening changed the run patch's structure.");
      return fresh.map((node, index) => {
        const original = source.at(index);
        if (original === undefined || original.type !== node.type)
          panic("Freshening changed the run patch's node kind.");
        if (node.type === "run" && original.type === "run") {
          const options = changes.get(original);
          if (options === undefined) return node;
          const change: RunPropertyChange = {
            type: "runPropertyChange",
            info: stampInfo(revision),
            boundaryJoins: options.boundaryJoins,
          };
          if (options.previous.formatting !== undefined)
            change.previousFormatting = options.previous.formatting;
          const changed = Object.assign({}, node);
          changed.propertyChanges = (node.propertyChanges ?? []).concat(change);
          return changed;
        }
        const children = childNodes(node);
        if (children === undefined) return node;
        const priorChildren = childNodes(original);
        if (priorChildren === undefined) panic("Freshening removed a run patch's child list.");
        const next = appendChanges(priorChildren, children);
        return next.every((child, childIndex) => child === children[childIndex])
          ? node
          : rebuildNode(node, next);
      });
    };
    const sourceFresh = freshened.value.at(0);
    if (sourceFresh === undefined) panic("Freshening removed the run patch's paragraph.");
    const content = asParagraphContent(appendChanges(tracked.content, sourceFresh.content));
    return editOne(
      {
        document,
        op,
        story,
        at: location,
        paragraph: withContent(paragraph, content),
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
  const replaced = replaceParagraphs({
    document,
    story,
    at: location,
    count: 1,
    replacement: [result],
  });
  if (replaced.isErr())
    return Result.err(refusal(op, replaced.error.reason, replaced.error.message));
  return Result.ok({
    document: replaced.value,
    inverse: cutHasResolutionProvenance({ content: paragraph.content, gaps: [from, to] })
      ? [contentRestoring(story, result, paragraph)]
      : inverse,
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
  if (structurallyEqual(formatting, paragraph.formatting)) {
    return unchanged(document);
  }
  if (revision !== undefined) {
    const next = withParagraphFormatting(paragraph, formatting);
    next.propertyChanges = [
      foldedParagraphPropertyChange({ paragraph, formatting, stamp: revision }),
    ];
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
  /** Whether the new half is the second, which ends with the paragraph's own mark. */
  madeSecond: boolean;
  resolutionJoin: number;
};

/**
 * Record a split as tracked: the new mark is an insertion, and the new
 * paragraph records a property change from the source's properties when it
 * states other ones and carries no change already. Returns the refusal when
 * the operation states a mark or property change of its own, or other run
 * properties for the paragraph's own mark: no tracked change records those.
 */
const trackSplit = ({
  op,
  stamp,
  paragraph,
  first,
  made,
  madeSecond,
  resolutionJoin,
}: TrackSplitOptions): DocumentOpRefusal | undefined => {
  if (op.firstMark !== undefined || (op.newParagraph?.propertyChanges?.length ?? 0) > 0) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      "A tracked split records its own mark and property change.",
    );
  }
  // A new second half ends with the paragraph's mark, which rejecting the split keeps.
  if (madeSecond && !sameMarkFormatting(made.formatting, paragraph.formatting)) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A paragraph property change does not record the paragraph mark's run properties.",
    );
  }
  first.pPrMark = { kind: "ins", info: stampInfo(stamp), resolutionJoin };
  if (!sameParagraphProperties(made.formatting, paragraph.formatting)) {
    made.propertyChanges = [
      madeSecond
        ? foldedParagraphPropertyChange({ paragraph, formatting: made.formatting, stamp })
        : paragraphPropertyChange(stampInfo(stamp), paragraph.formatting),
    ];
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
  if (op.firstSectionProperties !== undefined) first.sectionProperties = op.firstSectionProperties;
  if (op.revision !== undefined) {
    const madeHalf = newHalf === SPLIT_HALVES.FIRST ? first : second;
    const tracked = trackSplit({
      op,
      stamp: op.revision,
      paragraph,
      first,
      made: madeHalf,
      madeSecond: newHalf === SPLIT_HALVES.SECOND,
      resolutionJoin: cut.through.length,
    });
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
  const snapshotCut = cutHasResolutionProvenance({ content: paragraph.content, gaps: [gap] });
  const contentInverse = snapshotCut
    ? [
        contentRestoring(
          op.at.story,
          withContent(kept, [...placedFirst.content, ...placedSecond.content]),
          paragraph,
        ),
      ]
    : [];
  const inverseContent =
    cut.through.length === 0
      ? joinParagraphSeam(placedFirst.content, placedSecond.content)
      : paragraph.content;
  return Result.ok({
    document: committed.value.document,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: op.at.story,
        blockId: placedFirst.paraId ?? "",
        nextBlockId: placedSecond.paraId ?? "",
        depth: snapshotCut ? 0 : cut.through.length,
        survivor: newHalf === SPLIT_HALVES.FIRST ? SPLIT_HALVES.SECOND : SPLIT_HALVES.FIRST,
        expectedRetired: splitFieldsOf(madePlaced),
        expectedSurvivor: reviewFieldsOf(kept),
        ...(first.sectionProperties === undefined
          ? {}
          : { sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE }),
      },
      ...(structurallyEqual(inverseContent, paragraph.content)
        ? []
        : [contentRestoring(op.at.story, { ...kept, content: inverseContent }, paragraph)]),
      ...reviewSetting(
        op.at.story,
        kept.paraId ?? keptId,
        joinedReview(placedFirst, placedSecond),
        reviewFieldsOf(paragraph),
      ),
      ...contentInverse,
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
  if (depth !== 0) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A tracked paragraph mark cannot record an inline merge depth.",
    );
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
  // Primitive direct and tracked joins share physical paragraph semantics.
  // The editor intent compiler applies accepted-view group formatting.
  const formatting = joinedFormatting(leading, trailing);
  if (!sameParagraphProperties(formatting, trailing.formatting)) {
    second = withParagraphFormatting(trailing, formatting);
    second.propertyChanges = [
      foldedParagraphPropertyChange({ paragraph: trailing, formatting, stamp }),
    ];
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
  if (
    leading.sectionProperties !== undefined &&
    op.sectionBoundary !== SECTION_BOUNDARY_POLICIES.REMOVE
  ) {
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
  if (leading.sectionProperties !== undefined)
    split.firstSectionProperties = leading.sectionProperties;
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
  const contentInverse = [];
  if (
    depth === 0 &&
    merged.value.content.length < leading.content.length + trailing.content.length
  ) {
    const restoredContent = cutAt(merged.value.content, {
      offset: length,
      zeroWidthBefore: zeroWidthLeavesAt(leading.content, length).length,
    });
    if (!structurallyEqual(restoredContent.before, leading.content)) {
      contentInverse.push(
        contentRestoring(op.story, { ...leading, content: restoredContent.before }, leading),
      );
    }
    if (!structurallyEqual(restoredContent.after, trailing.content)) {
      contentInverse.push(
        contentRestoring(op.story, { ...trailing, content: restoredContent.after }, trailing),
      );
    }
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
  const joinedGap = {
    offset: length,
    zeroWidthBefore: zeroWidthLeavesAt(leading.content, length).length,
  };
  if (cutHasResolutionProvenance({ content: joined.content, gaps: [joinedGap] })) {
    return Result.ok({
      document: committed.value.document,
      inverse: [
        {
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story: op.story,
          expected: committed.value.paragraphs,
          blocks: [leading, trailing],
          ...(leading.sectionProperties === undefined
            ? {}
            : { sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE }),
        },
      ],
      touched: committed.value.touched,
    });
  }
  return Result.ok({
    document: committed.value.document,
    inverse: [
      split,
      ...contentInverse,
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
  if (
    op.sectionBoundaries !== SECTION_BOUNDARY_POLICIES.REPLACE &&
    !sameSectionBreaks(before, op.blocks)
  ) {
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
  const replaced = replaceParagraphs({
    document,
    story: op.story,
    at: start,
    count: before.length,
    replacement: op.blocks,
    ...(op.sectionView === undefined ? {} : { restoreSections: op.sectionView.restore }),
  });
  if (replaced.isErr())
    return Result.err(refusal(op, replaced.error.reason, replaced.error.message));
  return Result.ok({
    document: replaced.value,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: op.story,
        expected: op.blocks,
        blocks: before,
        ...(op.sectionBoundaries === SECTION_BOUNDARY_POLICIES.REPLACE
          ? { sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE }
          : {}),
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
  const review = cloneModel(op.review);
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
  const replaced = replaceParagraphs({
    document,
    story: op.story,
    at: located.value,
    count: 1,
    replacement: [next],
  });
  if (replaced.isErr())
    return Result.err(refusal(op, replaced.error.reason, replaced.error.message));
  return Result.ok({
    document: replaced.value,
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
  const content = cloneModel(op.content);
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
  const replaced = replaceParagraphs({
    document,
    story: op.story,
    at: located.value,
    count: 1,
    replacement: [next],
  });
  if (replaced.isErr())
    return Result.err(refusal(op, replaced.error.reason, replaced.error.message));
  return Result.ok({
    document: replaced.value,
    inverse: [contentRestoring(op.story, next, paragraph)],
    touched: touchedBetween([paragraph], [next]),
  });
};

const editNote = (document: Document, op: AddNoteOp | RemoveNoteOp): Applied => {
  const kind = op.type === DOCUMENT_OP_TYPES.ADD_NOTE ? op.note.type : op.story.kind;
  const id = op.type === DOCUMENT_OP_TYPES.ADD_NOTE ? op.note.id : op.story.id;
  if (!Number.isInteger(id) || id <= 0)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      "A normal note needs a positive integer id.",
    );
  const collection = kind === "footnote" ? document.package.footnotes : document.package.endnotes;
  if (op.type === DOCUMENT_OP_TYPES.ADD_NOTE) {
    if (collection?.some((note) => note.id === id))
      return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION, "The note id is already used.");
    if (op.note.noteType !== undefined && op.note.noteType !== "normal")
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "A reference can only create a normal note.",
      );
    if (op.note.content.length === 0)
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT,
        "A note needs an addressable paragraph.",
      );
    const note = cloneModel(op.note);
    const paragraphs = storyParagraphs({ content: note.content });
    const markers = paragraphs
      .flatMap(({ paragraph }) => leafSpans(paragraph.content))
      .filter(({ node }) => node.type === "noteMarker");
    if (
      markers.length > 1 ||
      markers.some(({ node }) => node.type === "noteMarker" && node.kind !== kind)
    )
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "A note needs one matching automatic reference mark.",
      );
    note.content = noteContentWithAutomaticMark({ note, customMark: false });
    const withNote: Document =
      note.type === "footnote"
        ? {
            ...document,
            package: {
              ...document.package,
              footnotes: [...(document.package.footnotes ?? []), note],
            },
          }
        : {
            ...document,
            package: {
              ...document.package,
              endnotes: [...(document.package.endnotes ?? []), note],
            },
          };
    const valid = validateOpsDocument(withNote);
    if (valid.isErr()) return refuse(op, valid.error.reason, valid.error.message);
    const reference = insertContent(withNote, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: op.at,
      slice: {
        content: [
          {
            type: "run",
            content: [{ type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id }],
          },
        ],
        openStart: 0,
        openEnd: 0,
      },
    });
    if (reference.isErr()) return reference;
    return storyLifecycleEdit(document, reference.value.document, op);
  }
  const note = collection?.find((candidate) => candidate.id === id);
  if (!note)
    return refuse(op, DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, "The note does not exist.");
  const located = locateRange(document, op, op.at, { ...op.at, offset: op.at.offset + 1 });
  if (located.isErr()) return Result.err(located.error);
  const matching = leafSpans(located.value.location.paragraph.content).some(
    ({ node, before, after }) =>
      before.offset === op.at.offset &&
      after.offset === op.at.offset + 1 &&
      ((kind === "footnote" && node.type === "footnoteRef") ||
        (kind === "endnote" && node.type === "endnoteRef")) &&
      node.id === id,
  );
  if (!matching)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "The position does not name the note reference.",
    );
  const removed = deleteRange(document, {
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: op.at,
    to: { ...op.at, offset: op.at.offset + 1 },
  });
  if (removed.isErr()) return removed;
  const stillReferenced = documentStories(removed.value.document).some((story) => {
    if (typeof story !== "string" && story.kind === kind && "id" in story && story.id === id)
      return false;
    return storyParagraphs(storyBody(removed.value.document, story)).some(({ paragraph }) =>
      leafSpans(paragraph.content).some(
        ({ node }) =>
          ((kind === "footnote" && node.type === "footnoteRef") ||
            (kind === "endnote" && node.type === "endnoteRef")) &&
          node.id === id,
      ),
    );
  });
  if (stillReferenced)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Another reference still owns this note.",
    );
  const next =
    kind === "footnote"
      ? {
          ...removed.value.document,
          package: {
            ...removed.value.document.package,
            footnotes:
              removed.value.document.package.footnotes?.filter(
                (candidate) => candidate.id !== id,
              ) ?? [],
          },
        }
      : {
          ...removed.value.document,
          package: {
            ...removed.value.document.package,
            endnotes:
              removed.value.document.package.endnotes?.filter((candidate) => candidate.id !== id) ??
              [],
          },
        };
  return storyLifecycleEdit(document, next, op);
};

const dispatch = (document: Document, op: DocumentOp): Applied => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
      return applyStoryLifecycle(document, op);
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
      return editNote(document, op);
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
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return applyNumberingSectionOp(document, op);
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
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return undefined;
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

/** Newly allocated revision ids; folding changes an existing record in place. */
const recordedRevisions = (
  before: Document,
  edit: DocumentEdit,
  stamp: RevisionStamp,
): number[] => {
  const touched = new Set([...edit.touched.modified, ...edit.touched.inserted].map(idKey));
  const paragraphs = documentStories(edit.document)
    .flatMap((story) => storyParagraphs(storyBody(edit.document, story)))
    .map(({ paragraph }) => paragraph)
    .filter(({ paraId }) => paraId !== undefined && touched.has(idKey(paraId)));
  const known = new Set(packageIdentityKeys(before.package));
  const revisions = paragraphs.flatMap((paragraph) =>
    stampedRevisionIds([paragraph], stamp, known),
  );
  // Every physical row record is reported, including later rows of a table.
  revisions.unshift(
    ...documentStories(edit.document)
      .map((story) =>
        stampedTableRowRevisionIds(storyBody(edit.document, story).content, stamp, known),
      )
      .flat(),
  );
  return revisions;
};

/** Apply one operation to a document that meets the seed contract. */
export const applyDocumentOp = (
  document: Document,
  input: DocumentOp,
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  const restored = restoreDocumentOp(input);
  if (restored.isErr()) return restored;
  const op = restored.value;
  const valid = validateOpsDocument(document);
  if (valid.isErr()) {
    return Result.err(refusal(op, valid.error.reason, valid.error.message));
  }
  if ("sectionView" in op && op.sectionView !== undefined) {
    const sections = document.package.document.sections;
    if (
      sections === undefined ||
      !structurallyEqual(captureSectionView(sections), op.sectionView.expected)
    )
      return Result.err(
        refusal(op, DOCUMENT_OP_REFUSAL_REASONS.STALE, "The derived section metadata changed."),
      );
  }
  const addressed = (() => {
    if ("at" in op && typeof op.at === "object" && "story" in op.at) return op.at.story;
    if ("from" in op) return op.from.story;
    if ("story" in op) return op.story;
    return undefined;
  })();
  if (
    addressed !== undefined &&
    op.type !== DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER &&
    !findStoryBody(document, addressed)
  ) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "The addressed story does not exist.",
      ),
    );
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
  const edit = applied.value;
  const beforeSections = document.package.document.sections;
  const afterSections = edit.document.package.document.sections;
  const structural =
    op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK ||
    op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS ||
    op.type === DOCUMENT_OP_TYPES.REPLACE_BLOCKS;
  const unchangedSectionMetadata =
    !structural ||
    beforeSections === afterSections ||
    (beforeSections !== undefined &&
      afterSections !== undefined &&
      beforeSections.length === afterSections.length &&
      beforeSections.every((section, index) => {
        const next = afterSections.at(index);
        return (
          next !== undefined &&
          section.properties === next.properties &&
          section.headers === next.headers &&
          section.footers === next.footers &&
          Object.hasOwn(section, "headers") === Object.hasOwn(next, "headers") &&
          Object.hasOwn(section, "footers") === Object.hasOwn(next, "footers")
        );
      }));
  if (
    structural &&
    !unchangedSectionMetadata &&
    beforeSections !== undefined &&
    afterSections !== undefined
  ) {
    const restore = captureSectionView(beforeSections);
    const expected = captureSectionView(afterSections);
    if (!structurallyEqual(restore, expected)) {
      const first = edit.inverse.at(0);
      if (first === undefined) panic("A section edit produced no inverse.");
      switch (first.type) {
        case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
        case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
        case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
          edit.inverse = [
            { ...first, sectionView: { expected, restore } },
            ...edit.inverse.slice(1),
          ];
          break;
        default:
          panic("A section edit produced a non-structural inverse.");
      }
    }
  }
  // Check all operation outputs before trusting them, including caller-supplied block/review records.
  if (hasMultipleParagraphPropertyChanges(edit.document)) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "A paragraph cannot carry more than one property change.",
      ),
    );
  }
  meetsContract(edit.document);
  return Result.ok({
    ...edit,
    inverse: edit.inverse.map(captureDocumentOp),
    revisions: stamp === undefined ? [] : recordedRevisions(document, edit, stamp),
  });
};

/** Apply a journal envelope only when its schema matches this reader. */
export const applyDocumentOpEnvelope = (
  document: Document,
  envelope: { schema: number; op: DocumentOp },
): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  if (envelope.schema !== DOCUMENT_OP_SCHEMA_VERSION)
    return Result.err(
      refusal(
        envelope.op,
        DOCUMENT_OP_REFUSAL_REASONS.UNSUPPORTED_SCHEMA,
        `Document operation schema ${envelope.schema} is unsupported; this reader accepts ${DOCUMENT_OP_SCHEMA_VERSION}.`,
      ),
    );
  return applyDocumentOp(document, envelope.op);
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
