import { panic, Result } from "better-result";

import {
  type BlockContent,
  type Document,
  finalizeParagraphFormatting,
  type Insertion,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
  type ParagraphFormatting,
  type Run,
  type RunPropertyChange,
} from "../model/document";
import { getSingularRunPropertyChange } from "../model/runPropertyReview";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import {
  deleteBlocks,
  markBeforeTableResolves,
  paragraphTableRunInBoundary,
} from "./blockDeletion";
import { insertBlocks } from "./blockInsertion";
import {
  blockListAt,
  blockListEndsInParagraph,
  captureSectionView,
  hasContainerFinalParagraphMark,
  type ParagraphLocation,
  replaceParagraphs,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { applyCommentOp } from "./comments";
import { setContentControl } from "./contentControls";
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
 * | `insertText`        | `deleteRange`, or `replaceInline` for source-cut provenance |
 * | `insertContent`     | `deleteRange`, or `replaceInline` for inserted range boundaries |
 * | `deleteRange`       | `insertContent`, or `replaceInline` for noninsertable source |
 * | `splitInline`       | `joinInline`                                              |
 * | `joinInline`        | `splitInline`                                             |
 * | `setRunProps`       | `setRunProps` per stretch of prior values, joining cuts   |
 * | `setParagraphProps` | `setParagraphProps` with the prior values                 |
 * | `splitBlock`        | `joinBlocks`, then authored content and review fields    |
 * | `joinBlocks`        | `splitBlock` with the retired paragraph's fields, then  |
 * |                     | authored content and the survivor's review fields         |
 * | `insertBlocks`      | an anchored `replaceBlocks`                              |
 * | `replaceBlocks`     | `replaceBlocks`, or `insertBlocks` for an empty replacement |
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
import {
  hasMultipleParagraphPropertyChanges,
  meetsContract,
  repeatedParagraphId,
  validateOpsDocument,
} from "./contract";
import {
  cloneIncomingCopyBoundaries,
  copyOriginBoundaryKey,
  copyOriginBoundaryKeysInBlocks,
  hasValidCopyOriginAncestors,
} from "./copyOriginAncestors";
import { keptDeletionBoundaries } from "./deletionBoundaries";
import {
  combineEditEffects,
  combineEdits,
  type DocumentEdit,
  type DocumentEditEffects,
} from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import { resolveGap, type ZeroWidthDefault } from "./gaps";
import { freshenIdentities, rebindRetainedIdentities } from "./identity";
import {
  countKeys,
  IDENTITY_SPACES,
  identityKeysIn,
  idKey,
  isParaId,
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
  packageReservedKeyIndex,
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
  joinAt,
  joinContent,
  type Joined,
  joinParagraphSeam,
  namesIds,
  patchedSet,
  patchRunsBetween,
  type RunDecoration,
  runsBetween,
  sameRunFormatting,
  splitAt,
} from "./inline";
import {
  asParagraphContent,
  compareGaps,
  type Gap,
  type InlineNode,
  leafSpans,
  nestedRecords,
  recordsBetween,
  spanningRecords,
  textsIn,
  withNestedRecords,
  zeroWidthLeavesAt,
} from "./leaves";
import { changedListBounds } from "./listSplice";
import { cloneModel } from "./modelClone";
import { replaceNoteParts } from "./noteLifecycle";
import { noteContentWithAutomaticMark, removedNoteStories } from "./noteMarks";
import { applyNumberingSectionOp } from "./numberingSections";
import { paragraphLength } from "./offsets";
import { applyPackageResourcesOp } from "./packageResources";
import { findStoryParagraph } from "./paragraphLocator";
import {
  contentRestoring,
  holdsEmptyRecord,
  identitiesOutside,
  nonemptyBlockReplacementEffects,
  paragraphIdsCollide,
  prepareResolutionParagraphOp,
  type ResolutionParagraphOp,
  reviewRestoring,
  touchedBetween,
  validateBlockReplacement,
  withContent,
} from "./paragraphPrimitive";
import {
  authoredParagraphFormatting,
  changesRecordingAuthored,
  getParagraphPropertySnapshot,
  registerParagraphPropertySnapshot,
  transferParagraphPropertySnapshot,
} from "./paragraphPropertySnapshot";
import { paragraphSnapshotTransfers } from "./paragraphSnapshotTransfers";
import { priorValues } from "./patch";
import { insertsInvertedRange, rangeBoundariesIn } from "./rangeBoundaries";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
  pendingParagraphMarkReason,
} from "./refusal";
import { type ApplyOps, resolveRevision } from "./resolve";
import {
  FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
  foldedParagraphPropertyChange,
  foldedRevisionIds,
  foldsIntoOtherSuggestion,
  isRemovedRevisionNode,
  namesMarkFormatting,
  paragraphPropertiesOf,
  paragraphPropertyChange,
  reviewFieldsOf,
  RUN_FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
  sameMarkFormatting,
  sameParagraphProperties,
  stampedRevisionIds,
  stampInfo,
  takesOverOtherSuggestion,
  withMarkFormatting,
  withReviewFields,
  WRAP_KINDS,
  type WrapKind,
  wrapTracked,
} from "./review";
import { setRevisionInfo } from "./revisionInfo";
import { isOwnRecord } from "./revisionOwner";
import {
  documentStories,
  findStoryBody,
  sameStory,
  textBoxHostIds,
  type TextBoxStory,
  textBoxStoryOf,
} from "./stories";
import {
  applyStoryLifecycle,
  captureStoryParts,
  storyLifecycleEdit,
} from "./storyLifecycle";
import { tableEditBoundaryRefusal } from "./tableEditBoundary";
import { applyTableEdit } from "./tableEdits";
import { blockInsertionPointAt } from "./tableLocation";
import { ownPendingTableInsertion } from "./tableOwnPending";
import { applyRowOp } from "./tableRows";
import { applyTableOp } from "./tables";
import { stampedTableRevisionIds } from "./tableTracking";
import { namingPackageStories, textBoxRoute } from "./textBoxRouting";
import {
  type AddNoteOp,
  type DeleteRangeOp,
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertContentOp,
  type InsertTextOp,
  type JoinBlocksOp,
  type JoinInlineOp,
  type NewIds,
  type OpStory,
  type ParagraphReviewFields,
  PROPERTY_REVIEW_POLICIES,
  type RemoveNoteOp,
  type ReplaceBlocksOp,
  REVISION_DECISIONS,
  type RevisionStamp,
  SECTION_BOUNDARY_POLICIES,
  type SetParagraphPropsOp,
  type SetParagraphReviewOp,
  type SetRunPropsOp,
  SPLIT_HALVES,
  type SplitBlockOp,
  type SplitHalf,
  type SplitInlineOp,
  type SplitParagraphFields,
  type StoryParts,
  type TableInsertionPoint,
  type TextPosition,
  type TouchedBlocks,
} from "./types";
import { captureDocumentOp, restoreDocumentOp } from "./wire";

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
  /**
   * Existing revision ids whose pending property changes the operation
   * folded its edit into, in document order: the changes now carry its
   * stamp and record the properties it set, so accepting or rejecting one
   * of them accepts or rejects the edit too. Disjoint from `revisions`; a
   * direct operation reports none.
   */
  folded: readonly number[];
};

/** What one operation did, before the revisions it recorded are read off it. */
type Applied = Result<DocumentEdit, DocumentOpRefusal>;

const UNTOUCHED: TouchedBlocks = Object.freeze({
  modified: [],
  inserted: [],
  removed: [],
});

const unchanged = (document: Document): Applied =>
  Result.ok({ document, inverse: [], touched: UNTOUCHED });

const refusal = (
  op: DocumentOp,
  reason: DocumentOpRefusalReason,
  message: string,
): DocumentOpRefusal =>
  new DocumentOpRefusal({ message, reason, opType: op?.type });

const refuse = (
  op: DocumentOp,
  reason: DocumentOpRefusalReason,
  message: string,
): Applied => Result.err(refusal(op, reason, message));

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
  if (
    !Number.isInteger(revisionId) ||
    revisionId < 0 ||
    revisionId > MAX_REVISION_ID
  ) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      `${revisionId} cannot be a revision id.`,
    );
  }
  const key = slotKey({ space: IDENTITY_SPACES.REVISION, id: stamp.id });
  if (packageIdentityKeyIndex(document.package).has(key)) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      `Revision id ${stamp.id} is already used in the package.`,
    );
  }
  return undefined;
};

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

/** Characters `insertText` does not carry: each is an inline atom of its own. */
const NON_TEXT_CHARACTER_PATTERN = /[\t\n\r]/u;

/** The paragraph carrying an id, compared as hex; the contract makes it unique. */
const findParagraph = (
  op: DocumentOp,
  paragraphs: readonly ParagraphLocation[],
  blockId: string,
): Result<ParagraphLocation, DocumentOpRefusal> => {
  const key = idKey(blockId);
  const match = paragraphs.find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === key,
  );
  return match === undefined
    ? Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
          `No paragraph is ${blockId}.`,
        ),
      )
    : Result.ok(match);
};

const locate = (
  document: Document,
  op: DocumentOp,
  story: OpStory,
  blockId: string,
): Result<ParagraphLocation, DocumentOpRefusal> => {
  const location = findStoryParagraph(storyBody(document, story), blockId);
  return location === undefined
    ? Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
          `No paragraph is ${blockId}.`,
        ),
      )
    : Result.ok(location);
};

const isGap = (value: Gap | DocumentOpRefusalReason): value is Gap =>
  typeof value === "object";

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
  if (
    !sameStory(from.story, to.story) ||
    idKey(from.blockId) !== idKey(to.blockId)
  ) {
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
  const start = resolveGap({ paragraph, position: from, fallback: "afterAll" });
  const end = resolveGap({ paragraph, position: to, fallback: "beforeAll" });
  const invalid = (reason: DocumentOpRefusalReason) =>
    Result.err(
      refusal(
        op,
        reason,
        `[${from.offset}, ${to.offset}) is not a range of ${from.blockId}.`,
      ),
    );
  if (!isGap(start)) return invalid(start);
  if (!isGap(end)) return invalid(end);
  const statesZeroWidth =
    from.zeroWidthBefore !== undefined || to.zeroWidthBefore !== undefined;
  if (
    start.offset > end.offset ||
    (start.offset === end.offset &&
      start.zeroWidthBefore > end.zeroWidthBefore &&
      statesZeroWidth)
  ) {
    return invalid(DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET);
  }
  const empty =
    start.offset === end.offset && start.zeroWidthBefore >= end.zeroWidthBefore;
  return Result.ok({
    story: from.story,
    location: located.value,
    from: start,
    to: end,
    empty,
  });
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
  const gap = resolveGap({
    paragraph: located.value.paragraph,
    position: at,
    fallback,
  });
  if (!isGap(gap)) {
    return Result.err(
      refusal(op, gap, `${at.offset} is not a position in ${at.blockId}.`),
    );
  }
  return Result.ok({ location: located.value, gap });
};

/** A position naming its paragraph by the id the paragraph carries. */
const positionAt = (
  story: OpStory,
  paragraph: Paragraph,
  gap: Gap,
): TextPosition => ({
  story,
  blockId: paragraph.paraId ?? "",
  offset: gap.offset,
  zeroWidthBefore: gap.zeroWidthBefore,
});

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

type Committed = {
  document: Document;
  paragraphs: Paragraph[];
  touched: TouchedBlocks;
};

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
    reserved: packageReservedKeyIndex(document.package),
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
    return Result.err(
      refusal(options.op, replaced.error.reason, replaced.error.message),
    );
  return Result.ok({
    document: replaced.value,
    paragraphs: paragraphs.value,
    touched: touchedBetween(before, paragraphs.value),
  });
};

/** An edit of one paragraph, committed, with the inverse built from the paragraph it became. */
const editOne = (
  commitOptions: Omit<Commit, "before" | "after"> & { paragraph: Paragraph },
  inverseOf: (
    result: Paragraph,
  ) => Result<readonly DocumentOp[], DocumentOpRefusal>,
): Applied => {
  const { paragraph, at, document } = commitOptions;
  if (paragraph === at.paragraph) {
    return unchanged(document);
  }
  const committed = commit({
    ...commitOptions,
    before: [at.paragraph],
    after: [paragraph],
  });
  if (committed.isErr()) {
    return Result.err(committed.error);
  }
  const [result] = committed.value.paragraphs;
  if (result === undefined) {
    return Result.ok({
      document: committed.value.document,
      inverse: [],
      touched: UNTOUCHED,
    });
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
  copiedBoundaries?: ReadonlySet<InlineNode>;
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
    copiedBoundaries?: ReadonlySet<InlineNode>;
  },
): Wrapped => {
  const wrapped = wrapTracked(options);
  switch (wrapped.kind) {
    case "wrapped":
      // A deletion of a range holding nothing but empty hyperlinks, fields or
      // content controls marks no run: no file or reader would show it,
      // while deleting the range directly removes them.
      if (options.kind === WRAP_KINDS.DELETION && !wrapped.recordsContent)
        return Result.err(
          refusal(
            op,
            DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            "A tracked deletion cannot record only empty containers.",
          ),
        );
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
  const { document, op, story, location, start, end, newIds, revision } =
    options;
  let { content } = options;
  if (revision !== undefined) {
    // The wrap cuts the records again; an insertion the direct one refuses stays refused.
    const inverse = insertionInverse(
      location.paragraph.content,
      content,
      start,
      end,
    );
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
      resolutionJoin:
        inverse.join + spanningRecords(content, [start, end], 0, 2).length,
      ...(options.copiedBoundaries === undefined
        ? {}
        : { copiedBoundaries: options.copiedBoundaries }),
      ...(op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT &&
      op.seamPolicy !== undefined
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
      if (
        cutHasResolutionProvenance({
          content: location.paragraph.content,
          gaps: [start],
        })
      )
        return Result.ok([contentRestoring(story, result, location.paragraph)]);
      const inverse = insertionInverse(
        location.paragraph.content,
        result.content,
        start,
        end,
      );
      if (inverse === undefined) {
        return Result.err(
          refusal(
            op,
            DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
            "The insertion would merge records that were separate.",
          ),
        );
      }
      // Range boundaries belong to this insertion. Ordinary deletion keeps
      // endpoints of balanced ranges, including a pair this insertion just
      // completed; undo must restore the source, even if it was unpaired.
      if (rangeBoundariesIn(inverse.removed.content).length > 0)
        return Result.ok([contentRestoring(story, result, location.paragraph)]);
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
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      "The text cannot be inserted.",
    );
  }
  const located = locatePosition(document, op, op.at, "afterAll");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location, gap } = located.value;
  const { offset } = gap;
  const content = insertTextInContent({
    items: location.paragraph.content,
    gap,
    text: op.text,
    runProps: op.runProps,
  });
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INSIDE_TRACKED_DELETION,
      `${offset} in ${op.at.blockId} is inside tracked-removed content.`,
    );
  }
  // The text went in at the gap: the zero-width leaves it names precede it.
  const start = gap;
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
  if (!hasValidCopyOriginAncestors(slice.content)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "The slice carries invalid copied-boundary ancestry.",
    );
  }
  if (
    !Number.isInteger(slice.openStart) ||
    !Number.isInteger(slice.openEnd) ||
    slice.openStart < 0 ||
    slice.openEnd < 0
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Open depths are counts.",
    );
  }
  if (textsIn(slice.content).some(hasIllegalXmlCharacters)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      "The slice holds text that cannot be written.",
    );
  }
  const incoming = paragraphIdsIn(slice.content);
  if (paragraphIdsCollide({ document, incoming })) {
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
  const incomingBoundaryKeys = new Set(
    rangeBoundariesIn(slice.content)
      .map(copyOriginBoundaryKey)
      .filter((key): key is string => key !== undefined),
  );
  const copiedSlice =
    op.revision === undefined || incomingBoundaryKeys.size === 0
      ? { content: slice.content, copiedBoundaries: new Set<InlineNode>() }
      : cloneIncomingCopyBoundaries(
          slice.content,
          new Set(
            [
              ...copyOriginBoundaryKeysInBlocks(
                storyBody(document, op.at.story).content,
              ),
            ].filter((key) => incomingBoundaryKeys.has(key)),
          ),
        );
  const copiedBoundaries = copiedSlice.copiedBoundaries;
  const content = insertSliceAt(location.paragraph.content, gap, {
    ...slice,
    content: copiedSlice.content,
  });
  if (content === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      `The slice's open ends do not fit ${op.at.offset} in ${op.at.blockId}.`,
    );
  }
  if (
    insertsInvertedRange({
      body: storyBody(document, op.at.story),
      paragraph: location.paragraph,
      candidateContent: content,
      incomingContent: slice.content,
    })
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "The insertion would create a range whose end precedes its start.",
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
    copiedBoundaries,
  });
};

/** A merge's content, or the refusal naming why it cannot be made. */
const joinedContent = (
  op: DocumentOp,
  joined: Joined,
  where: string,
): Result<
  { content: Paragraph["content"]; retired: NewIds },
  DocumentOpRefusal
> => {
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

const isCount = (value: number): boolean =>
  Number.isInteger(value) && value >= 0;

/** Find a note story's source-owned automatic mark inside a replacement range. */
const automaticNoteMarkBetween = (
  story: OpStory,
  paragraph: Paragraph,
  from: Gap,
  to: Gap,
): { before: Gap; after: Gap } | undefined => {
  if (
    story === "main" ||
    (story.kind !== "footnote" && story.kind !== "endnote")
  )
    return undefined;
  return leafSpans(paragraph.content).find(
    ({ node, before, after }) =>
      node.type === "noteMarker" &&
      node.kind === story.kind &&
      compareGaps(from, before) <= 0 &&
      compareGaps(after, to) <= 0,
  );
};

/**
 * `items` without the leaves between `from` and `to`, except the `kept` leaves
 * inside that range, which stay where they are. Deletes right to left so each
 * cut's gaps still name the original positions.
 */
const deletedAround = (
  items: readonly ParagraphContent[],
  from: Gap,
  to: Gap,
  kept: readonly { before: Gap; after: Gap }[],
): ParagraphContent[] => {
  const cuts: [Gap, Gap][] = [];
  let start = from;
  for (const leaf of kept) {
    cuts.push([start, leaf.before]);
    start = leaf.after;
  }
  cuts.push([start, to]);
  let content = [...items];
  for (const [cutFrom, cutTo] of cuts.toReversed()) {
    if (compareGaps(cutFrom, cutTo) < 0)
      content = deleteBetween(content, cutFrom, cutTo).content;
  }
  return content;
};

const deleteRange = (document: Document, op: DeleteRangeOp): Applied => {
  const join = op.join ?? 0;
  if (!isCount(join)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A join depth is a count.",
    );
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
  const { paragraph } = location;
  const { expected } = op;
  if (empty) {
    return expected === undefined || expected.content.length === 0
      ? unchanged(document)
      : refuse(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.STALE,
          "The range to delete is empty.",
        );
  }
  const { content: remaining, removed } = deleteBetween(
    paragraph.content,
    from,
    to,
  );
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
  const protectedMark = automaticNoteMarkBetween(story, paragraph, from, to);
  // One end of a comment or move range the deletion covers, whose other end
  // stands outside it, stays where it is: the range shrinks to the
  // deletion's edge instead of losing an end.
  const keptBoundaries =
    op.revision === undefined && join === 0
      ? keptDeletionBoundaries({
          body: storyBody(document, story),
          paragraph,
          from,
          to,
        })
      : [];
  let content = remaining;
  if (protectedMark !== undefined) {
    // A leaf counts when the range covers any part of it, including part of a text leaf.
    const hasOtherSurvivingContent = leafSpans(paragraph.content).some(
      ({ node, ancestors, before, after }) =>
        compareGaps(before, to) < 0 &&
        compareGaps(from, after) < 0 &&
        node.type !== "noteMarker" &&
        !ancestors.some(isRemovedRevisionNode),
    );
    if (!hasOtherSurvivingContent) {
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
        "The range contains only the note story's automatic marker.",
      );
    }
  }
  const kept = [
    ...keptBoundaries,
    ...(protectedMark === undefined ? [] : [protectedMark]),
  ].toSorted((left, right) => compareGaps(left.before, right.before));
  if (kept.length > 0) {
    content = deletedAround(paragraph.content, from, to, kept);
    if (structurallyEqual(content, paragraph.content)) {
      return unchanged(document);
    }
  }
  if (op.revision !== undefined) {
    const wrapped = wrappedContent(op, {
      items: paragraph.content,
      from,
      to,
      kind: WRAP_KINDS.DELETION,
      stamp: op.revision,
      resolutionJoin: spanningRecords(paragraph.content, [from, to], 0, 2)
        .length,
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
  let retired: NewIds = {};
  if (join > 0) {
    const joined = joinedContent(
      op,
      joinAt(content, from, join),
      `at ${from.offset}`,
    );
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
      if (
        kept.length > 0 ||
        cutHasResolutionProvenance({
          content: paragraph.content,
          gaps: [from, to],
        }) ||
        insertsInvertedRange({
          body: storyBody(document, story),
          paragraph,
          candidateContent: paragraph.content,
          incomingContent: removed.content,
        })
      )
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
      if (
        cutHasResolutionProvenance({
          content: location.paragraph.content,
          gaps: [gap],
        })
      )
        return Result.ok([
          contentRestoring(op.at.story, result, location.paragraph),
        ]);
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
  // A seam join merges runs: it keeps records apart that are not alike or
  // carry ids, and leaves the inside of a run that already runs across.
  if (
    op.ifAlike === true &&
    (joined.isErr() ||
      namesIds(joined.value.retired) ||
      spanningRecords(location.paragraph.content, [gap], 0, 1).some(
        ({ type }) => type === "run",
      ))
  )
    return unchanged(document);
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
        return Result.ok([
          contentRestoring(op.at.story, result, location.paragraph),
        ]);
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
const statesValues = (
  formatting: object | undefined,
  expected: object,
): boolean => {
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
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A join depth is a count.",
    );
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
    // Editor actions update the pending review; low-level patches retain its authorship.
    const changes = new Map<Run, Parameters<RunDecoration>[0]>();
    let takesOverSuggestion = false;
    const recordChange: RunDecoration = (options) => {
      const { patched: patchedRun, previous } = options;
      const pending = getSingularRunPropertyChange(previous.propertyChanges);
      if (takesOverOtherSuggestion(pending, revision)) {
        takesOverSuggestion = true;
        return patchedRun;
      }
      if (
        pending !== undefined &&
        op.propertyReview !== PROPERTY_REVIEW_POLICIES.APPEND
      ) {
        return patchedRun;
      }
      // The replaced record consumes no fresh identity before it is coalesced.
      const next = Object.assign({}, patchedRun);
      delete next.propertyChanges;
      changes.set(next, options);
      return next;
    };
    const tracked = patchRunsBetween(
      paragraph.content,
      from,
      to,
      op.patch,
      op.whenEmpty,
      recordChange,
    );
    if (takesOverSuggestion)
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.PENDING_RUN_PROPERTY_CHANGE,
        RUN_FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
      );
    if (tracked === undefined) {
      return unchanged(document);
    }
    // Retained source records cut by the patch consume their identities before
    // this action introduces its coalesced property-change records.
    const freshened = withFreshIds({
      document,
      op,
      before: [paragraph],
      after: [withContent(paragraph, tracked.content)],
      newIds: op.newIds,
    });
    if (freshened.isErr()) return Result.err(freshened.error);
    const sourceFresh =
      freshened.value.at(0) ??
      panic("Freshening removed the run patch's paragraph.");
    // A pending change the patch folds whole keeps its id, as a paragraph's
    // does; a piece the patch cuts off another record carrying that id takes
    // the stamp's.
    const carriedAfterPatch = countKeys(identityKeysIn([sourceFresh]));
    const carriedOutside = identitiesOutside(document, [paragraph]);
    const foldedInfo = (pending: RunPropertyChange | undefined) => {
      const info = stampInfo(revision);
      if (pending === undefined) return info;
      const key = slotKey({
        space: IDENTITY_SPACES.REVISION,
        id: pending.info.id,
      });
      if (carriedOutside.has(key) || carriedAfterPatch.has(key)) return info;
      info.id = pending.info.id;
      return info;
    };
    const priorIdentities = new WeakMap<InlineNode, InlineNode>();
    const appendChanges = (
      source: readonly InlineNode[],
      fresh: readonly InlineNode[],
    ): InlineNode[] => {
      if (source.length !== fresh.length)
        panic("Freshening changed the run patch's structure.");
      return fresh.map((node, index) => {
        const original = source.at(index);
        if (original === undefined || original.type !== node.type)
          panic("Freshening changed the run patch's node kind.");
        if (node.type === "run" && original.type === "run") {
          const options = changes.get(original);
          if (options === undefined) return node;
          const prior = Object.assign({}, node);
          if (options.previous.propertyChanges !== undefined)
            prior.propertyChanges = options.previous.propertyChanges;
          const pending = getSingularRunPropertyChange(
            options.previous.propertyChanges,
          );
          if (
            pending !== undefined &&
            sameRunFormatting(pending.previousFormatting, node.formatting)
          ) {
            // Back to the properties the pending change started from:
            // nothing is left to review, so the change goes, and the run
            // states them as rejecting it would.
            const settled = Object.assign({}, node);
            if (pending.previousFormatting === undefined)
              delete settled.formatting;
            else settled.formatting = pending.previousFormatting;
            priorIdentities.set(settled, prior);
            return settled;
          }
          const change: RunPropertyChange = {
            type: "runPropertyChange",
            info: foldedInfo(pending),
            boundaryJoins: options.boundaryJoins,
          };
          const previousFormatting =
            pending === undefined
              ? options.previous.formatting
              : pending.previousFormatting;
          if (previousFormatting !== undefined)
            change.previousFormatting = previousFormatting;
          if (pending !== undefined) {
            if (
              pending.boundaryJoins === undefined ||
              change.boundaryJoins === undefined
            ) {
              delete change.boundaryJoins;
            } else {
              change.boundaryJoins = [
                ...new Set([...pending.boundaryJoins, ...change.boundaryJoins]),
              ];
            }
          }
          const changed = Object.assign({}, node);
          changed.propertyChanges = [change];
          priorIdentities.set(changed, prior);
          return changed;
        }
        const children = nestedRecords(node);
        if (children === undefined) return node;
        const priorChildren = nestedRecords(original);
        if (priorChildren === undefined)
          panic("Freshening removed a run patch's child list.");
        const next = appendChanges(priorChildren, children);
        return next.every((child, childIndex) => child === children[childIndex])
          ? node
          : withNestedRecords(node, next);
      });
    };
    const content = asParagraphContent(
      appendChanges(tracked.content, sourceFresh.content),
    );
    const restorePriorIdentities = (
      nodes: readonly InlineNode[],
    ): InlineNode[] =>
      nodes.map((node) => {
        const prior = priorIdentities.get(node) ?? node;
        const children = nestedRecords(prior);
        return children === undefined
          ? prior
          : withNestedRecords(prior, restorePriorIdentities(children));
      });
    const rebound =
      rebindRetainedIdentities({
        before: [
          withContent(
            paragraph,
            asParagraphContent(restorePriorIdentities(content)),
          ),
        ],
        after: [withContent(paragraph, content)],
      }).at(0) ??
      panic("A run patch preserves its paragraph during identity rebinding.");
    return editOne(
      {
        document,
        op,
        story,
        at: location,
        paragraph: rebound,
        newIds: op.newIds,
      },
      (result) => Result.ok([contentRestoring(story, result, paragraph)]),
    );
  }
  const patched = patchRunsBetween(
    paragraph.content,
    from,
    to,
    op.patch,
    op.whenEmpty,
  );
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
    const joined = joinedContent(
      op,
      joinAt(content, gap, depth),
      `at ${gap.offset}`,
    );
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

  // A field whose result runs the patch changed from different values has no
  // one patch over its unit that restores them; the paragraph as it was does.
  const restorable = patched === undefined || patched.restoring !== undefined;
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
    if (namesIds(retiredAt.from))
      first.newIds = concatIds(first.newIds ?? {}, retiredAt.from);
    if (namesIds(retiredAt.to))
      last.newIds = concatIds(last.newIds ?? {}, retiredAt.to);
  }
  const replaced = replaceParagraphs({
    document,
    story,
    at: location,
    count: 1,
    replacement: [result],
  });
  if (replaced.isErr())
    return Result.err(
      refusal(op, replaced.error.reason, replaced.error.message),
    );
  return Result.ok({
    document: replaced.value,
    inverse:
      !restorable ||
      cutHasResolutionProvenance({
        content: paragraph.content,
        gaps: [from, to],
      })
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
  const finalized = finalizeParagraphFormatting(formatting);
  if (finalized === undefined) {
    delete next.formatting;
  } else {
    next.formatting = finalized;
  }
  return next;
};

const setParagraphProps = (
  document: Document,
  op: SetParagraphPropsOp,
): Applied => {
  const located = locate(document, op, op.story, op.blockId);
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { paragraph } = located.value;
  if (
    op.expected !== undefined &&
    !statesValues(paragraph.formatting, op.expected)
  ) {
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
  const formatting = finalizeParagraphFormatting(
    patchedSet(paragraph.formatting, op.patch, op.whenEmpty),
  );
  const currentFormatting = finalizeParagraphFormatting(paragraph.formatting);
  if (structurallyEqual(formatting, currentFormatting)) {
    return unchanged(document);
  }
  const next = withParagraphFormatting(paragraph, formatting);
  registerParagraphPropertySnapshot({
    paragraph: next,
    snapshot: patchedSet(
      getParagraphPropertySnapshot(paragraph),
      op.patch,
      op.whenEmpty,
    ),
  });
  if (revision !== undefined) {
    if (foldsIntoOtherSuggestion(paragraph, revision))
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
        FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
      );
    const authored = authoredParagraphFormatting(next);
    const pending = paragraph.propertyChanges;
    // Stated, not effective, properties: an explicit value still holds if the
    // style changes, so it is not the inherited default.
    if (
      pending?.length === 1 &&
      structurallyEqual(
        finalizeParagraphFormatting(pending.at(0)?.previousFormatting),
        authored,
      )
    ) {
      // Back to the properties the pending change started from: nothing is
      // left to review, so the change goes.
      delete next.propertyChanges;
    } else {
      // The change records what the paragraph now states as its current
      // properties, as a reopened file reads them.
      next.propertyChanges = [
        foldedParagraphPropertyChange({
          paragraph,
          formatting: authored,
          stamp: revision,
        }),
      ];
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
  // Edited directly, the paragraph keeps its pending property changes, each
  // now recording what the paragraph states; undoing restores them as they
  // were.
  if (paragraph.propertyChanges !== undefined) {
    next.propertyChanges = changesRecordingAuthored(
      paragraph.propertyChanges,
      next,
    );
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
      paragraph: next,
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
    : withMarkFormatting(
        paragraphPropertiesOf(first.formatting),
        second.formatting,
      );

/** The review fields of two paragraphs once joined: the first's properties, the second's mark. */
const joinedReview = (
  first: Paragraph,
  second: Paragraph,
): ParagraphReviewFields => {
  const review: ParagraphReviewFields = {};
  const formatting = joinedFormatting(first, second);
  if (formatting !== undefined) review.formatting = formatting;
  if (second.propertyChanges !== undefined)
    review.propertyChanges = second.propertyChanges;
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
  if (
    op.firstMark !== undefined ||
    (op.newParagraph?.propertyChanges?.length ?? 0) > 0
  ) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      "A tracked split records its own mark and property change.",
    );
  }
  // A new second half ends with the paragraph's mark, which rejecting the split keeps.
  if (
    madeSecond &&
    !sameMarkFormatting(made.formatting, paragraph.formatting)
  ) {
    return refusal(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      "A paragraph property change does not record the paragraph mark's run properties.",
    );
  }
  first.pPrMark = {
    kind: "ins",
    info: stampInfo(stamp),
    resolutionJoin,
    // A first half that keeps the paragraph's identity is the paragraph a
    // rejection gives back.
    ...(madeSecond ? { survivor: "first" as const } : {}),
  };
  if (!sameParagraphProperties(made.formatting, paragraph.formatting)) {
    if (madeSecond && foldsIntoOtherSuggestion(paragraph, stamp))
      return refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
        FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
      );
    made.propertyChanges = [
      madeSecond
        ? foldedParagraphPropertyChange({
            paragraph,
            formatting: made.formatting,
            stamp,
          })
        : paragraphPropertyChange(stampInfo(stamp), paragraph, made.formatting),
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
  if (packageParagraphIdIndex(document.package).has(idKey(op.newBlockId))) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      `${op.newBlockId} is already used in the package.`,
    );
  }
  if (op.newHalf !== undefined && !isSplitHalf(op.newHalf)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A half is first or second.",
    );
  }
  const located = locatePosition(document, op, op.at, "insertion");
  if (located.isErr()) {
    return Result.err(located.error);
  }
  const { location, gap } = located.value;
  const { paragraph } = location;
  const newHalf =
    op.newHalf ??
    (gap.offset === paragraphLength(paragraph)
      ? SPLIT_HALVES.SECOND
      : SPLIT_HALVES.FIRST);
  if (
    newHalf === SPLIT_HALVES.SECOND &&
    (op.newParagraph?.propertyChanges?.length ?? 0) > 0
  ) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "The second half takes the paragraph's property changes with its mark.",
    );
  }
  const cut = cutAt(paragraph.content, gap);
  const fields =
    op.newParagraph ??
    (paragraph.formatting === undefined
      ? {}
      : { formatting: paragraph.formatting });
  const made: Paragraph = {
    ...fields,
    type: "paragraph",
    paraId: op.newBlockId,
    content: [],
  };
  if (finalizeParagraphFormatting(made.formatting) === undefined)
    delete made.formatting;
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
    if (paragraph.propertyChanges !== undefined)
      second.propertyChanges = paragraph.propertyChanges;
    if (paragraph.sectionProperties !== undefined) {
      second.sectionProperties = paragraph.sectionProperties;
    }
    if (paragraph.pPrMark !== undefined) second.pPrMark = paragraph.pPrMark;
  }
  if (op.firstMark !== undefined) {
    first.pPrMark = op.firstMark;
  }
  if (op.firstSectionProperties !== undefined)
    first.sectionProperties = op.firstSectionProperties;
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
  const [placedFirst = first, placedSecond = second] =
    committed.value.paragraphs;
  const kept = newHalf === SPLIT_HALVES.FIRST ? placedSecond : placedFirst;
  const madePlaced =
    newHalf === SPLIT_HALVES.FIRST ? placedFirst : placedSecond;
  const keptId = paragraph.paraId ?? op.at.blockId;
  const snapshotCut = cutHasResolutionProvenance({
    content: paragraph.content,
    gaps: [gap],
  });
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
        survivor:
          newHalf === SPLIT_HALVES.FIRST
            ? SPLIT_HALVES.SECOND
            : SPLIT_HALVES.FIRST,
        expectedRetired: splitFieldsOf(madePlaced),
        expectedSurvivor: reviewFieldsOf(kept),
        ...(first.sectionProperties === undefined
          ? {}
          : { sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE }),
      },
      ...(structurallyEqual(inverseContent, paragraph.content)
        ? []
        : [
            contentRestoring(
              op.at.story,
              { ...kept, content: inverseContent },
              paragraph,
            ),
          ]),
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
const trackJoin = ({
  document,
  op,
  stamp,
  at,
  leading,
  trailing,
}: TrackJoinOptions): Applied => {
  const depth = op.depth ?? 0;
  if (!isCount(depth)) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A join depth is a count.",
    );
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
  const mark = leading.pPrMark;
  // Another owner's added break keeps their insertion, and the join is
  // recorded as its deletion beside it.
  const deletesAddedMark =
    (mark?.kind === "ins" || mark?.kind === "moveTo") &&
    mark.deletion === undefined &&
    !isOwnRecord(mark.info, stamp);
  if (mark !== undefined && !deletesAddedMark) {
    return refuse(
      op,
      pendingParagraphMarkReason(mark),
      `The mark of ${op.blockId} already carries a tracked change.`,
    );
  }
  // The second paragraph's own mark already goes, before a table that cannot
  // receive what runs in (its first row or cell is under review): accepting
  // both would run the first paragraph's words into that table.
  const trailingMark = trailing.pPrMark;
  const trailingMarkGoes =
    trailingMark?.kind === "del" ||
    trailingMark?.kind === "moveFrom" ||
    trailingMark?.deletion !== undefined;
  const list = blockListAt(storyBody(document, op.story).content, at.list);
  const afterTrailing = list[at.index + 2];
  const boundary =
    afterTrailing?.type === "table"
      ? paragraphTableRunInBoundary(trailing, afterTrailing)
      : undefined;
  if (
    trailingMarkGoes &&
    boundary !== undefined &&
    !markBeforeTableResolves({ blocks: list, boundary })
  ) {
    // Only a table under review is a refusal of its own; a section break or
    // a shape with no receiving paragraph is one no tracked join records yet.
    return refuse(
      op,
      boundary.type === "blocked" && boundary.reason === "review"
        ? DOCUMENT_OP_REFUSAL_REASONS.JOIN_INTO_PENDING_TABLE
        : DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      `Joining ${op.blockId} would run it into the table after ${op.nextBlockId}.`,
    );
  }
  const first: Paragraph = {
    ...leading,
    pPrMark:
      mark === undefined
        ? { kind: "del", info: stampInfo(stamp) }
        : { ...mark, deletion: { info: stampInfo(stamp) } },
  };
  let second = trailing;
  // Primitive direct and tracked joins share physical paragraph semantics.
  // The editor intent compiler applies accepted-view group formatting.
  const formatting = joinedFormatting(leading, trailing);
  if (!sameParagraphProperties(formatting, trailing.formatting)) {
    if (foldsIntoOtherSuggestion(trailing, stamp))
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
        FOLD_INTO_OTHER_SUGGESTION_MESSAGE,
      );
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
  const [placedFirst = first, placedSecond = second] =
    committed.value.paragraphs;
  const inverse: SetParagraphReviewOp[] = [
    reviewRestoring(op.story, placedFirst, leading),
  ];
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
  // An empty address historically selects the first idless paragraph.
  const unaddressed =
    op.blockId === "" || op.nextBlockId === ""
      ? storyParagraphs(storyBody(document, op.story))
      : undefined;
  const endpoint = (id: string) =>
    unaddressed === undefined
      ? locate(document, op, op.story, id)
      : findParagraph(op, unaddressed, id);
  const first = endpoint(op.blockId);
  if (first.isErr()) {
    return Result.err(first.error);
  }
  const next = endpoint(op.nextBlockId);
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
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A half is first or second.",
    );
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
    !equalForStaleness(
      reviewFieldsOf(survivor),
      reviewFieldsOf(op.expectedSurvivor),
    )
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
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "A join depth is a count.",
    );
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
  const newHalf =
    survivorSide === SPLIT_HALVES.SECOND
      ? SPLIT_HALVES.FIRST
      : SPLIT_HALVES.SECOND;
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
    merged.value.content.length <
      leading.content.length + trailing.content.length
  ) {
    const restoredContent = cutAt(merged.value.content, {
      offset: length,
      zeroWidthBefore: zeroWidthLeavesAt(leading.content, length).length,
    });
    if (!structurallyEqual(restoredContent.before, leading.content)) {
      contentInverse.push(
        contentRestoring(
          op.story,
          { ...leading, content: restoredContent.before },
          leading,
        ),
      );
    }
    if (!structurallyEqual(restoredContent.after, trailing.content)) {
      contentInverse.push(
        contentRestoring(
          op.story,
          { ...trailing, content: restoredContent.after },
          trailing,
        ),
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
  if (
    cutHasResolutionProvenance({ content: joined.content, gaps: [joinedGap] })
  ) {
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
      ...reviewSetting(
        op.story,
        survivor.paraId ?? "",
        splitReview,
        reviewFieldsOf(survivor),
      ),
    ],
    touched: committed.value.touched,
  });
};

const replaceBlocks = (document: Document, op: ReplaceBlocksOp): Applied => {
  const [firstExpected] = op.expected;
  if (firstExpected === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "A replacement names the paragraphs it replaces and the ones it puts in their place.",
    );
  }
  const body = storyBody(document, op.story);
  const found: ParagraphLocation[] = [];
  for (const expected of op.expected) {
    if (expected.paraId === undefined) {
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        "A replaced paragraph has no id.",
      );
    }
    // Empty ids also match an unidentified paragraph in the legacy lookup.
    const located =
      expected.paraId === ""
        ? findParagraph(op, storyParagraphs(body), expected.paraId)
        : locate(document, op, op.story, expected.paraId);
    if (located.isErr()) {
      return Result.err(located.error);
    }
    found.push(located.value);
  }
  const [start] = found;
  if (start === undefined) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "Nothing to replace.",
    );
  }
  const contiguous = found.every(
    (location, offset) =>
      sameBlockList(location.list, start.list) &&
      location.index === start.index + offset,
  );
  if (!contiguous) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      "The replaced paragraphs are not adjacent.",
    );
  }
  const stale = found.some(
    (location, offset) =>
      !equalForStaleness(location.paragraph, op.expected[offset]),
  );
  if (stale) {
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The paragraphs to replace have changed.",
    );
  }
  const before = found.map(({ paragraph }) => paragraph);
  const valid = validateBlockReplacement({ document, op, before });
  if (valid.isErr()) return valid;
  let inverseAt: TableInsertionPoint | undefined;
  if (op.blocks.length === 0) {
    // The insertion inverse must be able to place every captured source mark.
    if (
      found.some((location) => hasContainerFinalParagraphMark(body, location))
    )
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        "An empty replacement cannot restore a captured container-final paragraph mark through insertion.",
      );
    // The inverse uses ordinary insertion, whose canonicalization removes empty
    // formatting fields. Refuse captures whose field ownership it cannot restore.
    if (
      before.some(
        (paragraph) =>
          Object.hasOwn(paragraph, "formatting") &&
          finalizeParagraphFormatting(paragraph.formatting) === undefined,
      )
    )
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        "An empty replacement cannot restore an owned empty paragraph-formatting field through insertion.",
      );
    const originalList = blockListAt(body.content, start.list);
    const removalList = [...originalList];
    removalList.splice(start.index, before.length);
    inverseAt = blockInsertionPointAt({
      blocks: removalList,
      index: start.index,
    });
    if (before.some(({ sectionProperties }) => sectionProperties !== undefined))
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
        "An empty replacement cannot remove a section endpoint.",
      );
    if (inverseAt === undefined)
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        "An empty replacement needs a surviving paragraph or table boundary.",
      );
  }
  const replaced = replaceParagraphs({
    document,
    story: op.story,
    at: start,
    count: before.length,
    replacement: op.blocks,
    ...(op.sectionView === undefined
      ? {}
      : { restoreSections: op.sectionView.restore }),
  });
  if (replaced.isErr())
    return Result.err(
      refusal(op, replaced.error.reason, replaced.error.message),
    );
  if (op.blocks.length === 0) {
    const cellStep = start.list.findLastIndex(
      (step) => step.kind === "tableCell",
    );
    if (
      cellStep >= 0 &&
      !blockListEndsInParagraph(
        blockListAt(
          storyBody(replaced.value, op.story).content,
          start.list.slice(0, cellStep + 1),
        ),
      )
    )
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        "An empty replacement must leave its cell ending with a paragraph.",
      );
    if (inverseAt === undefined)
      return panic(
        "A validated empty replacement lost its insertion boundary.",
      );
    return Result.ok({
      document: replaced.value,
      inverse: [
        {
          type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          story: op.story,
          at: inverseAt,
          blocks: before,
        },
      ],
      touched: touchedBetween(before, []),
    });
  }
  return Result.ok({
    document: replaced.value,
    ...nonemptyBlockReplacementEffects(op, before),
  });
};

const applyParagraphPrimitive = (
  document: Document,
  op: ResolutionParagraphOp,
): Applied => {
  const located = locate(document, op, op.story, op.blockId);
  if (located.isErr()) return Result.err(located.error);
  const prepared = prepareResolutionParagraphOp({
    document,
    op,
    at: located.value,
  });
  if (prepared.isErr()) return Result.err(prepared.error);
  if (prepared.value.type === "unchanged") return unchanged(document);
  const { paragraph, inverse, touched } = prepared.value;
  const replaced = replaceParagraphs({
    document,
    story: op.story,
    at: located.value,
    count: 1,
    replacement: [paragraph],
  });
  if (replaced.isErr())
    return Result.err(
      refusal(op, replaced.error.reason, replaced.error.message),
    );
  return Result.ok({ document: replaced.value, inverse: [inverse], touched });
};

/**
 * A note list with a new note before the first note of a greater id: notes
 * added concurrently take the same places whichever is added first.
 */
const withNoteById = <Note extends { id: number }>(
  notes: readonly Note[] | undefined,
  note: Note,
): Note[] => {
  const list = [...(notes ?? [])];
  const at = list.findIndex(({ id }) => id > note.id);
  list.splice(at < 0 ? list.length : at, 0, note);
  return list;
};

type EditNoteOptions = {
  document: Document;
  op: AddNoteOp | RemoveNoteOp;
  applyOps: ApplyOps;
};
const editNote = ({ document, op, applyOps }: EditNoteOptions): Applied => {
  const kind =
    op.type === DOCUMENT_OP_TYPES.ADD_NOTE ? op.note.type : op.story.kind;
  const id = op.type === DOCUMENT_OP_TYPES.ADD_NOTE ? op.note.id : op.story.id;
  if (!Number.isInteger(id) || id <= 0)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      "A normal note needs a positive integer id.",
    );
  const collection =
    kind === "footnote"
      ? document.package.footnotes
      : document.package.endnotes;
  if (op.type === DOCUMENT_OP_TYPES.ADD_NOTE) {
    if (collection?.some((note) => note.id === id))
      return refuse(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        "The note id is already used.",
      );
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
      markers.some(
        ({ node }) => node.type === "noteMarker" && node.kind !== kind,
      )
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
              footnotes: withNoteById(document.package.footnotes, note),
            },
          }
        : {
            ...document,
            package: {
              ...document.package,
              endnotes: withNoteById(document.package.endnotes, note),
            },
          };
    const valid = validateOpsDocument(withNote);
    if (valid.isErr())
      return refuse(op, valid.error.reason, valid.error.message);
    const reference = insertContent(withNote, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: op.at,
      slice: {
        content: [
          {
            type: "run",
            content: [
              { type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id },
            ],
          },
        ],
        openStart: 0,
        openEnd: 0,
      },
      ...(op.revision === undefined ? {} : { revision: op.revision }),
      ...(op.newIds === undefined ? {} : { newIds: op.newIds }),
    });
    if (reference.isErr()) return reference;
    return storyLifecycleEdit(document, reference.value.document, op);
  }
  const note = collection?.find((candidate) => candidate.id === id);
  if (!note)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The note does not exist.",
    );
  const located = locateRange(document, op, op.at, {
    ...op.at,
    offset: op.at.offset + 1,
  });
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
    ...(op.revision === undefined ? {} : { revision: op.revision }),
    ...(op.newIds === undefined ? {} : { newIds: op.newIds }),
  });
  if (removed.isErr()) return removed;
  const stillReferenced = documentStories(removed.value.document).some(
    (story) => {
      if (
        typeof story !== "string" &&
        story.kind === kind &&
        "id" in story &&
        story.id === id
      )
        return false;
      return storyParagraphs(storyBody(removed.value.document, story)).some(
        ({ paragraph }) =>
          leafSpans(paragraph.content).some(
            ({ node, ancestors }) =>
              ((kind === "footnote" && node.type === "footnoteRef") ||
                (kind === "endnote" && node.type === "endnoteRef")) &&
              node.id === id &&
              !ancestors.some(
                (ancestor) =>
                  ancestor.type === "deletion" &&
                  ancestor.info.id === op.revision?.id,
              ),
          ),
      );
    },
  );
  if (stillReferenced)
    return refuse(
      op,
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Another reference still owns this note.",
    );
  if (op.revision !== undefined)
    return storyLifecycleEdit(document, removed.value.document, op);
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
              removed.value.document.package.endnotes?.filter(
                (candidate) => candidate.id !== id,
              ) ?? [],
          },
        };
  const cleanup = replaceNoteParts({
    document: removed.value.document,
    parts: next.package,
    op,
    applyOps,
  });
  return cleanup.map((edit) => combineEdits(document, [removed.value, edit]));
};

/** Direct text edits own the lifecycle of note bodies whose last reference they remove. */
const withoutOrphanedNotes = (before: Document, after: Document): Document => {
  const removed = new Set(
    removedNoteStories(before, after).map(
      ({ kind, id }) => `${kind}:${String(id)}`,
    ),
  );
  if (removed.size === 0) return after;
  return {
    ...after,
    package: {
      ...after.package,
      ...(after.package.footnotes === undefined
        ? {}
        : {
            footnotes: after.package.footnotes.filter(
              ({ id }) => !removed.has(`footnote:${String(id)}`),
            ),
          }),
      ...(after.package.endnotes === undefined
        ? {}
        : {
            endnotes: after.package.endnotes.filter(
              ({ id }) => !removed.has(`endnote:${String(id)}`),
            ),
          }),
    },
  };
};

type OperationExecutionPurpose = "apply" | "inspect-inverse";
type DispatchOptions = {
  document: Document;
  op: DocumentOp;
  purpose: OperationExecutionPurpose;
};
const dispatchCore = ({ document, op, purpose }: DispatchOptions): Applied => {
  const applyOps = (current: Document, ops: readonly DocumentOp[]) =>
    applyDocumentOpsForPurpose({ document: current, ops, purpose });
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
      return applyCommentOp(document, op, applyOps);
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
      return applyStoryLifecycle(document, op);
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
      return editNote({ document, op, applyOps });
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
      return deleteBlocks({ document, op, applyOps });
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return insertBlocks({ document, op, applyOps, validateBlockReplacement });
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
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
      return applyParagraphPrimitive(document, op);
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
      return resolveRevision(document, op, applyOps);
    case DOCUMENT_OP_TYPES.SET_REVISION_INFO:
      return setRevisionInfo(document, op);
    case DOCUMENT_OP_TYPES.DELETE_TABLE: {
      // The author's own pending column goes outright before the rest of
      // the table is suggested for deletion.
      const pending = ownPendingTableInsertion(document, op);
      if (pending === undefined) return applyTableOp(document, op);
      const { expected: _checked, ...rest } = op;
      return applyOps(document, [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: op.story,
          revisionIds: pending.revisionIds,
          decision: REVISION_DECISIONS.REJECT,
        },
        { ...rest, blockId: pending.blockId },
      ]).map(({ document: edited, inverse, touched }) => ({
        document: edited,
        inverse,
        touched,
      }));
    }
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
      return applyTableOp(document, op);
    case DOCUMENT_OP_TYPES.INSERT_ROW:
    case DOCUMENT_OP_TYPES.DELETE_ROW:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
      return applyRowOp(document, op);
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
      return applyPackageResourcesOp(document, op);
    case DOCUMENT_OP_TYPES.SET_CONTENT_CONTROL:
      return setContentControl(document, op);
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
    case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE:
      return applyTableEdit(document, op);
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

/**
 * A direct operation owns the note records made unreachable by its text edit.
 * Keeping this in dispatch makes the lifecycle edit part of the operation's
 * result, so its inverse and touched-story contract include the removed body.
 */
const dispatch = ({ document, op, purpose }: DispatchOptions): Applied => {
  const applied = dispatchCore({ document, op, purpose });
  if (applied.isErr()) return applied;
  if (
    op.type === DOCUMENT_OP_TYPES.ADD_NOTE ||
    op.type === DOCUMENT_OP_TYPES.REMOVE_NOTE ||
    op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS
  )
    return applied;
  const cleaned =
    "revision" in op && op.revision !== undefined
      ? applied.value.document
      : withoutOrphanedNotes(document, applied.value.document);
  let edit = applied.value;
  if (cleaned !== edit.document) {
    const cleanup = replaceNoteParts({
      document: edit.document,
      parts: cleaned.package,
      op,
      applyOps: (current, ops) =>
        applyDocumentOpsForPurpose({ document: current, ops, purpose }),
    });
    if (cleanup.isErr()) return cleanup;
    edit = combineEdits(document, [edit, cleanup.value]);
  }
  // Inspection applies ordinary validation and cleanup, but never inspects
  // the inverse it just generated (reference swaps would otherwise recurse).
  if (purpose === "inspect-inverse") return Result.ok(edit);
  const inverseCleanup = withoutOrphanedNotes(applied.value.document, document);
  if (inverseCleanup === document) return Result.ok(edit);

  // Adding the first reference to an authored orphan note must not let the
  // ordinary inverse delete its pre-existing declaration. Replay only this
  // uncommon case, retaining every inverse owned by the primary edit.
  const undone = applyDocumentOpsForPurpose({
    document: edit.document,
    ops: edit.inverse,
    purpose: "inspect-inverse",
  });
  if (undone.isErr())
    return panic(
      "The primary edit's inverse must apply before note repair.",
      undone.error,
    );
  const owned = {
    footnotes: null,
    endnotes: null,
    body: { comments: null },
  } as const satisfies StoryParts;
  const original = captureStoryParts(document, owned);
  const relationships = [...(document.package.relationships ?? [])];
  const matchesSource = (candidate: Document) =>
    structurallyEqual(captureStoryParts(candidate, owned), original) &&
    Object.hasOwn(candidate.package, "relationships") ===
      Object.hasOwn(document.package, "relationships") &&
    (candidate.package.relationships === undefined) ===
      (document.package.relationships === undefined) &&
    structurallyEqual(
      [...(candidate.package.relationships ?? [])],
      relationships,
    );
  if (matchesSource(undone.value.document)) return Result.ok(edit);
  const pruning = replaceNoteParts({
    document,
    parts: inverseCleanup.package,
    op,
    applyOps: (current, ops) =>
      applyDocumentOpsForPurpose({
        document: current,
        ops,
        purpose: "inspect-inverse",
      }),
  });
  if (pruning.isErr())
    return panic(
      "Source note pruning must apply before inverse repair.",
      pruning.error,
    );
  for (const key of ["footnotes", "endnotes"] as const) {
    const field = { [key]: null };
    if (
      !structurallyEqual(
        captureStoryParts(undone.value.document, field),
        captureStoryParts(pruning.value.document, field),
      )
    )
      return panic("Note repair may only restore inverse-pruned declarations.");
  }
  const repaired = applyDocumentOpsForPurpose({
    document: undone.value.document,
    ops: pruning.value.inverse,
    purpose: "inspect-inverse",
  });
  if (repaired.isErr())
    return panic(
      "The source note repair must meet the primary inverse.",
      repaired.error,
    );
  if (!matchesSource(repaired.value.document))
    return panic(
      "Note repair must restore its source declarations and comment state.",
    );
  return Result.ok({
    ...edit,
    inverse: [...edit.inverse, ...pruning.value.inverse],
  });
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
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
    case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
      return op.revision;
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
    case DOCUMENT_OP_TYPES.SET_TABLE:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
    case DOCUMENT_OP_TYPES.SET_REVISION_INFO:
    case DOCUMENT_OP_TYPES.SET_CONTENT_CONTROL:
      return undefined;
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

/**
 * The blocks an edit replaced and the ones it left in their place. Shared
 * prefix and suffix blocks can carry only identities already present.
 */
const changedRevisionBlocks = (
  previous: readonly BlockContent[] | undefined,
  current: BlockContent[],
): { before: readonly BlockContent[]; after: BlockContent[] } => {
  if (previous === current) return { before: [], after: [] };
  if (previous === undefined) return { before: [], after: current };
  const { start, previousEnd, end } = changedListBounds(previous, current);
  return {
    before: previous.slice(start, previousEnd),
    after: current.slice(start, end),
  };
};

/** The revision ids a tracked operation's records carry. */
export type RecordedRevisionIds = Pick<
  AppliedDocumentOp,
  "revisions" | "folded"
>;

const NO_RECORDED_REVISIONS: RecordedRevisionIds = Object.freeze({
  revisions: [],
  folded: [],
});

/**
 * The revision ids of operations applied in turn, as one operation's: every
 * step's new ids, and the existing ids a step folded into that no step of
 * them allocated.
 */
export const combineRecordedRevisions = (
  steps: readonly RecordedRevisionIds[],
): RecordedRevisionIds => {
  const revisions = steps.flatMap((step) => step.revisions);
  const allocated = new Set(revisions);
  const folded = new Set(
    steps.flatMap((step) => step.folded).filter((id) => !allocated.has(id)),
  );
  return { revisions, folded: [...folded] };
};

/**
 * Newly allocated revision ids, and the existing ones whose pending property
 * changes the operation folded its edit into.
 */
const recordedRevisions = (
  before: Document,
  edit: DocumentEdit,
  stamp: RevisionStamp,
  boxes: readonly TextBoxStory[],
): RecordedRevisionIds => {
  const touched = new Set(
    [...edit.touched.modified, ...edit.touched.inserted].map(idKey),
  );
  const known = packageIdentityKeyIndex(before.package);
  const tableRevisions: number[] = [];
  const paragraphRevisions: number[] = [];
  const folded: number[] = [];
  for (const story of [...documentStories(edit.document), ...boxes]) {
    const blocks = changedRevisionBlocks(
      findStoryBody(before, story)?.content,
      storyBody(edit.document, story).content,
    );
    if (blocks.after.length === 0) continue;
    tableRevisions.push(
      ...stampedTableRevisionIds({ blocks: blocks.after, stamp, known }),
    );
    const replaced: Paragraph[] = [];
    for (const { paragraph } of storyParagraphs({ content: blocks.after })) {
      const { paraId } = paragraph;
      if (paraId === undefined || !touched.has(idKey(paraId))) continue;
      replaced.push(paragraph);
      paragraphRevisions.push(...stampedRevisionIds([paragraph], stamp, known));
    }
    folded.push(
      ...foldedRevisionIds({
        before: storyParagraphs({ content: blocks.before }).map(
          ({ paragraph }) => paragraph,
        ),
        after: replaced,
        stamp,
        known,
      }),
    );
  }
  // Every physical row record is reported before paragraph records.
  return {
    revisions: tableRevisions.concat(paragraphRevisions),
    folded: [...new Set(folded)],
  };
};

/** Content edits retain authored snapshots across immutable paragraph replacements. */
const preserveParagraphPropertySnapshots = (options: {
  before: Document;
  edit: DocumentEdit;
  boxes: readonly TextBoxStory[];
}): void => {
  for (const { source, target } of paragraphSnapshotTransfers(options))
    transferParagraphPropertySnapshot({ source, target });
};

type ApplyDocumentOpOptions = {
  document: Document;
  input: DocumentOp;
  purpose: OperationExecutionPurpose;
};
const applyDocumentOpForPurpose = ({
  document,
  input,
  purpose,
}: ApplyDocumentOpOptions): Result<AppliedDocumentOp, DocumentOpRefusal> => {
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
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.STALE,
          "The derived section metadata changed.",
        ),
      );
  }
  const malformedTable = tableEditBoundaryRefusal(op);
  if (malformedTable !== undefined) return Result.err(malformedTable);
  const addressed =
    op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT
      ? undefined
      : addressedStory(op);
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
  const story = addressedStory(op);
  const routes = textBoxRoute(document, op, story);
  if (routes === undefined) {
    // An operation compiled in a box's own story names it already.
    const box = story === undefined ? undefined : textBoxStoryOf(story);
    return applyAddressedOp({
      document,
      op,
      purpose,
      boxes: box === undefined ? [] : [box],
    });
  }
  let current = document;
  const edits: DocumentEdit[] = [];
  const recorded: RecordedRevisionIds[] = [];
  for (const route of routes) {
    const applied = applyAddressedOp({
      document: current,
      op: route.op,
      purpose,
      boxes: route.box === undefined ? [] : [route.box],
    });
    if (applied.isErr()) return applied;
    // The paragraphs drawing the box hold its body: they change with it.
    const hosts =
      route.box === undefined ||
      findStoryBody(current, route.box) ===
        findStoryBody(applied.value.document, route.box)
        ? []
        : textBoxHostIds(route.box);
    current = applied.value.document;
    const { touched } = applied.value;
    edits.push({
      ...applied.value,
      touched: {
        ...touched,
        modified: [
          ...touched.modified,
          ...hosts.filter((id) => !touched.modified.includes(id)),
        ],
      },
    });
    recorded.push(applied.value);
  }
  const combined = combineEdits(document, edits);
  return Result.ok({
    ...combined,
    inverse: combined.inverse.map(namingPackageStories),
    ...combineRecordedRevisions(recorded),
  });
};

/** The story an operation's text or blocks are in, for those that name one. */
const addressedStory = (op: DocumentOp): OpStory | undefined => {
  if (op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT)
    switch (op.anchor.kind) {
      case "range":
        return op.anchor.from.story;
      case "point":
        return op.anchor.at.story;
      case "revision":
        return op.anchor.story;
      case "reply":
        return undefined;
      default: {
        const unreachable: never = op.anchor;
        return unreachable;
      }
    }
  if ("at" in op && typeof op.at === "object" && "story" in op.at)
    return op.at.story;
  if ("from" in op) return op.from.story;
  if ("story" in op) return op.story;
  return undefined;
};

type ApplyAddressedOpOptions = {
  document: Document;
  op: DocumentOp;
  purpose: OperationExecutionPurpose;
  /** The text box story the operation was routed to, if any. */
  boxes: readonly TextBoxStory[];
};
const applyAddressedOp = ({
  document,
  op,
  purpose,
  boxes,
}: ApplyAddressedOpOptions): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  const stamp = stampOf(op);
  const badStamp =
    stamp === undefined ? undefined : stampRefusal(document, op, stamp);
  if (badStamp !== undefined) {
    return Result.err(badStamp);
  }
  const applied = dispatch({ document, op, purpose });
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
          Object.hasOwn(section, "headers") ===
            Object.hasOwn(next, "headers") &&
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
  if (hasMultipleParagraphPropertyChanges(edit.document, boxes)) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "A paragraph cannot carry more than one property change.",
      ),
    );
  }
  // The contract keeps paragraph ids unique: an edit that would repeat one,
  // however its operation came by it, is refused here, where it is owned.
  const repeatedId = repeatedParagraphId(edit.document);
  if (repeatedId !== undefined)
    return Result.err(refusal(op, repeatedId.reason, repeatedId.message));
  meetsContract(edit.document);
  preserveParagraphPropertySnapshots({ before: document, edit, boxes });
  // A direct join removes one of its paragraphs into the other.
  if (op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS) {
    const gone = edit.touched.removed.find(
      (id) =>
        idKey(id) === idKey(op.blockId) || idKey(id) === idKey(op.nextBlockId),
    );
    if (gone !== undefined)
      edit.touched = {
        ...edit.touched,
        absorbed: [
          ...(edit.touched.absorbed ?? []),
          {
            absorbed: gone,
            survivor:
              idKey(gone) === idKey(op.blockId) ? op.nextBlockId : op.blockId,
          },
        ],
      };
  }
  return Result.ok({
    ...edit,
    inverse: edit.inverse.map(captureDocumentOp),
    ...(stamp === undefined
      ? NO_RECORDED_REVISIONS
      : recordedRevisions(document, edit, stamp, boxes)),
  });
};

/** Apply one operation to a document that meets the seed contract. */
export const applyDocumentOp = (
  document: Document,
  input: DocumentOp,
): Result<AppliedDocumentOp, DocumentOpRefusal> =>
  applyDocumentOpForPurpose({ document, input, purpose: "apply" });

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
 * An inverse whose whole-part restores expect the parts they meet when the
 * inverse runs, in order, on `document`.
 *
 * A restore records the parts its operation left, and refuses as stale when
 * the document it meets holds other ones. A caller that changes the document
 * after an operation (a derivation over the whole document, such as list
 * rendering) and puts its own inverse first rebases the restores onto that
 * document, so undo meets exactly what each restore expects. The inverse is
 * returned as it is when it holds no restore, or when it does not apply.
 */
export const inverseMeetingDocument = (
  document: Document,
  inverse: readonly DocumentOp[],
): readonly DocumentOp[] => {
  if (!inverse.some((op) => op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS))
    return inverse;
  let current = document;
  const rebased: DocumentOp[] = [];
  for (const op of inverse) {
    const next =
      op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS
        ? {
            ...op,
            expected: cloneModel(captureStoryParts(current, op.expected)),
          }
        : op;
    const applied = applyDocumentOp(current, next);
    if (applied.isErr()) return inverse;
    current = applied.value.document;
    rebased.push(next);
  }
  return rebased;
};

type ApplyDocumentOpsOptions = {
  document: Document;
  ops: readonly DocumentOp[];
  purpose: OperationExecutionPurpose;
};
const applyDocumentOpsForPurpose = ({
  document,
  ops,
  purpose,
}: ApplyDocumentOpsOptions): Result<AppliedDocumentOp, DocumentOpRefusal> => {
  let current = document;
  const effects: (DocumentEditEffects & { document?: never })[] = [];
  const recorded: RecordedRevisionIds[] = [];
  for (const op of ops) {
    const applied = applyDocumentOpForPurpose({
      document: current,
      input: op,
      purpose,
    });
    if (applied.isErr()) {
      return applied;
    }
    const { document: next, inverse, touched } = applied.value;
    current = next;
    effects.push({ inverse, touched });
    recorded.push(applied.value);
  }
  return Result.ok({
    ...combineEditEffects(current, effects),
    ...combineRecordedRevisions(recorded),
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
): Result<AppliedDocumentOp, DocumentOpRefusal> =>
  applyDocumentOpsForPurpose({ document, ops, purpose: "apply" });
