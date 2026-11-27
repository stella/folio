/** Paragraph insertion, isolated from the inline and table operation implementations. */
import { panic, Result } from "better-result";

import {
  type BlockContent,
  type Document,
  finalizeParagraphFormatting,
  type Paragraph,
  type ParagraphMarkChange,
  type ParagraphMarkPropertyChange,
  type ParagraphPropertyChange,
  type TrackedChangeInfo,
} from "../model/document";
import {
  isOpaqueFollower,
  markBeforeTableResolves,
  markGoes,
  paragraphTableRunInBoundary,
} from "./blockDeletion";
import {
  blockListAt,
  blockListEndsInParagraph,
  canSpliceMainBodyWithoutSectionRepair,
  endsItsContainer,
  hasContainerFinalParagraphMark,
  replaceParagraphs,
  storyBody,
} from "./blocks";
import { validateOpsDocument } from "./contract";
import type { DocumentEdit, DocumentEditEffects } from "./edits";
import { structurallyEqual } from "./equality";
import { firstSurvivorFields } from "./firstSurvivor";
import { freshenIdentities } from "./identity";
import {
  type IdentityKeyIndex,
  idKey,
  isParaId,
  type KeyMembership,
  packageIdentityKeyIndex,
  slotKey,
} from "./ids";
import { asMoveHalf } from "./moves";
import { paragraphLength } from "./offsets";
import { findStoryParagraph } from "./paragraphLocator";
import {
  nonemptyBlockReplacementEffects,
  validateBlockReplacement as validateReplacement,
} from "./paragraphPrimitive";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  pendingParagraphMarkReason,
} from "./refusal";
import type { ApplyOps } from "./resolve";
import {
  foldedParagraphPropertyChange,
  paragraphMarkPropertyChange,
  paragraphPropertyFoldConflict,
  sameMarkFormatting,
  sameParagraphProperties,
  stampInfo,
  WRAP_KINDS,
  wrapTracked,
} from "./review";
import { identitySlots } from "./slots";
import { locateBlockInsertion } from "./tableLocation";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertBlocksOp,
  OP_STORIES,
  type ReplaceBlocksOp,
  type RevisionStamp,
  SECTION_BOUNDARY_POLICIES,
} from "./types";

type InsertedContentOptions = {
  paragraph: Paragraph;
  revision: RevisionStamp;
  move: InsertBlocksOp["move"];
};
const insertedContent = ({
  paragraph,
  revision,
  move,
}: InsertedContentOptions): Result<Paragraph["content"], DocumentOpRefusal> => {
  const wrapped = wrapTracked({
    items: paragraph.content,
    from: { offset: 0, zeroWidthBefore: 0 },
    to: {
      offset: paragraphLength(paragraph),
      zeroWidthBefore: Number.MAX_SAFE_INTEGER,
    },
    kind: WRAP_KINDS.INSERTION,
    stamp: revision,
  });
  if (wrapped.kind === "refused")
    return Result.err(
      new DocumentOpRefusal({
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        reason: wrapped.reason,
        message: "The inserted content cannot be tracked.",
      }),
    );
  const content =
    wrapped.kind === "wrapped" ? wrapped.content : paragraph.content;
  return Result.ok(
    move === true ? asMoveHalf(content, revision.id, "to") : content,
  );
};

type BlockInsertionOptions = {
  document: Document;
  op: InsertBlocksOp;
  applyOps: ApplyOps;
  validateBlockReplacement: ValidateBlockReplacement;
};

export type ValidateBlockReplacement = (options: {
  document: Document;
  op: ReplaceBlocksOp;
  before: readonly Paragraph[];
}) => Result<void, DocumentOpRefusal>;

/** A paragraph's fields a join on rejection takes from the paragraph after. */
const joinedFields = ({
  paraId: _paraId,
  textId: _textId,
  preservedAttributes: _preservedAttributes,
  renderedPageBreakBefore: _renderedPageBreakBefore,
  content: _content,
  pPrMark: _pPrMark,
  ...fields
}: Paragraph): Omit<
  Paragraph,
  | "paraId"
  | "textId"
  | "preservedAttributes"
  | "renderedPageBreakBefore"
  | "content"
  | "pPrMark"
> => fields;

type MovedPendingBreakOptions = {
  /** The paragraph whose inserted mark keeps it on rejection. */
  anchor: Paragraph;
  /** The blocks after the anchor in its block list. */
  after: readonly BlockContent[];
  blocks: readonly Paragraph[];
};

/** Whether two records were made by one change: alike in all but their ids. */
const sameChange = (
  left: TrackedChangeInfo,
  right: TrackedChangeInfo,
): boolean => structurallyEqual({ ...left, id: 0 }, { ...right, id: 0 });

/**
 * Whether `mark` is a plain inserted break of the change `info` records that
 * keeps the paragraph before it.
 */
const keepsFirstBreakOf = (
  mark: ParagraphMarkChange | undefined,
  info: TrackedChangeInfo,
): boolean =>
  mark?.kind === "ins" &&
  sameChange(mark.info, info) &&
  mark.survivor === "first" &&
  mark.takenOverMark === undefined &&
  mark.deletion === undefined &&
  (mark.resolutionJoin ?? 0) === 0;

/**
 * The anchor and the new paragraphs of a direct insertion after the anchor's
 * pending break, the break moved to the last new paragraph, as a break typed
 * at the anchor's end leaves it; `undefined` where rejecting the moved break
 * would not give the insertion into the rejected document. Rejecting joins
 * the last new paragraph through the paragraphs that insertion added into
 * the one holding the anchor's former mark, keeping the new paragraph and
 * taking that holder's properties. So the paragraphs after the anchor must
 * hold only that insertion's content, the holder no record of its own, and
 * the anchor and the last new paragraph its properties. Any other difference
 * is one only a property change of the pending insertion records, and a
 * direct insertion makes no tracked change.
 */
const movedPendingBreak = ({
  anchor,
  after,
  blocks,
}: MovedPendingBreakOptions): Paragraph[] | undefined => {
  const mark = anchor.pPrMark;
  const last = blocks.at(-1);
  if (
    mark === undefined ||
    !keepsFirstBreakOf(mark, mark.info) ||
    anchor.sectionProperties !== undefined ||
    last === undefined ||
    last.pPrMark !== undefined
  )
    return undefined;
  let holder: Paragraph | undefined;
  for (const block of after) {
    if (
      block.type !== "paragraph" ||
      !block.content.every(
        (node) => node.type === "insertion" && sameChange(node.info, mark.info),
      )
    )
      return undefined;
    if (block.pPrMark === undefined) {
      holder = block;
      break;
    }
    if (!keepsFirstBreakOf(block.pPrMark, mark.info)) return undefined;
  }
  if (holder === undefined) return undefined;
  const fields = joinedFields(holder);
  if (
    !structurallyEqual(joinedFields(anchor), fields) ||
    !structurallyEqual(joinedFields(last), fields)
  )
    return undefined;
  const { pPrMark: _moved, ...own } = anchor;
  return [own, ...blocks.slice(0, -1), { ...last, pPrMark: mark }];
};

const canonicalizeInsertedParagraph = (paragraph: Paragraph): Paragraph => {
  if (finalizeParagraphFormatting(paragraph.formatting) !== undefined)
    return paragraph;
  const canonical = { ...paragraph };
  delete canonical.formatting;
  return canonical;
};

const paragraphReplacement =
  ({
    op,
    before,
    anchor,
  }: {
    op: InsertBlocksOp;
    before: boolean;
    anchor: Paragraph;
  }) =>
  (blocks: readonly Paragraph[], replacedAnchor = anchor): ReplaceBlocksOp => ({
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: op.story,
    expected: [anchor],
    blocks: before ? [...blocks, replacedAnchor] : [replacedAnchor, ...blocks],
  });

type TakenOverPropertyChangesOptions = {
  /** The anchor whose mark the last new paragraph takes over. */
  source: Paragraph;
  /** The last new paragraph's properties. */
  current: Paragraph["formatting"];
  revision: RevisionStamp;
};

/**
 * The property changes the paragraph that takes over `source`'s mark holds:
 * `source`'s own pending one, with other properties folded into it (keeping
 * its id and original properties, as a split's second half does), or a new
 * one from `source`'s properties. A suggestion does not fold into another's
 * change.
 */
const takenOverPropertyChanges = ({
  source,
  current,
  revision,
}: TakenOverPropertyChangesOptions): Result<
  ParagraphPropertyChange[] | undefined,
  DocumentOpRefusal
> => {
  const pending =
    (source.propertyChanges?.length ?? 0) > 0
      ? source.propertyChanges
      : undefined;
  if (
    sameParagraphProperties(
      finalizeParagraphFormatting(source.formatting),
      current,
    )
  )
    return Result.ok(pending);
  if (pending === undefined)
    return Result.ok([
      {
        type: "paragraphPropertyChange",
        info: stampInfo(revision),
        ...(source.formatting === undefined
          ? {}
          : { previousFormatting: source.formatting }),
      },
    ]);
  const conflict = paragraphPropertyFoldConflict(source, revision);
  if (conflict !== undefined)
    return Result.err(
      new DocumentOpRefusal({
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
        message: conflict,
      }),
    );
  return Result.ok([
    foldedParagraphPropertyChange({
      paragraph: source,
      formatting: current,
      stamp: revision,
    }),
  ]);
};

/**
 * The run-property change of the mark the last new paragraph takes over:
 * `source`'s own while the mark keeps its run properties, else a new one
 * from them. A pending one is resolved first, as a split's is.
 */
const takenOverMarkChange = ({
  source,
  current,
  revision,
}: TakenOverPropertyChangesOptions): Result<
  ParagraphMarkPropertyChange | undefined,
  DocumentOpRefusal
> => {
  if (sameMarkFormatting(source.formatting, current))
    return Result.ok(source.markPropertyChange);
  if (source.markPropertyChange !== undefined)
    return Result.err(
      new DocumentOpRefusal({
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_RUN_PROPERTY_CHANGE,
        message:
          "Resolve the paragraph mark's existing run-property change first.",
      }),
    );
  return Result.ok(paragraphMarkPropertyChange(stampInfo(revision), source));
};

type TakenOverMarksOptions = {
  /** The mark of the anchor, when the last new paragraph takes it over. */
  mark: ParagraphMarkChange | undefined;
  /** The anchor's new mark, which this insertion adds. */
  added: ParagraphMarkChange;
};

type TakenOverMarks = {
  anchor: ParagraphMarkChange;
  /** The mark of the paragraph that takes over the anchor's. */
  taker: ParagraphMarkChange | undefined;
};

/**
 * The marks of the anchor and of the paragraph that takes over its former
 * one: an added mark goes with it. A break pending deletion or moving away
 * is refused (`PENDING_BREAK_TAKEOVER`), the stated boundary of taking over
 * a mark: no placement of its record keeps both laws. On the new paragraph,
 * accepting would run that paragraph on instead of the anchor's words, so
 * accepting differs from the direct edit; beside the anchor's new mark,
 * rejecting the insertion alone would take the deletion with that mark, so
 * the original does not come back.
 */
const takenOverMarks = ({
  mark,
  added,
}: TakenOverMarksOptions): Result<TakenOverMarks, DocumentOpRefusal> => {
  if (mark === undefined) return Result.ok({ anchor: added, taker: undefined });
  if (
    (mark.kind === "ins" || mark.kind === "moveTo") &&
    mark.deletion === undefined
  )
    return Result.ok({ anchor: added, taker: mark });
  return Result.err(
    new DocumentOpRefusal({
      opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_BREAK_TAKEOVER,
      message:
        "A tracked insertion cannot take over the mark of a paragraph whose break is pending deletion or moving away.",
    }),
  );
};

type TrackedParagraphReplacementOptions = {
  op: InsertBlocksOp;
  revision: RevisionStamp;
  source: Paragraph;
  anchor: Paragraph;
  takesOverMark: boolean;
  canonicalBlocks: readonly Paragraph[];
  replacement: ReturnType<typeof paragraphReplacement>;
  usedElsewhere: () => KeyMembership;
};

const trackedParagraphReplacement = ({
  op,
  revision,
  source,
  anchor,
  takesOverMark,
  canonicalBlocks,
  replacement,
  usedElsewhere,
}: TrackedParagraphReplacementOptions): Result<
  ReplaceBlocksOp,
  DocumentOpRefusal
> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  const taker = canonicalBlocks.at(-1);
  if (
    takesOverMark &&
    ((taker?.propertyChanges?.length ?? 0) > 0 ||
      taker?.markPropertyChange !== undefined)
  )
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
      "The pasted paragraph would give its break two property histories.",
    );
  // A move's destination marks its paragraphs moved to, not inserted.
  const markKind = op.move === true ? "moveTo" : "ins";
  const former = takesOverMark ? source.pPrMark : undefined;
  const survivor = firstSurvivorFields(former);
  const marks = takenOverMarks({
    mark: former,
    added: { kind: markKind, info: stampInfo(revision), ...survivor },
  });
  if (marks.isErr()) return marks;
  const blocks: Paragraph[] = [];
  for (const [index, paragraph] of canonicalBlocks.entries()) {
    const wrapped = insertedContent({
      paragraph,
      revision,
      move: op.move,
    });
    if (wrapped.isErr()) return wrapped;
    const content = wrapped.value;
    if (!takesOverMark || index < canonicalBlocks.length - 1) {
      blocks.push({
        ...paragraph,
        content,
        pPrMark: {
          kind: markKind,
          info: stampInfo(revision),
          // After a final paragraph each new break is rejected into the
          // paragraph before it, which ends up being the anchor, as the
          // anchor's own new break is.
          ...(takesOverMark ? survivor : {}),
        },
      });
      continue;
    }
    // The anchor's former mark, with the review records it holds, as a
    // split's second half takes them: its properties before are the
    // anchor's, and other properties fold into a pending change of them.
    const taken = {
      source,
      current: finalizeParagraphFormatting(paragraph.formatting),
      revision,
    };
    const propertyChanges = takenOverPropertyChanges(taken);
    if (propertyChanges.isErr()) return propertyChanges;
    const markChange = takenOverMarkChange(taken);
    if (markChange.isErr()) return markChange;
    blocks.push({
      ...paragraph,
      content,
      ...(propertyChanges.value === undefined
        ? {}
        : { propertyChanges: propertyChanges.value }),
      ...(markChange.value === undefined
        ? {}
        : { markPropertyChange: markChange.value }),
      ...(marks.value.taker === undefined
        ? {}
        : { pPrMark: marks.value.taker }),
    });
  }
  const anchorBlocks: Paragraph[] = [];
  if (takesOverMark) {
    // Its property changes went with its former mark.
    const {
      propertyChanges: _moved,
      markPropertyChange: _markMoved,
      ...kept
    } = source;
    anchorBlocks.push({ ...kept, pPrMark: marks.value.anchor });
  }
  // The anchor's own records are carried over and keep their ids, those of
  // its mark on the paragraph that took the mark over; only the records this
  // insertion creates take new ones.
  const moved = new Set(
    takesOverMark ? identitySlots(source).map(slotKey) : [],
  );
  const fresh = freshenIdentities({
    before: takesOverMark ? [source, ...canonicalBlocks] : canonicalBlocks,
    after: [...anchorBlocks, ...blocks],
    newIds: op.newIds ?? {},
    usedElsewhere: () => {
      const elsewhere = usedElsewhere();
      return { has: (key) => !moved.has(key) && elsewhere.has(key) };
    },
  });
  switch (fresh.kind) {
    case "needsIds":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        "Inserted paragraph review records need more revision ids.",
      );
    case "invalidId":
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        "An inserted review record has an invalid revision id.",
      );
    case "fresh": {
      const [replacedAnchor, ...inserted] = takesOverMark
        ? fresh.paragraphs
        : [anchor, ...fresh.paragraphs];
      return Result.ok(replacement(inserted, replacedAnchor));
    }
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
};

type PrepareTrackedParagraphInsertionOptions = {
  document: Document;
  op: InsertBlocksOp;
  /** Source record addressed by the operation and its next virtual neighbor. */
  anchor: Paragraph;
  /** Omitted at the end of the virtual main-story list. */
  following?: Paragraph;
  /** A neighbor in the already validated virtual document represented by both indexes. */
  validatedNeighbor?: Paragraph;
  paragraphIds: Pick<IdentityKeyIndex, "count">;
  recordIds: Pick<IdentityKeyIndex, "count">;
};

type PreparedTrackedParagraphInsertion =
  | { type: "unsupported" }
  | {
      type: "replacement";
      before: Paragraph;
      operation: ReplaceBlocksOp;
      effect: DocumentEditEffects;
    };

/**
 * Prepare one step without splicing. The batch must prove that the virtual
 * main story has a supported trailing section and that the source and its
 * neighbor are adjacent top-level paragraphs (or the source is last).
 */
type FinalMarkTakeoverRefusal = {
  reason: DocumentOpRefusal["reason"];
  message: string;
};

/**
 * Why a tracked insertion after a container's final paragraph cannot take
 * over that paragraph's mark, by the record it would move: a mark under
 * review or a pending property or mark run-property change, each refused by
 * its own reason. A review carrier's role does not move to another paragraph
 * yet.
 */
const finalMarkTakeoverRefusal = (
  paragraph: Paragraph,
): FinalMarkTakeoverRefusal => {
  if (paragraph.reviewCarrier !== undefined)
    return {
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message:
        "A tracked insertion cannot take over the mark of a paragraph that carries a table's review.",
    };
  if (paragraph.pPrMark !== undefined)
    return {
      reason: pendingParagraphMarkReason(paragraph.pPrMark),
      message:
        "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of a pending change.",
    };
  if ((paragraph.propertyChanges?.length ?? 0) === 0)
    return {
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_RUN_PROPERTY_CHANGE,
      message:
        "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark run properties free of a pending change.",
    };
  return {
    reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
    message:
      "A tracked insertion after a paragraph no paragraph follows needs the anchor's properties free of a pending change.",
  };
};

export const prepareTrackedParagraphInsertion = ({
  document,
  op,
  anchor: source,
  following,
  validatedNeighbor,
  paragraphIds,
  recordIds,
}: PrepareTrackedParagraphInsertionOptions): Result<
  PreparedTrackedParagraphInsertion,
  DocumentOpRefusal
> => {
  if (
    op.story !== OP_STORIES.MAIN ||
    op.revision === undefined ||
    op.move === true ||
    op.blocks.length !== 1 ||
    (op.at.type !== "after" && op.at.type !== "before") ||
    source.paraId === undefined ||
    idKey(source.paraId) !== idKey(op.at.blockId) ||
    (document.package.footnotes?.length ?? 0) > 0 ||
    (document.package.endnotes?.length ?? 0) > 0
  )
    return Result.ok({ type: "unsupported" });
  const incoming = op.blocks.at(0);
  const takesOverMark = op.at.type === "after" && following === undefined;
  const anchor = op.at.type === "after" ? (following ?? source) : source;
  if (
    incoming === undefined ||
    anchor.paraId === undefined ||
    !isParaId(anchor.paraId) ||
    source.sectionProperties !== undefined ||
    anchor.sectionProperties !== undefined ||
    incoming.sectionProperties !== undefined ||
    incoming.pPrMark !== undefined ||
    incoming.propertyChanges !== undefined ||
    incoming.markPropertyChange !== undefined ||
    incoming.reviewCarrier !== undefined ||
    !incoming.content.every(
      (node) =>
        node.type === "run" &&
        node.content.every((leaf) => leaf.type === "text"),
    )
  )
    return Result.ok({ type: "unsupported" });
  if (
    takesOverMark &&
    (source.pPrMark !== undefined ||
      source.reviewCarrier !== undefined ||
      source.markPropertyChange !== undefined ||
      (source.propertyChanges?.length ?? 0) > 0)
  )
    return Result.err(
      new DocumentOpRefusal({
        opType: op.type,
        ...finalMarkTakeoverRefusal(source),
      }),
    );
  const canonicalBlocks = [canonicalizeInsertedParagraph(incoming)];
  const replacement = paragraphReplacement({
    op,
    before: !takesOverMark,
    anchor,
  });
  const validate = (prepared: ReplaceBlocksOp) => {
    // An unchanged, validated neighbor contributes no removals. Its ids stay
    // occupied in both virtual indexes, so an inserted copy still collides.
    const [added, retained] = prepared.blocks;
    const includesValidatedNeighbor =
      validatedNeighbor === anchor &&
      retained === anchor &&
      prepared.blocks.length === 2 &&
      added !== undefined;
    const checked = validateReplacement({
      document,
      op: includesValidatedNeighbor
        ? { ...prepared, blocks: [added], expected: [] }
        : prepared,
      before: includesValidatedNeighbor ? [] : [anchor],
      paragraphIds,
      recordIds,
    });
    if (checked.isOk()) return checked;
    const { reason, message } = checked.error;
    return Result.err(
      new DocumentOpRefusal({ opType: op.type, reason, message }),
    );
  };
  const draft = validate(replacement(canonicalBlocks));
  if (draft.isErr()) return draft;
  const prepared = trackedParagraphReplacement({
    op,
    revision: op.revision,
    source,
    anchor,
    takesOverMark,
    canonicalBlocks,
    replacement,
    usedElsewhere: () => ({ has: (key) => recordIds.count(key) > 0 }),
  });
  if (prepared.isErr()) return prepared;
  const operation = prepared.value;
  const validated = validate(operation);
  if (validated.isErr()) return validated;
  return Result.ok({
    type: "replacement",
    before: anchor,
    operation,
    effect: {
      ...nonemptyBlockReplacementEffects(operation, [anchor]),
      touched: {
        modified: takesOverMark ? [source.paraId] : [],
        inserted: incoming.paraId === undefined ? [] : [incoming.paraId],
        removed: [],
      },
    },
  });
};

/**
 * Whether a paragraph before `index` in `blocks` owns the paragraph its
 * pending mark joins. Independent review can carry such a mark across the
 * blocks between, and an insertion cannot record that conditional ownership.
 */
const followsFirstParagraphOwner = (
  blocks: readonly BlockContent[],
  index: number,
): boolean =>
  blocks
    .slice(0, index)
    .some(
      (block) =>
        block.type === "paragraph" && block.pPrMark?.survivor === "first",
    );

type InsertedMarkBeforeTableOptions = {
  blocks: readonly BlockContent[];
  /** The paragraph the new one goes after. */
  source: Paragraph;
  /** Where the table stands in `blocks`, which the new paragraph goes before. */
  index: number;
  boundary: ReturnType<typeof paragraphTableRunInBoundary>;
};

/**
 * Whether a paragraph added right before the table at `index` can hold its
 * own inserted mark: one the table can receive, or, before a table itself
 * under review, one whose rejection removes the emptied paragraph without
 * crossing the table. The place reads alike named from the paragraph before
 * it or from the table, and each insertion there resolves on its own. After a
 * break that goes, the new paragraph would receive the words before it.
 */
const insertedMarkResolvesBeforeTable = ({
  blocks,
  source,
  index,
  boundary,
}: InsertedMarkBeforeTableOptions): boolean =>
  markBeforeTableResolves({ blocks, boundary }) ||
  (boundary.type === "blocked" &&
    boundary.reason === "review" &&
    !markGoes(source) &&
    blockListEndsInParagraph(blocks) &&
    !followsFirstParagraphOwner(blocks, index));

/**
 * Paragraphs some of which end a section (`w:pPr/w:sectPr`), inserted as
 * the paragraphs and then each break set on its paragraph, as setting a
 * section endpoint does. A suggested insertion records each new paragraph's
 * mark inserted, so the break travels with that mark: rejecting it removes
 * the paragraph and joins its section into the one after it, and accepting
 * keeps the break. A section ends only in the main story; a paragraph
 * elsewhere keeps no break an insertion could give it.
 */
const insertSectionBreaks = ({
  document,
  op,
  applyOps,
}: Omit<BlockInsertionOptions, "validateBlockReplacement">): Result<
  DocumentEdit,
  DocumentOpRefusal
> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  if (op.story !== OP_STORIES.MAIN)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      "A section break ends a section of the main story only.",
    );
  const breaks: DocumentOp[] = [];
  const paragraphs = op.blocks.map(({ sectionProperties, ...paragraph }) => {
    if (sectionProperties !== undefined)
      breaks.push({
        type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
        endpoint: { type: "paragraph", blockId: paragraph.paraId ?? "" },
        expected: { type: "omitted" },
        properties: { type: "present", value: sectionProperties },
      });
    return paragraph;
  });
  const applied = applyOps(document, [
    { ...op, blocks: paragraphs },
    ...breaks,
  ]);
  if (applied.isErr())
    return refuse(applied.error.reason, applied.error.message);
  // Suggested after the story's last paragraph, the new last paragraph keeps
  // the story's own mark, which no tracked change records: rejecting would
  // leave its break standing.
  if (
    op.revision !== undefined &&
    paragraphs.some((paragraph, index) => {
      if (op.blocks[index]?.sectionProperties === undefined) return false;
      const mark = findStoryParagraph(
        storyBody(applied.value.document, op.story),
        paragraph.paraId ?? "",
      )?.paragraph.pPrMark;
      return mark?.kind !== "ins" && mark?.kind !== "moveTo";
    })
  )
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
      "A suggested section break needs a paragraph mark the suggestion adds.",
    );
  return Result.ok({
    ...applied.value,
    touched: {
      modified: applied.value.touched.modified,
      inserted: op.blocks.flatMap(({ paraId }) =>
        paraId === undefined ? [] : [paraId],
      ),
      removed: [],
    },
  });
};

export const insertBlocks = ({
  document,
  op,
  applyOps,
  validateBlockReplacement,
}: BlockInsertionOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const refuse = (reason: DocumentOpRefusal["reason"], message: string) =>
    Result.err(new DocumentOpRefusal({ reason, message, opType: op.type }));
  const firstInserted = op.blocks.at(0);
  if (firstInserted === undefined) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
      "No paragraphs to insert.",
    );
  }
  const existingMark = op.blocks.find(
    ({ pPrMark }) => pPrMark !== undefined,
  )?.pPrMark;
  if (op.revision !== undefined && existingMark !== undefined)
    return refuse(
      pendingParagraphMarkReason(existingMark),
      "A tracked insertion would overwrite a pasted paragraph's pending break.",
    );
  if (op.at.type === "containerStart") {
    const located = locateBlockInsertion(document, op);
    if (located.isErr()) return located;
    const first = blockListAt(
      storyBody(document, op.story).content,
      located.value.list,
    ).at(0);
    // A list that starts with a paragraph is named by it: one place, one
    // insertion, whichever way it is addressed.
    if (first?.type === "paragraph" && first.paraId !== undefined)
      return insertBlocks({
        document,
        op: { ...op, at: { type: "before", blockId: first.paraId } },
        applyOps,
        validateBlockReplacement,
      });
  }
  if (
    op.at.type === "beforeTable" ||
    op.at.type === "afterTable" ||
    op.at.type === "containerStart"
  ) {
    if (
      op.blocks.some(({ sectionProperties }) => sectionProperties !== undefined)
    )
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
        "Inserting section breaks requires section operations.",
      );
    if (
      op.blocks.some(({ paraId }) => paraId === undefined || !isParaId(paraId))
    )
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "Every inserted paragraph needs a usable id.",
      );
    const located = locateBlockInsertion(document, op);
    if (located.isErr()) return located;
    let paragraphs = op.blocks.map((paragraph) => {
      if (finalizeParagraphFormatting(paragraph.formatting) !== undefined)
        return paragraph;
      const canonical = { ...paragraph };
      delete canonical.formatting;
      return canonical;
    });
    if (op.revision !== undefined) {
      const list = blockListAt(
        storyBody(document, op.story).content,
        located.value.list,
      );
      if (followsFirstParagraphOwner(list, located.value.index))
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "Tracked table-side insertion cannot cross an earlier pending first-paragraph owner.",
        );
      const following = list.at(located.value.index);
      // After a table that ends a block content control's content, the
      // paragraph stays in the control. Its mark joins nothing there:
      // rejecting removes the emptied paragraph by an empty replacement,
      // reinserted after the table, and never reaches past the control. A
      // control that ends its container is refused below, as a container
      // keeps a final paragraph.
      const endsControl =
        op.at.type === "afterTable" &&
        following === undefined &&
        located.value.list.at(-1)?.kind === "blockSdt";
      if (!endsControl && !list.some((block) => block.type === "paragraph"))
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "Tracked table-side insertion needs a surviving paragraph for its review inverse.",
        );
      if (following?.type === "table") {
        const boundary = paragraphTableRunInBoundary(firstInserted, following);
        if (boundary.type === "blocked" && boundary.reason !== "review")
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            `Tracked table-side insertion cannot resolve a ${boundary.reason} boundary.`,
          );
        if (boundary.type === "receiver" && !blockListEndsInParagraph(list))
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            "Tracked table-side insertion needs a paragraph terminator after its table receiver.",
          );
      } else if (
        !endsControl &&
        following?.type !== "paragraph" &&
        !isOpaqueFollower(following)
      )
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          "Tracked table-side insertion needs a following paragraph or a proven table boundary.",
        );
      const tracked: Paragraph[] = [];
      for (const paragraph of paragraphs) {
        const content = insertedContent({
          paragraph,
          revision: op.revision,
          move: op.move,
        });
        if (content.isErr()) return content;
        tracked.push({
          ...paragraph,
          content: content.value,
          pPrMark: {
            kind: op.move === true ? "moveTo" : "ins",
            info: stampInfo(op.revision),
          },
        });
      }
      const fresh = freshenIdentities({
        before: paragraphs,
        after: tracked,
        newIds: op.newIds ?? {},
        usedElsewhere: () => packageIdentityKeyIndex(document.package),
      });
      switch (fresh.kind) {
        case "fresh":
          paragraphs = fresh.paragraphs;
          break;
        case "needsIds":
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
            "Inserted paragraph review records need more revision ids.",
          );
        case "invalidId":
          return refuse(
            DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
            "An inserted review record has an invalid revision id.",
          );
        default: {
          const unreachable: never = fresh;
          return unreachable;
        }
      }
    }
    const inserted = replaceParagraphs({
      document,
      story: op.story,
      at: located.value,
      count: 0,
      replacement: paragraphs,
    });
    if (inserted.isErr())
      return refuse(inserted.error.reason, inserted.error.message);
    const insertedBody = storyBody(inserted.value, op.story);
    if (
      paragraphs.some((paragraph, index) =>
        hasContainerFinalParagraphMark(insertedBody, {
          list: located.value.list,
          index: located.value.index + index,
          paragraph,
        }),
      )
    )
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        "Insertion would leave a tracked mark on a container-final paragraph.",
      );
    const valid = validateOpsDocument(inserted.value);
    if (valid.isErr()) return refuse(valid.error.reason, valid.error.message);
    return Result.ok({
      document: inserted.value,
      inverse: [
        {
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story: op.story,
          expected: paragraphs,
          blocks: [],
        },
      ],
      touched: {
        modified: [],
        inserted: paragraphs.flatMap(({ paraId }) =>
          paraId === undefined ? [] : [paraId],
        ),
        removed: [],
      },
    });
  }
  if (op.at.type !== "before" && op.at.type !== "after") {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      "Insert before or after a paragraph or table.",
    );
  }
  const body = storyBody(document, op.story);
  const location = findStoryParagraph(body, op.at.blockId);
  if (location === undefined) {
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The insertion anchor does not exist.",
    );
  }
  if (
    op.blocks.some(({ sectionProperties }) => sectionProperties !== undefined)
  )
    return insertSectionBreaks({ document, op, applyOps });
  // Use the following paragraph as the unchanged boundary whenever possible.
  const list = blockListAt(body.content, location.list);
  const following = list[location.index + 1];
  const before = op.at.type === "before" || following?.type === "paragraph";
  const anchor =
    op.at.type === "after" && following?.type === "paragraph"
      ? following
      : location.paragraph;
  // With no paragraph after the anchor, a tracked insertion is recorded as
  // a reader reads it: the anchor's mark is the inserted one, and the last
  // new paragraph holds the mark the anchor had, with the anchor's paragraph
  // properties as its prior ones. Rejecting removes the new text and the
  // anchor's mark, which leaves the anchor as it was; accepting leaves both.
  // Before a table the new paragraph's own mark is inserted where it
  // resolves there (`insertedMarkResolvesBeforeTable`). Otherwise keep the
  // source mark on the new last paragraph, so rejecting joins back into the
  // anchor.
  // At the end of the list, before a range boundary between blocks that no
  // resolution may carry content across, or before any other block, such as
  // a content control, there is no mark after the new paragraph's for it to
  // resolve against, so the anchor's mark is the inserted one.
  // Directly after a paragraph that ends a section, with a block after it,
  // the new paragraphs open the section that block is in: the break stays on
  // its paragraph, with what it ended. Tracked, the new paragraphs hold their
  // own inserted marks and the anchor keeps its own, so its break never goes
  // with a mark this insertion adds.
  const opensNextSection =
    !before &&
    following !== undefined &&
    location.paragraph.sectionProperties !== undefined;
  const takesOverMark =
    op.revision !== undefined &&
    !before &&
    !opensNextSection &&
    (following?.type !== "table" ||
      !insertedMarkResolvesBeforeTable({
        blocks: list,
        source: location.paragraph,
        index: location.index + 1,
        boundary: paragraphTableRunInBoundary(firstInserted, following),
      }));
  // The last new paragraph takes the anchor's mark with its review records
  // (`trackedParagraphReplacement`), except onto the container's new final
  // paragraph, which keeps no record a pending change moves to it.
  if (
    takesOverMark &&
    (location.paragraph.reviewCarrier !== undefined ||
      (following === undefined &&
        (location.paragraph.pPrMark !== undefined ||
          location.paragraph.markPropertyChange !== undefined ||
          (location.paragraph.propertyChanges?.length ?? 0) > 0)))
  ) {
    const { reason, message } = finalMarkTakeoverRefusal(location.paragraph);
    return refuse(reason, message);
  }
  const replacement = paragraphReplacement({ op, before, anchor });
  const canonicalBlocks = op.blocks.map(canonicalizeInsertedParagraph);
  // A direct paragraph after this pending break would be joined into its
  // anchor when the mark is rejected, so the break moves to the last new
  // paragraph (`movedPendingBreak`) where rejecting it there gives the
  // insertion into the rejected document.
  let movedBreak: ReplaceBlocksOp | undefined;
  if (
    op.revision === undefined &&
    op.at.type === "after" &&
    following?.type === "paragraph" &&
    location.paragraph.pPrMark?.survivor === "first"
  ) {
    const moved = movedPendingBreak({
      anchor: location.paragraph,
      after: list.slice(location.index + 1),
      blocks: canonicalBlocks,
    });
    // Any other difference is one only a property change of the pending
    // insertion records, and a direct edit records none: the insertion is
    // resolved first.
    if (moved === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_MARK_INSERTION,
        "Insertion after a pending paragraph mark would change what rejecting that mark leaves.",
      );
    movedBreak = {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: op.story,
      expected: [location.paragraph],
      blocks: moved,
    };
  }
  const sectionBoundaries = opensNextSection
    ? { sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE }
    : {};
  // A draft ordinary insertion has no other work when its main-body splice
  // has no section or note repair and cannot add inline stories or review
  // records. The tracked replacement below still runs full validation.
  const draft: ReplaceBlocksOp = movedBreak ?? {
    ...replacement(canonicalBlocks),
    ...sectionBoundaries,
  };
  const direct =
    op.revision !== undefined &&
    op.story === OP_STORIES.MAIN &&
    location.list.length === 0 &&
    anchor.paraId !== undefined &&
    isParaId(anchor.paraId) &&
    canSpliceMainBodyWithoutSectionRepair(body) &&
    (document.package.footnotes?.length ?? 0) === 0 &&
    (document.package.endnotes?.length ?? 0) === 0 &&
    canonicalBlocks.every(
      ({
        pPrMark,
        propertyChanges,
        markPropertyChange,
        reviewCarrier,
        content,
      }) =>
        pPrMark === undefined &&
        propertyChanges === undefined &&
        markPropertyChange === undefined &&
        reviewCarrier === undefined &&
        content.every(
          (node) =>
            node.type === "run" &&
            node.content.every((leaf) => leaf.type === "text"),
        ),
    )
      ? validateBlockReplacement({ document, op: draft, before: [anchor] }).map(
          () => undefined,
        )
      : applyOps(document, [draft]);
  if (direct.isErr())
    return Result.err(
      new DocumentOpRefusal({
        reason: direct.error.reason,
        message: direct.error.message,
        opType: op.type,
      }),
    );
  // Every other insertion is followed by another inserted paragraph or the
  // unchanged anchor in this block list.
  const lastInserted = canonicalBlocks.at(-1);
  if (
    direct.value !== undefined &&
    !before &&
    lastInserted?.pPrMark !== undefined &&
    endsItsContainer(storyBody(direct.value.document, op.story), {
      list: location.list,
      index: location.index + canonicalBlocks.length,
    })
  )
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
      "Insertion would leave a tracked mark on a container-final paragraph.",
    );
  const touched = {
    // The anchor's mark becomes the inserted one when it takes over, and its
    // own again when a direct insertion moves its pending break on.
    modified:
      (takesOverMark || movedBreak !== undefined) &&
      location.paragraph.paraId !== undefined
        ? [location.paragraph.paraId]
        : [],
    inserted: op.blocks.flatMap(({ paraId }) =>
      paraId === undefined ? [] : [paraId],
    ),
    removed: [],
  };
  if (op.revision === undefined)
    return Result.ok({
      ...(direct.value ??
        panic("An ordinary insertion needs its direct edit.")),
      touched,
    });
  const prepared = trackedParagraphReplacement({
    op,
    revision: op.revision,
    source: location.paragraph,
    anchor,
    takesOverMark,
    canonicalBlocks,
    replacement,
    usedElsewhere: () => packageIdentityKeyIndex(document.package),
  });
  if (prepared.isErr()) return prepared;
  const tracked = applyOps(document, [
    { ...prepared.value, ...sectionBoundaries },
  ]);
  return tracked.isErr()
    ? Result.err(
        new DocumentOpRefusal({
          reason: tracked.error.reason,
          message: tracked.error.message,
          opType: op.type,
        }),
      )
    : Result.ok({ ...tracked.value, touched });
};
