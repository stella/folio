/**
 * `resolveRevision`: accepting or rejecting tracked changes, expanded to the
 * primitive operations that carry it out.
 *
 * The expansion follows the order an editor resolves a review in. First each
 * paragraph's inline changes (insertions, deletions and moves, nested ones
 * included, and run property changes) become one `replaceInline`, and its
 * property changes and kept marks one `setParagraphReview`. Then the marks
 * that go are removed, from the last paragraph to the first: each leaves the
 * next paragraph, which takes the content before it, so a chain of removed
 * marks ends in the paragraph after the last.
 */

import { panic, Result } from "better-result";

import { INSERTION_SEAM_POLICIES } from "../model/content";
import { MAX_REVISION_ID } from "../model/document";
import type {
  BlockContent,
  Deletion,
  Document,
  DocumentBody,
  KeptCutPiece,
  Paragraph,
  ParagraphMarkChangeKind,
  ParagraphPropertyChange,
  Run,
  RunPropertyChange,
  Table,
} from "../model/document";
import { getSingularRunPropertyChange } from "../model/runPropertyReview";
import {
  isOpaqueFollower,
  markBeforeTableResolves,
  paragraphBlockRunInTarget,
  paragraphMarkOutcome,
  paragraphTableRunInBoundary,
  type TableRunInTarget,
} from "./blockDeletion";
import {
  blockListAt,
  endsItsContainer,
  type ParagraphLocation,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { isCommentExhausted } from "./comments";
import {
  acceptCopyOriginAncestor,
  rejectCopyOriginAncestorsInNodes,
} from "./copyOriginAncestors";
import { keptDeletionBoundaryPaths } from "./deletionBoundaries";
import { combineEdits, type DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { canonicalDeferredGroups, deferredGroupKey } from "./identity";
import {
  IDENTITY_SPACES,
  identityKeysIn,
  idKey,
  reservedIdentityKeysIn,
  slotKey,
} from "./ids";
import { joinParagraphSeam } from "./inline";
import {
  holdsKeptCut,
  keptCutFrom,
  keptCutPiece,
  pendingAtEdge,
  settleKeptCuts,
  standsInKeptCut,
  withEdgeKeptCutsWidened,
  withKeptCutPieceConsumed,
  withKeptCutPieceRecorded,
  withKeptCuts,
  withKeptCutsRecorded,
  withoutKeptCuts,
  withoutKeptCutsRecorded,
} from "./keptCut";
import { handJoin, seamHandsOver, withJoinRetired } from "./keptJoins";
import {
  alikeDepth,
  asParagraphContent,
  childNodes,
  fieldResultOf,
  type InlineNode,
  mergeAlike,
  mergeLists,
  ownRunPropertyChanges,
  rebuildNode,
  sameOwnFields,
  withFieldResult,
} from "./leaves";
import { replaceNoteParts } from "./noteLifecycle";
import { automaticNoteMarkRuns, noteReferenceSpans } from "./noteMarks";
import {
  paragraphHoldsContent,
  paragraphHoldsContentOnceAccepted,
} from "./offsets";
import {
  isBlockBoundary,
  isKeptBoundaryCarrier,
  noRelocatedBoundaries,
  pairedCommentRanges,
  paragraphRemains,
  planParagraphBoundaryRemoval,
  type RelocatedBoundaries,
  removedContentBoundaries,
  wellFormedRanges,
  withBoundariesAtEdge,
  withSeamBoundariesInside,
} from "./rangeBoundaries";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { edgeChain, edgeChainPastRemoved } from "./resolvedEdges";
import { reachableRowIds, resolveTableRows } from "./resolveTableRows";
import {
  isAddedRevision,
  isRemovedRevisionNode,
  isTrackedWrapper,
  reviewFieldsOf,
  type TrackedWrapper,
  withMarkFormatting,
} from "./review";
import { runsMergeable } from "./runMerge";
import { resolveSectionChange, sectionResolutionPatch } from "./sectionReview";
import { identitySlots, withInlineIdentity } from "./slots";
import { documentStories } from "./stories";
import { sectionEnds, storyLifecycleTouched } from "./storyLifecycle";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type OpStory,
  type ParagraphReviewFields,
  type ResolveRevisionOp,
  REVISION_DECISIONS,
  type RevisionDecision,
  SECTION_BOUNDARY_POLICIES,
} from "./types";

/** Applies operations atomically; `applyDocumentOps`, passed in to keep the modules acyclic. */
export type ApplyOps = (
  document: Document,
  ops: readonly DocumentOp[],
) => Result<DocumentEdit, DocumentOpRefusal>;

const refusal = (
  op: ResolveRevisionOp,
  reason: DocumentOpRefusal["reason"],
  message: string,
) => new DocumentOpRefusal({ message, reason, opType: op.type });

const precedingSectionOwner = (
  previous: BlockContent | undefined,
  paragraph: Paragraph,
): "unavailable" | "available" | "same" | "occupied" => {
  if (previous?.type !== "paragraph") return "unavailable";
  if (previous.sectionProperties === undefined) return "available";
  return JSON.stringify(previous.sectionProperties) ===
    JSON.stringify(paragraph.sectionProperties)
    ? "same"
    : "occupied";
};

const runParagraphIntoTable = (
  paragraph: Paragraph,
  target: TableRunInTarget,
): Table => {
  let replacement: Paragraph | Table = {
    ...target.receiver,
    content: joinParagraphSeam(paragraph.content, target.receiver.content),
  };
  for (const table of target.tables.toReversed()) {
    const row = table.rows.at(0);
    const cell = row?.cells.at(0);
    if (!row || !cell) panic("A validated table receiver retains its path.");
    replacement = {
      ...table,
      rows: [
        {
          ...row,
          cells: [
            { ...cell, content: [replacement, ...cell.content.slice(1)] },
            ...row.cells.slice(1),
          ],
        },
        ...table.rows.slice(1),
      ],
    };
  }
  if (replacement.type !== "table")
    panic("A table receiver has a nonempty table path.");
  return replacement;
};

const NO_COMMENT_RANGES: ReadonlySet<number> = new Set();
const NO_BOUNDARY_PATHS: ReadonlySet<string> = new Set();

/** Whether a decision removes the content of a tracked change of this kind. */
const removesContent = (
  node: InlineNode,
  decision: RevisionDecision,
): boolean =>
  decision === REVISION_DECISIONS.ACCEPT
    ? !isAddedRevision(node)
    : isAddedRevision(node);

type NoteReference = {
  kind: "footnote" | "endnote";
  id: number;
  added: boolean;
};

/** Note bodies whose selected reference revision is removed by this decision. */
const removedNoteReferences = (
  document: Document,
  op: ResolveRevisionOp,
  ids: ReadonlySet<number>,
): NoteReference[] => {
  const references = new Map<string, NoteReference>();
  for (const { paragraph } of storyParagraphs(storyBody(document, op.story))) {
    for (const { node, ancestors } of noteReferenceSpans(paragraph.content)) {
      const revision = ancestors.find(
        (ancestor) =>
          isTrackedWrapper(ancestor) &&
          ids.has(ancestor.info.id) &&
          removesContent(ancestor, op.decision),
      );
      if (revision === undefined || !isTrackedWrapper(revision)) continue;
      const reference = {
        kind: node.type === "footnoteRef" ? "footnote" : "endnote",
        id: node.id,
        added: isAddedRevision(revision),
      } as const;
      references.set(`${reference.kind}:${reference.id}`, reference);
    }
  }
  return [...references.values()];
};

const hasNoteReference = (
  document: Document,
  reference: NoteReference,
): boolean =>
  documentStories(document).some((story) =>
    storyParagraphs(storyBody(document, story)).some(({ paragraph }) =>
      noteReferenceSpans(paragraph.content).some(
        ({ node }) =>
          node.type ===
            (reference.kind === "footnote" ? "footnoteRef" : "endnoteRef") &&
          node.id === reference.id,
      ),
    ),
  );

/**
 * The paragraph a rejected mark that keeps the first paragraph leaves: the
 * first paragraph's identity, and everything the mark that stays carries,
 * which is the next paragraph's. Where that mark is the first paragraph's
 * former one, the first paragraph comes back as it was once the reviews
 * pending on it resolve; where it is another break the same edit inserted,
 * that break says itself which paragraph its rejection keeps.
 */
const joinedKeepingFirst = (
  paragraph: Paragraph,
  next: Paragraph,
  content: Paragraph["content"],
  receivesSection: boolean,
): Paragraph => {
  const {
    paraId: _paraId,
    textId: _textId,
    preservedAttributes: _preservedAttributes,
    renderedPageBreakBefore: _renderedPageBreakBefore,
    sectionProperties: _sectionProperties,
    ...marked
  } = next;
  // A section the rejected mark ended moves into the joined paragraph;
  // otherwise the section break is the one the staying mark ends.
  const sectionProperties = receivesSection
    ? paragraph.sectionProperties
    : next.sectionProperties;
  return {
    ...marked,
    content,
    ...(paragraph.paraId === undefined ? {} : { paraId: paragraph.paraId }),
    ...(paragraph.textId === undefined ? {} : { textId: paragraph.textId }),
    ...(paragraph.preservedAttributes === undefined
      ? {}
      : { preservedAttributes: paragraph.preservedAttributes }),
    ...(paragraph.renderedPageBreakBefore === undefined
      ? {}
      : { renderedPageBreakBefore: paragraph.renderedPageBreakBefore }),
    ...(sectionProperties === undefined ? {} : { sectionProperties }),
  };
};

/** Whether a mark of this kind recorded a break being added. */
const markWasAdded = (kind: ParagraphMarkChangeKind): boolean =>
  kind === "ins" || kind === "moveTo";

type PropertyChange = RunPropertyChange | ParagraphPropertyChange;

/**
 * The property changes left once some are removed, and the formatting to
 * restore when the last ones go.
 *
 * Each change records the formatting before it, so a change removed before
 * one that is kept hands its record to that one, which then says what the
 * formatting was before both. Removed changes at the end restore what the
 * first of them recorded.
 */
type RemovedPropertyChanges<Change extends PropertyChange> =
  | { kind: "unchanged" }
  | { kind: "keepLive"; remaining: Change[] }
  | {
      kind: "restore";
      remaining: Change[];
      previous: Change["previousFormatting"] | undefined;
    };

const removePropertyChanges = <Change extends PropertyChange>(
  changes: readonly Change[],
  remove: (change: Change) => boolean,
): RemovedPropertyChanges<Change> => {
  const remaining: Change[] = [];
  let removedPrevious: Change["previousFormatting"] | undefined = undefined;
  let removedAny = false;
  let removing = false;
  const removedRunChanges: RunPropertyChange[] = [];
  for (const change of changes) {
    if (remove(change)) {
      if (!removing) {
        removedPrevious = change.previousFormatting;
        removing = true;
      }
      if (change.type === "runPropertyChange") removedRunChanges.push(change);
      removedAny = true;
      continue;
    }
    if (!removing) {
      remaining.push(change);
      continue;
    }
    const rebased: Change = { ...change };
    delete rebased.previousFormatting;
    if (removedPrevious !== undefined) {
      Object.assign(rebased, { previousFormatting: removedPrevious });
    }
    if (rebased.type === "runPropertyChange" && removedRunChanges.length > 0) {
      // A later suggestion inherits the cuts of a rejected earlier one,
      // so rejecting it can still restore the original run topology.
      if (
        rebased.boundaryJoins === undefined ||
        removedRunChanges.some((removed) => removed.boundaryJoins === undefined)
      ) {
        delete rebased.boundaryJoins;
      } else {
        rebased.boundaryJoins = [
          ...new Set([
            ...rebased.boundaryJoins,
            ...removedRunChanges.flatMap(
              (removed) => removed.boundaryJoins ?? [],
            ),
          ]),
        ];
      }
    }
    removedRunChanges.length = 0;
    remaining.push(rebased);
    removedPrevious = undefined;
    removing = false;
  }
  if (!removedAny) return { kind: "unchanged" };
  if (!removing) return { kind: "keepLive", remaining };
  return { kind: "restore", remaining, previous: removedPrevious };
};

/** Records resolution emptied, as distinct from ones that held nothing before it. */
type Emptied = { has: (node: InlineNode) => boolean };

const NOTHING_EMPTIED: Emptied = { has: () => false };

type Resolution = {
  ids: ReadonlySet<number> & { add: (id: number) => void };
  decision: RevisionDecision;
  /** The containers this resolution has emptied so far. */
  emptied: WeakSet<InlineNode>;
  /** Actual payload edges accepted after source identities were restored. */
  acceptedClosed: Record<"first" | "last", WeakSet<InlineNode>>;
  /** Unconsumed source-cut edges on the actual rebuilt container. */
  cutEdges: WeakMap<InlineNode, CutEdges>;
  /** Removal cuts supplied entirely by a deferred seam from a different edit. */
  deferredOnlyRemovals: WeakSet<NonNullable<TrackedWrapper["resolutionJoins"]>>;
  /** Comment ranges with a start and an end in the paragraphs resolved. */
  pairedComments: ReadonlySet<number>;
  /** Source occurrences retained by the current paragraph's accepted deletions. */
  keptDeletionPaths: ReadonlySet<string>;
  /** Boundaries this resolution kept from removed content. */
  relocated: RelocatedBoundaries;
  /**
   * Identity handoffs whose retained piece was not beside the removed prefix:
   * a replacement's new content, or a paragraph break it inserted, stands
   * between them. They reach that piece later in the same story.
   */
  outstanding: RetainedIdentity[];
};

const createResolution = (
  ids: Set<number>,
  decision: RevisionDecision,
): Resolution => ({
  ids,
  decision,
  emptied: new WeakSet(),
  acceptedClosed: { first: new WeakSet(), last: new WeakSet() },
  cutEdges: new WeakMap(),
  deferredOnlyRemovals: new WeakSet(),
  pairedComments: NO_COMMENT_RANGES,
  keptDeletionPaths: NO_BOUNDARY_PATHS,
  relocated: noRelocatedBoundaries(),
  outstanding: [],
});

const resolveRun = (run: Run, { ids, decision }: Resolution): Run => {
  const changes = run.propertyChanges ?? [];
  const selected = (change: (typeof changes)[number]): boolean =>
    ids.has(change.info.id);
  if (!changes.some(selected)) {
    return run;
  }
  const next: Run = { ...run };
  let remaining = changes.filter((change) => !selected(change));
  if (decision === REVISION_DECISIONS.REJECT) {
    const removed = removePropertyChanges(changes, selected);
    if (removed.kind !== "unchanged") {
      remaining = removed.remaining;
    }
    if (removed.kind === "restore") {
      if (removed.previous === undefined) {
        delete next.formatting;
      } else {
        next.formatting = removed.previous;
      }
    }
  }
  if (remaining.length > 0) {
    next.propertyChanges = remaining;
  } else {
    delete next.propertyChanges;
  }
  return next;
};

/** Records with the run property changes of their runs resolved, the same list when none is. */
const resolveRunsIn = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): readonly InlineNode[] => {
  let changed = false;
  const out = nodes.map((node) => {
    let next: InlineNode = node;
    if (node.type === "run") next = resolveRun(node, resolution);
    else {
      const children = childNodes(node);
      const inner =
        children === undefined ? children : resolveRunsIn(children, resolution);
      if (inner !== children && inner !== undefined)
        next = rebuildNode(node, inner);
    }
    if (next !== node) changed = true;
    return next;
  });
  return changed ? out : nodes;
};

/**
 * A field with the run property changes on the runs of its result resolved:
 * a run property patch over the field records them there.
 */
const resolveFieldRuns = (
  field: Extract<InlineNode, { type: "simpleField" | "complexField" }>,
  resolution: Resolution,
): InlineNode => {
  const result = fieldResultOf(field) ?? [];
  const next = resolveRunsIn(result, resolution);
  return next === result ? field : withFieldResult(field, next);
};

/**
 * Whether `empty` is a piece of `other` that resolution left with nothing: a
 * container of the same kind and fields that the resolution emptied, where
 * the other holds something or was emptied too (two emptied pieces are one
 * piece, whichever emptied first). A container that held nothing before (an
 * empty content control or hyperlink) is markup of its own and is never merged.
 */
const emptyPieceOf = (
  empty: InlineNode,
  other: InlineNode,
  emptied: Emptied,
): boolean => {
  const emptyChildren = childNodes(empty);
  const otherChildren = childNodes(other);
  return (
    empty.type !== "run" &&
    emptyChildren !== undefined &&
    otherChildren !== undefined &&
    emptyChildren.length === 0 &&
    (otherChildren.length > 0 || emptied.has(other)) &&
    emptied.has(empty) &&
    sameOwnFields(empty, other)
  );
};

/**
 * An emptied piece folded into its neighbour, which keeps the first one's ids;
 * `undefined` when neither is one. When both are, the first stays as it is,
 * still emptied.
 */
const foldedPiece = (
  left: InlineNode,
  right: InlineNode,
  emptied: Emptied,
): InlineNode | undefined => {
  if (emptyPieceOf(right, left, emptied)) return left;
  if (emptyPieceOf(left, right, emptied))
    return rebuildNode(left, childNodes(right) ?? []);
  return undefined;
};

/**
 * Two records meeting beside a container resolution emptied, with no change
 * resolved between them: the emptied piece folds into its neighbour, down the
 * chain of alike containers at their facing edges when it sits inside one.
 * Nothing else merges there.
 */
const foldAtSeam = (
  left: InlineNode,
  right: InlineNode,
  emptied: Emptied,
): InlineNode[] => {
  const folded = foldedPiece(left, right, emptied);
  if (folded !== undefined) return [folded];
  const leftChildren = childNodes(left);
  const rightChildren = childNodes(right);
  const last = leftChildren?.at(-1);
  const first = rightChildren?.at(0);
  if (
    left.type === "run" ||
    leftChildren === undefined ||
    rightChildren === undefined ||
    last === undefined ||
    first === undefined ||
    !sameOwnFields(left, right)
  ) {
    return [left, right];
  }
  const inner = foldAtSeam(last, first, emptied);
  if (inner.length !== 1) return [left, right];
  return [
    rebuildNode(left, [
      ...leftChildren.slice(0, -1),
      ...inner,
      ...rightChildren.slice(1),
    ]),
  ];
};

/** A pending formatting review retains a recorded seam until it can be joined. */
type DeferredRunReviewSeamOptions = {
  left: InlineNode;
  right: InlineNode;
  depth: number;
};
const deferRunReviewSeam = ({
  left,
  right,
  depth,
}: DeferredRunReviewSeamOptions): InlineNode[] => {
  if (depth === 0) return [left, right];
  if (left.type !== "run" || right.type !== "run") {
    const leftChildren = childNodes(left);
    const rightChildren = childNodes(right);
    const last = leftChildren?.at(-1);
    const first = rightChildren?.at(0);
    if (
      leftChildren === undefined ||
      rightChildren === undefined ||
      last === undefined ||
      first === undefined ||
      !sameOwnFields(left, right)
    )
      return [left, right];
    const deferred = deferRunReviewSeam({
      left: last,
      right: first,
      depth: depth - 1,
    });
    const deferredLeft =
      deferred.at(0) ?? panic("A deferred seam retains its left node.");
    const deferredRight =
      deferred.at(1) ?? panic("A deferred seam retains its right node.");
    return [
      deferredLeft === last
        ? left
        : rebuildNode(left, [...leftChildren.slice(0, -1), deferredLeft]),
      deferredRight === first
        ? right
        : rebuildNode(right, [deferredRight, ...rightChildren.slice(1)]),
    ];
  }
  if (sameOwnFields(left, right)) return [left, right];
  const retain = (run: Run, side: "before" | "after"): Run => {
    const change = getSingularRunPropertyChange(run.propertyChanges);
    if (
      change === undefined ||
      change.boundaryJoins === undefined ||
      change.boundaryJoins.includes(side)
    )
      return run;
    return {
      ...run,
      propertyChanges: [
        { ...change, boundaryJoins: [...change.boundaryJoins, side] },
      ],
    };
  };
  return [retain(left, "after"), retain(right, "before")];
};

/**
 * Two records meeting where a change was resolved, merged as far as they are
 * alike. A piece of a cut container that resolution emptied (one `emptied`
 * holds) goes into the piece it was cut from, which keeps the first one's ids.
 * Two alike containers merged meet inside too, and that seam is merged the
 * same way, emptied pieces included, so a merge leaves no seam it would merge
 * again.
 */
export const mergeAtSeam = (
  left: InlineNode,
  right: InlineNode,
  emptied: Emptied = NOTHING_EMPTIED,
): InlineNode[] => {
  if (alikeDepth(left, right) > 0) {
    const leftChildren = childNodes(left);
    const rightChildren = childNodes(right);
    return leftChildren === undefined || rightChildren === undefined
      ? mergeAlike([left], [right])
      : [rebuildNode(left, mergedAtSeam(leftChildren, rightChildren, emptied))];
  }
  const folded = foldedPiece(left, right, emptied);
  return folded === undefined
    ? deferRunReviewSeam({ left, right, depth: Number.POSITIVE_INFINITY })
    : [folded];
};

/** Two lists end to end, merged at the seam as {@link mergeAtSeam} merges. */
const mergedAtSeam = (
  left: readonly InlineNode[],
  right: readonly InlineNode[],
  emptied: Emptied,
): InlineNode[] => {
  const last = left.at(-1);
  const first = right.at(0);
  if (last === undefined || first === undefined) {
    return [...left, ...right];
  }
  return [
    ...left.slice(0, -1),
    ...mergeAtSeam(last, first, emptied),
    ...right.slice(1),
  ];
};

type RetainedIdentity = NonNullable<
  NonNullable<Deletion["resolutionJoins"]>["retainedAfter"]
>[number];

const hasSlots = (
  node: InlineNode,
  expected: RetainedIdentity["target"],
): boolean => {
  const actual = identitySlots(node);
  return expected.every((target) =>
    actual.some((slot) => slot.space === target.space && slot.id === target.id),
  );
};

/** Only recorded slot references change; arbitrary numbers in the model never do. */
type RestoreRetainedIdentitiesOptions = {
  nodes: readonly InlineNode[];
  transfers: readonly RetainedIdentity[];
  resolution: Resolution;
};
const restoreRetainedIdentities = ({
  nodes,
  transfers,
  resolution,
}: RestoreRetainedIdentitiesOptions): {
  nodes: InlineNode[];
  unmatched: RetainedIdentity[];
} => {
  const references = new Map<string, RetainedIdentity["source"][number]>();
  const unmatched: RetainedIdentity[] = [];
  let out = [...nodes];
  for (const transfer of transfers) {
    const source = transfer.source.map(
      (slot) => references.get(slotKey(slot)) ?? slot,
    );
    let matched = false;
    const restore = (node: InlineNode): InlineNode => {
      const matches = hasSlots(node, transfer.target);
      matched ||= matches;
      let next = matches
        ? withInlineIdentity(
            node,
            identitySlots(node).map((slot) => {
              const index = transfer.target.findIndex(
                (target) =>
                  slot.space === target.space && slot.id === target.id,
              );
              return index < 0
                ? slot.id
                : (source.at(index)?.id ??
                    panic(
                      "A validated identity transfer has its source slot.",
                    ));
            }),
          )
        : node;
      const children = childNodes(next);
      if (children !== undefined)
        next = rebuildNode(next, children.map(restore));
      return next;
    };
    out = out.map(restore);
    if (!matched) {
      unmatched.push({ ...transfer, source });
      continue;
    }
    for (const [index, target] of transfer.target.entries()) {
      const restored = source.at(index);
      if (restored === undefined || restored.space !== target.space)
        panic("Validated identity transfers retain slot order and space.");
      references.set(slotKey(target), restored);
      // Selection follows the actual retained record through a source-ID
      // transfer, rather than resolving an unrelated later use of its old ID.
      if (
        target.space === IDENTITY_SPACES.REVISION &&
        resolution.ids.has(target.id)
      )
        resolution.ids.add(restored.id);
    }
  }
  // A later pending deletion may remove the fragment whose identity just
  // returned. Rebase its explicit slot references, preserving that lineage.
  const rebase = (node: InlineNode): InlineNode => {
    let next = node;
    if (isTrackedWrapper(node) && node.resolutionJoins !== undefined) {
      const joins = Object.assign({}, node.resolutionJoins);
      if (joins.retainedAfter !== undefined) {
        joins.retainedAfter = joins.retainedAfter.map((entry) => ({
          depth: entry.depth,
          source: entry.source.map(
            (slot) => references.get(slotKey(slot)) ?? slot,
          ),
          target: entry.target.map(
            (slot) => references.get(slotKey(slot)) ?? slot,
          ),
        }));
      }
      if (joins.retainedEnclosing !== undefined) {
        joins.retainedEnclosing = joins.retainedEnclosing.map((entry) => ({
          levels: entry.levels,
          source: entry.source.map(
            (slot) => references.get(slotKey(slot)) ?? slot,
          ),
          target: entry.target.map(
            (slot) => references.get(slotKey(slot)) ?? slot,
          ),
        }));
      }
      if (joins.deferredRemove !== undefined) {
        joins.deferredRemove = canonicalDeferredGroups(
          joins.deferredRemove.map((group) => ({
            ...group,
            blockers: [
              ...new Set(
                group.blockers.map((id) => {
                  const restored = references.get(
                    slotKey({ space: IDENTITY_SPACES.REVISION, id }),
                  );
                  return restored?.space === IDENTITY_SPACES.REVISION
                    ? restored.id
                    : id;
                }),
              ),
            ],
          })),
        );
      }
      next = Object.assign({}, node, { resolutionJoins: joins });
    }
    const children = childNodes(next);
    return children === undefined
      ? next
      : rebuildNode(next, children.map(rebase));
  };
  return { nodes: out.map(rebase), unmatched };
};

/**
 * Hand outstanding identities to the pieces they name among `nodes`, content
 * that follows the removed prefixes in the story. A target id names one
 * record, so the piece is found wherever the replacement left it.
 */
const restoreOutstandingIdentities = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): readonly InlineNode[] => {
  if (resolution.outstanding.length === 0) return nodes;
  const restored = restoreRetainedIdentities({
    nodes,
    transfers: resolution.outstanding,
    resolution,
  });
  resolution.outstanding = restored.unmatched;
  return restored.nodes;
};

const isIdentitySlot = (
  slot: unknown,
): slot is RetainedIdentity["source"][number] => {
  if (
    typeof slot !== "object" ||
    slot === null ||
    !("space" in slot) ||
    !("id" in slot)
  )
    return false;
  // Revision/control ids allow zero; note-record reserved ids do not apply here.
  const retainedId = slot.id;
  return (
    (slot.space === IDENTITY_SPACES.REVISION ||
      slot.space === IDENTITY_SPACES.CONTROL) &&
    typeof retainedId === "number" &&
    Number.isInteger(retainedId) &&
    retainedId >= 0 &&
    retainedId <= MAX_REVISION_ID
  );
};

/**
 * Handoffs to records above a wrapper name paired slots one level up or more;
 * whether that record holds them is checked where it is resolved.
 */
const validEnclosingHandoffs = (handoffs: unknown): boolean =>
  handoffs === undefined ||
  (Array.isArray(handoffs) &&
    handoffs.every((entry: unknown) => {
      if (
        typeof entry !== "object" ||
        entry === null ||
        !("levels" in entry) ||
        !("source" in entry) ||
        !("target" in entry)
      )
        return false;
      const { levels, source, target } = entry;
      return (
        typeof levels === "number" &&
        Number.isInteger(levels) &&
        levels >= 1 &&
        levels <= MAX_REVISION_ID &&
        Array.isArray(source) &&
        Array.isArray(target) &&
        source.length > 0 &&
        source.length === target.length &&
        source.every(isIdentitySlot) &&
        target.every(isIdentitySlot) &&
        new Set(source.map(slotKey)).size === source.length &&
        new Set(target.map(slotKey)).size === target.length &&
        source.every((slot, index) => slot.space === target.at(index)?.space)
      );
    }));

/** Journal-supplied provenance must name actual source slots at its recorded depth. */
const validRetainedIdentities = (node: InlineNode): boolean => {
  if (!isTrackedWrapper(node) || node.resolutionJoins === undefined)
    return true;
  const joins = node.resolutionJoins;
  if (
    typeof joins !== "object" ||
    joins === null ||
    ![joins.before, joins.after, joins.remove].every(
      (depth) =>
        Number.isInteger(depth) && depth >= 0 && depth <= MAX_REVISION_ID,
    )
  )
    return false;
  if (joins.deferredRemove !== undefined) {
    const groups = joins.deferredRemove;
    if (!Array.isArray(groups) || groups.length === 0) return false;
    const groupKeys = new Set<string>();
    for (const deferred of groups) {
      if (
        typeof deferred !== "object" ||
        deferred === null ||
        !Number.isInteger(deferred.depth) ||
        deferred.depth < 0 ||
        deferred.depth > MAX_REVISION_ID ||
        !Array.isArray(deferred.blockers) ||
        deferred.blockers.length === 0 ||
        !deferred.blockers.every(
          (id: unknown) =>
            typeof id === "number" &&
            Number.isInteger(id) &&
            id >= 0 &&
            id <= MAX_REVISION_ID,
        ) ||
        new Set(deferred.blockers).size !== deferred.blockers.length ||
        !deferred.blockers.includes(node.info.id) ||
        // An insertion holds a seam it stands in; a deletion or move source
        // holds one at the edge where its piece meets it.
        (isAddedRevision(node)
          ? deferred.edge !== undefined
          : deferred.edge !== "start" && deferred.edge !== "end")
      )
        return false;
      const key = deferredGroupKey(deferred);
      if (groupKeys.has(key)) return false;
      groupKeys.add(key);
    }
  }
  if (!validEnclosingHandoffs(joins.retainedEnclosing)) return false;
  if (joins.retainedAfter === undefined) return true;
  if (!Array.isArray(joins.retainedAfter)) return false;
  type AtDepthOptions = {
    records: readonly InlineNode[];
    depth: number;
    source: RetainedIdentity["source"];
  };
  const atDepth = ({ records, depth, source }: AtDepthOptions): boolean => {
    if (depth === 0) return records.some((record) => hasSlots(record, source));
    return records.some((record) =>
      atDepth({ records: childNodes(record) ?? [], depth: depth - 1, source }),
    );
  };
  const validSlot = (
    slot: unknown,
  ): slot is RetainedIdentity["source"][number] => {
    if (
      typeof slot !== "object" ||
      slot === null ||
      !("space" in slot) ||
      !("id" in slot)
    )
      return false;
    // Revision/control ids allow zero; note-record reserved ids do not apply here.
    const retainedId = slot.id;
    return (
      (slot.space === IDENTITY_SPACES.REVISION ||
        slot.space === IDENTITY_SPACES.CONTROL) &&
      typeof retainedId === "number" &&
      Number.isInteger(retainedId) &&
      retainedId >= 0 &&
      retainedId <= MAX_REVISION_ID
    );
  };
  const validEntry = (entry: unknown): boolean => {
    if (
      typeof entry !== "object" ||
      entry === null ||
      !("depth" in entry) ||
      !("source" in entry) ||
      !("target" in entry)
    )
      return false;
    const { depth, source, target } = entry;
    if (
      typeof depth !== "number" ||
      !Number.isInteger(depth) ||
      depth < 0 ||
      depth > MAX_REVISION_ID ||
      !Array.isArray(source) ||
      !Array.isArray(target) ||
      !source.every(validSlot) ||
      !target.every(validSlot) ||
      source.length === 0 ||
      source.length !== target.length
    )
      return false;
    return (
      new Set(source.map(slotKey)).size === source.length &&
      new Set(target.map(slotKey)).size === target.length &&
      source.every((slot, index) => slot.space === target.at(index)?.space) &&
      atDepth({ records: node.content, depth, source })
    );
  };
  return joins.retainedAfter.every(validEntry);
};

/**
 * A range boundary standing on its own in a list, held by no record, or a
 * pending deletion a rejection kept holding only such boundaries.
 */
const isStandingBoundary = (node: InlineNode | undefined): boolean =>
  node !== undefined && isKeptBoundaryCarrier(node);

/**
 * How many range boundaries kept records stand for: a pending deletion
 * holding only kept ones stands for those it holds, so a count recorded
 * while it is pending still names them once rejecting it leaves them
 * standing on their own.
 */
const keptBoundaryCount = (nodes: readonly InlineNode[]): number =>
  nodes.reduce(
    (sum, node) =>
      sum + (node.type === "deletion" ? keptBoundaryCount(node.content) : 1),
    0,
  );

/**
 * Kept markers standing at a seam: the earliest revision they were kept from
 * and how many boundaries they stand for.
 */
type KeptAtSeam = { from: string; count: number };

/** Kept records standing in a row, and the boundaries they stand for. */
type KeptRun = { records: number; boundaries: number };
const NO_KEPT_RUN: KeptRun = { records: 0, boundaries: 0 };

/**
 * The kept records standing from `from`, stepping by `step`, that stand for
 * at most `limit` boundaries; `at` reads a position, `undefined` past what
 * may count.
 */
const keptRun = ({
  at,
  from,
  step,
  limit,
}: {
  at: (index: number) => InlineNode | undefined;
  from: number;
  step: 1 | -1;
  limit: number;
}): KeptRun => {
  let records = 0;
  let boundaries = 0;
  for (;;) {
    const node = at(from + step * records);
    if (node === undefined || !isStandingBoundary(node)) break;
    const more = keptBoundaryCount([node]);
    if (boundaries + more > limit) break;
    boundaries += more;
    records += 1;
  }
  return { records, boundaries };
};

/**
 * The content an acceptance leaves in a record that is no tracked change:
 * when it took everything there but pending insertions, those insertions
 * are marked so rejecting the last of them removes the record too, as
 * accepting that content after rejecting them does.
 */
const withEmptiedHolder = (
  node: InlineNode,
  children: readonly InlineNode[],
  nodes: InlineNode[],
  resolution: Resolution,
): InlineNode[] => {
  const pendingInsertion = (child: InlineNode): boolean =>
    isAddedRevision(child) &&
    !resolution.ids.has(child.info.id) &&
    child.resolutionJoins !== undefined;
  if (
    resolution.decision !== REVISION_DECISIONS.ACCEPT ||
    isTrackedWrapper(node) ||
    nodes.length === 0 ||
    !nodes.every(pendingInsertion) ||
    children.every(pendingInsertion)
  )
    return nodes;
  return nodes.map((child) => {
    if (
      !isAddedRevision(child) ||
      child.resolutionJoins === undefined ||
      child.resolutionJoins.emptiedHolder === true
    )
      return child;
    const marked = Object.assign({}, child, {
      resolutionJoins: Object.assign({}, child.resolutionJoins, {
        emptiedHolder: true as const,
      }),
    });
    carryResolutionState(resolution, child, marked);
    return marked;
  });
};

/** What has become of a deferred seam's blocker. */
type BlockerFate = "pending" | "accepted" | "removed";

type DeferredRemovalSettlement = {
  fateOf: (blocker: number) => BlockerFate;
  deferredOnlyRemovals?: Resolution["deferredOnlyRemovals"];
};

/**
 * The one rule for a deferred seam whose blockers leave the story. An
 * accepted blocker retires the seam; a removed one, rejected or emptied by
 * resolving the changes it held, no longer stands between the pieces, and
 * once none stands the seam joins when its holder is removed.
 */
const settleDeferredRemovals = (
  nodes: readonly InlineNode[],
  options: DeferredRemovalSettlement,
): readonly InlineNode[] => {
  const { fateOf, deferredOnlyRemovals } = options;
  let changed = false;
  const out: InlineNode[] = [];
  /** Seams whose every blocker is accepted, by key: see placeAcceptedSeamKept. */
  const acceptedKept = new Map<string, DeferredRemovalGroup>();
  for (const node of nodes) {
    let next = node;
    if (
      isTrackedWrapper(node) &&
      node.resolutionJoins?.deferredRemove !== undefined
    ) {
      const groups = node.resolutionJoins.deferredRemove;
      if (
        groups.some((group) =>
          group.blockers.some((id) => fateOf(id) !== "pending"),
        )
      ) {
        const joins = Object.assign({}, node.resolutionJoins);
        const remainingGroups: DeferredRemovalGroup[] = [];
        for (const group of groups) {
          const fates = group.blockers.map(fateOf);
          if (fates.every((fate) => fate === "pending")) {
            remainingGroups.push(group);
            continue;
          }
          if (
            group.edge === undefined &&
            fates.every((fate) => fate === "accepted")
          )
            acceptedKept.set(deferredGroupKey(group), group);
          if (fates.includes("accepted")) continue;
          const remaining = group.blockers.filter(
            (_, index) => fates[index] === "pending",
          );
          if (remaining.length === 0) {
            if (group.edge === "start")
              joins.before = Math.max(joins.before, group.depth);
            else if (group.edge === "end")
              joins.after = Math.max(joins.after, group.depth);
            else joins.remove = Math.max(joins.remove, group.depth);
          } else remainingGroups.push({ ...group, blockers: remaining });
        }
        delete joins.deferredRemove;
        if (remainingGroups.length > 0)
          joins.deferredRemove = canonicalDeferredGroups(remainingGroups);
        if (
          (node.resolutionJoins.remove === 0 && joins.remove > 0) ||
          deferredOnlyRemovals?.has(node.resolutionJoins) === true
        )
          deferredOnlyRemovals?.add(joins);
        next = Object.assign({}, node, { resolutionJoins: joins });
      }
    }
    const children = childNodes(next);
    const settled =
      children === undefined
        ? children
        : settleDeferredRemovals(children, options);
    if (children !== undefined && settled !== children)
      next = rebuildNode(next, settled ?? []);
    changed ||= next !== node;
    out.push(next);
  }
  for (const group of acceptedKept.values())
    changed = placeAcceptedSeamKept({ out, group, fateOf }) || changed;
  return changed ? out : nodes;
};

/**
 * Whether resolving as `fateOf` says leaves nothing of `node`: a tracked
 * change whose content goes, or a record holding only such changes.
 */
const emptiedBy = (
  node: InlineNode,
  fateOf: DeferredRemovalSettlement["fateOf"],
): boolean => {
  if (isTrackedWrapper(node)) {
    const fate = fateOf(node.info.id);
    return isAddedRevision(node) ? fate === "removed" : fate === "accepted";
  }
  const children = childNodes(node);
  return (
    children !== undefined &&
    children.length > 0 &&
    children.every((child) => emptiedBy(child, fateOf))
  );
};

/**
 * Boundaries a seam kept standing beside pending insertions made after the
 * revision they were kept from, once every one of those insertions is
 * accepted: they go into the piece of the cut record they touch, as
 * rejecting that revision after accepting the insertions places them. They
 * are the kept markers beside the blockers naming the blockers' side, or
 * both, as their cut's piece. They stand where they are when no such piece
 * can hold them. `out` is changed in place; whether it changed.
 */
const placeAcceptedSeamKept = ({
  out,
  group,
  fateOf,
}: {
  out: InlineNode[];
  group: DeferredRemovalGroup;
  fateOf: DeferredRemovalSettlement["fateOf"];
}): boolean => {
  const blocker = (node: InlineNode | undefined): boolean =>
    node !== undefined &&
    isTrackedWrapper(node) &&
    group.blockers.includes(node.info.id);
  const indices = out.flatMap((node, index) => (blocker(node) ? [index] : []));
  const first = indices.at(0);
  const last = indices.at(-1);
  if (first === undefined || last === undefined) return false;
  /**
   * How many kept markers from `from`, stepping by `step`, name `piece` or
   * both as their cut's piece.
   */
  const namedFrom = (
    from: number,
    step: 1 | -1,
    piece: KeptCutPiece,
  ): number => {
    let records = 0;
    for (let index = from; ; index += step) {
      const node = index < 0 ? undefined : out[index];
      const named = node === undefined ? undefined : keptCutPiece(node);
      if (
        node === undefined ||
        !standsInKeptCut(node) ||
        (named !== piece && named !== "both")
      )
        return records;
      records += 1;
    }
  };
  // The boundaries stand just after the blockers before the seam, or just
  // before those after it.
  const preceding = namedFrom(first - 1, -1, "after");
  const forward = (index: number) => ({
    index,
    records: namedFrom(index, 1, "before"),
  });
  const placed = [
    forward(last + 1),
    { index: first - preceding, records: preceding },
    ...indices.map((index) => forward(index + 1)),
  ].find(
    ({ index, records }) =>
      records > 0 &&
      (blocker(out.at(index - 1)) || blocker(out.at(index + records))),
  );
  if (placed === undefined) return false;
  const { index: start, records } = placed;
  // The cut's piece across the blockers must stand once they are resolved:
  // with nothing left of it, the cut has one side and they stand beside it,
  // as resolving its emptying first leaves them.
  const far = start < first ? out.at(last + 1) : out.at(first - 1);
  if (far === undefined || emptiedBy(far, fateOf)) return false;
  // The accepted blocker beside them stays between the pieces. Rejecting the
  // revision after accepting it puts them into the piece only when the
  // blocker's record facing them is alike the piece, or when the blocker's
  // own join there reaches deeper than the cut, so that it continues the
  // record the cut split; otherwise they stand beside it.
  const blockerBefore = blocker(out.at(start - 1));
  const beside = out.at(blockerBefore ? start - 1 : start + records);
  const piece = out.at(blockerBefore ? start + records : start - 1);
  const edgeRecord =
    beside === undefined
      ? undefined
      : childNodes(beside)?.at(blockerBefore ? -1 : 0);
  const besideJoins =
    beside !== undefined && isTrackedWrapper(beside)
      ? beside.resolutionJoins
      : undefined;
  const ownJoin =
    (blockerBefore ? besideJoins?.after : besideJoins?.before) ?? 0;
  const continuesPiece =
    ownJoin > group.depth ||
    (edgeRecord !== undefined &&
      piece !== undefined &&
      sameOwnFields(edgeRecord, piece));
  if (!continuesPiece) return false;
  const standing = out.slice(start, start + records);
  const keptFrom = standing
    .flatMap((node) => keptCutFrom(node) ?? [])
    .toSorted()
    .at(0);
  const boundaries = standing.map((node) => withoutKeptCuts(node));
  // A piece still held in a pending removal, or content made after the
  // revision they were kept from, is no piece that takes them.
  const takes = (node: InlineNode | undefined): node is InlineNode =>
    node !== undefined &&
    !blocker(node) &&
    !isStandingBoundary(node) &&
    !(node.type === "deletion" || node.type === "moveFrom") &&
    !(
      isAddedRevision(node) &&
      keptFrom !== undefined &&
      (node.info.date ?? "") > keptFrom
    );
  const after = out.at(start + records);
  const before = start > 0 ? out.at(start - 1) : undefined;
  const intoAfter = takes(after)
    ? withBoundariesAtEdge(after, boundaries, group.depth, "start")
    : undefined;
  if (intoAfter !== undefined) {
    out.splice(start, records + 1, intoAfter);
    return true;
  }
  const intoBefore = takes(before)
    ? withBoundariesAtEdge(before, boundaries, group.depth, "end")
    : undefined;
  if (intoBefore === undefined) return false;
  out.splice(start - 1, records + 1, intoBefore);
  return true;
};

/** Update an explicit pending seam before removing any of its blocking revisions. */
const prepareDeferredRemovals = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): readonly InlineNode[] =>
  settleDeferredRemovals(nodes, {
    deferredOnlyRemovals: resolution.deferredOnlyRemovals,
    fateOf(id) {
      if (!resolution.ids.has(id)) return "pending";
      return resolution.decision === REVISION_DECISIONS.ACCEPT
        ? "accepted"
        : "removed";
    },
  });

type CutEdges = { first?: number; last?: number };
/** A handoff waiting for the record `levels` above the list to be emptied. */
type EnclosingHandoff = { levels: number; transfer: RetainedIdentity };

type ResolvedList = {
  nodes: InlineNode[];
  changed: boolean;
  edges?: CutEdges;
  /** Handoffs naming records above this list. */
  enclosing?: EnclosingHandoff[];
};

/** Comment ids whose complete range belongs to selected content being removed. */
const removedCommentRangeIds = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): ReadonlySet<number> => {
  const ids = new Set<number>();
  const visit = (
    items: readonly InlineNode[],
    removedWithAncestor = false,
  ): void => {
    for (const node of items) {
      const removedHere =
        removedWithAncestor ||
        (isTrackedWrapper(node) &&
          resolution.ids.has(node.info.id) &&
          removesContent(node, resolution.decision));
      if (isTrackedWrapper(node) && resolution.ids.has(node.info.id)) {
        if (removesContent(node, resolution.decision)) {
          const starts = new Set<number>();
          const ends = new Set<number>();
          const scan = (children: readonly InlineNode[]): void => {
            for (const child of children) {
              if (child.type === "commentRangeStart") starts.add(child.id);
              if (child.type === "commentRangeEnd") ends.add(child.id);
              scan(childNodes(child) ?? []);
            }
          };
          scan(node.content);
          for (const id of starts) if (ends.has(id)) ids.add(id);
        }
      }
      visit(childNodes(node) ?? [], removedHere);
    }
  };
  visit(nodes);
  return ids;
};

const withoutCommentAnchors = (
  nodes: readonly InlineNode[],
  removedIds: ReadonlySet<number>,
): readonly InlineNode[] => {
  if (removedIds.size === 0) return nodes;
  return nodes.flatMap((node) => {
    if (
      (node.type === "commentRangeStart" || node.type === "commentRangeEnd") &&
      removedIds.has(node.id)
    )
      return [];
    const children = childNodes(node);
    if (children === undefined) return [node];
    const filtered = withoutCommentAnchors(children, removedIds);
    return [filtered === children ? node : rebuildNode(node, filtered)];
  });
};

/**
 * A list with its tracked changes resolved. Where a change was resolved,
 * the records left meeting are merged as far as they are alike.
 */
/** What a resolution knows about `from`, now about `to`, its rebuilt copy. */
const carryResolutionState = (
  resolution: Resolution,
  from: InlineNode,
  to: InlineNode,
): void => {
  const edges = resolution.cutEdges.get(from);
  if (edges !== undefined) resolution.cutEdges.set(to, edges);
  if (resolution.emptied.has(from)) resolution.emptied.add(to);
  if (resolution.acceptedClosed.first.has(from))
    resolution.acceptedClosed.first.add(to);
  if (resolution.acceptedClosed.last.has(from))
    resolution.acceptedClosed.last.add(to);
};

type SourceOwnerSide = "before" | "after";

/**
 * A recorded join on one side identifies the surviving piece of this cut, as
 * does a removal cut that meets no piece on its other side (see
 * {@link withPayloadEdgesClosed}).
 */
const oneSidedKeptPiece = (
  joins: TrackedWrapper["resolutionJoins"],
): SourceOwnerSide | undefined => {
  if (joins === undefined) return undefined;
  if (joins.before === 0 && joins.after > 0) return "after";
  if (joins.after === 0 && joins.before > 0) return "before";
  const plain = joins.plainSides ?? [];
  if (joins.remove === 0 || plain.length !== 1) return undefined;
  return plain.includes("after") ? "before" : "after";
};

/**
 * The content an accepted insertion leaves, with a pending change at an edge
 * the insertion joins to nothing marked as meeting no piece there when its
 * removal cut records levels: the change was made inside the content, where
 * nothing stood beyond that edge, so its cut stops there. Removing it later
 * joins nothing across the edge and keeps its boundaries in the piece on its
 * other side, as removing it before the acceptance does.
 */
const withPayloadEdgesClosed = (
  content: readonly InlineNode[],
  joins: TrackedWrapper["resolutionJoins"],
  resolution: Resolution,
): readonly InlineNode[] => {
  if (joins === undefined) return content;
  const closed = (
    node: InlineNode | undefined,
    side: SourceOwnerSide,
  ): InlineNode | undefined => {
    if (
      node === undefined ||
      !isTrackedWrapper(node) ||
      resolution.ids.has(node.info.id) ||
      node.resolutionJoins === undefined ||
      node.resolutionJoins.remove === 0 ||
      node.resolutionJoins.plainSides?.includes(side) === true
    )
      return undefined;
    return Object.assign({}, node, {
      resolutionJoins: Object.assign({}, node.resolutionJoins, {
        plainSides: [...(node.resolutionJoins.plainSides ?? []), side],
      }),
    });
  };
  let out: InlineNode[] | undefined;
  const first =
    joins.before === 0 ? closed(content.at(0), "before") : undefined;
  if (first !== undefined) {
    out = [...content];
    out[0] = first;
  }
  const lastAt = content.length - 1;
  const last =
    joins.after === 0
      ? closed((out ?? content).at(lastAt), "after")
      : undefined;
  if (last !== undefined) {
    out ??= [...content];
    out[lastAt] = last;
  }
  return out ?? content;
};
type SourceOwnerRoute = { owner: SourceOwner; depth: number };
type SourceOwnerRoutes = Record<SourceOwnerSide, SourceOwnerRoute | undefined>;
type SourceOwnerState =
  | { type: "source" }
  | { type: "merged"; into: SourceOwner }
  | {
      type: "removed";
      routes: SourceOwnerRoutes;
      /** The occurrences either side as the list held them, never rerouted. */
      adjacent: Record<SourceOwnerSide, SourceOwner | undefined>;
    };
type SourceOwner = { state: SourceOwnerState };
/** A record in a resolved list and the source occurrence it came from. */
type SourceOccurrence = {
  node: InlineNode;
  owner: SourceOwner | undefined;
  /**
   * The record the list held before this resolution: for content a resolved
   * wrapper left in its place, that wrapper.
   */
  source?: InlineNode;
  /** Content of an added revision this resolution accepts: new, not source. */
  payload?: true;
};

/** `node` in the place of `occurrence`, from the same source. */
const occurrenceOf = (
  node: InlineNode,
  occurrence: SourceOccurrence | undefined,
): SourceOccurrence => ({
  node,
  owner: occurrence?.owner,
  ...(occurrence?.source !== undefined && { source: occurrence.source }),
  ...(occurrence?.payload === true && { payload: true as const }),
});

const sourceOwnerRoot = (owner: SourceOwner): SourceOwner => {
  let root = owner;
  while (root.state.type === "merged") {
    root = root.state.into;
  }
  let current = owner;
  while (current.state.type === "merged") {
    const next = current.state.into;
    current.state = { type: "merged", into: root };
    current = next;
  }
  return root;
};

const sourceOwnerDepthAt = ({
  owner,
  source,
  side,
  depth,
}: {
  owner: SourceOwner | undefined;
  source: SourceOwner | undefined;
  side: SourceOwnerSide;
  depth: number;
}): number | undefined => {
  if (owner === undefined || source === undefined) return undefined;
  const ownerRoot = sourceOwnerRoot(owner);
  const visited = new Set<SourceOwner>();
  let current: SourceOwner | undefined = source;
  const path: { owner: SourceOwner; depth: number }[] = [];
  const compress = (target: SourceOwner): number => {
    let routeDepth = Number.POSITIVE_INFINITY;
    for (const entry of path.toReversed()) {
      routeDepth = Math.min(routeDepth, entry.depth);
      if (entry.owner.state.type !== "removed")
        panic("Only removed source owners have directional routes.");
      entry.owner.state.routes[side] = { owner: target, depth: routeDepth };
    }
    return routeDepth;
  };
  while (current !== undefined) {
    if (visited.has(current))
      panic("Directional source-owner routes are acyclic.");
    if (sourceOwnerRoot(current) === ownerRoot) {
      return Math.min(depth, compress(current));
    }
    visited.add(current);
    if (current.state.type !== "removed") {
      compress(current);
      return undefined;
    }
    const route: SourceOwnerRoute | undefined = current.state.routes[side];
    if (route === undefined) {
      compress(current);
      return undefined;
    }
    // A removed occurrence between two others the resolution removes was cut
    // into content that goes with it; its depth measured that content, not
    // the content the route reaches, so it does not limit the route.
    const { adjacent } = current.state;
    const enclosed =
      adjacent.before?.state.type === "removed" &&
      adjacent.after?.state.type === "removed";
    path.push({
      owner: current,
      depth: enclosed ? Number.POSITIVE_INFINITY : route.depth,
    });
    current = route.owner;
  }
  return undefined;
};

/** Two tracked changes of one kind, author and date: pieces of one change. */
const sameStamp = (left: TrackedWrapper, right: TrackedWrapper): boolean =>
  left.type === right.type &&
  structurallyEqual({ ...left.info, id: 0 }, { ...right.info, id: 0 });

/** Every tracked change in `nodes`, at any depth. */
const trackedWithin = (nodes: readonly InlineNode[]): TrackedWrapper[] =>
  nodes.flatMap((node) => [
    ...(isTrackedWrapper(node) ? [node] : []),
    ...trackedWithin(childNodes(node) ?? []),
  ]);

type DeferredRemovalGroup = NonNullable<
  NonNullable<TrackedWrapper["resolutionJoins"]>["deferredRemove"]
>[number];

/**
 * The ids of the pending changes in `nodes`, at any depth, that can hold a
 * deferred seam: the only records a deferred seam may name as its blockers.
 * An insertion holds one it stands in; a deletion or move source holds one at
 * the edge where a piece it holds meets the seam.
 */
const seamHolderIds = (nodes: readonly InlineNode[]): Set<number> =>
  new Set(
    trackedWithin(nodes).flatMap((node) =>
      node.resolutionJoins !== undefined &&
      (isAddedRevision(node) ||
        node.type === "deletion" ||
        node.type === "moveFrom")
        ? [node.info.id]
        : [],
    ),
  );

/**
 * The ids a seam join retires, each with the id that takes its place: where
 * `left` and `right` met as `merged`, a tracked piece on the right's facing
 * edge merged into the left's piece at its level, whose id the merge kept.
 */
const rejoinedPieceIds = ({
  left,
  right,
  merged,
}: {
  left: InlineNode;
  right: InlineNode;
  merged: InlineNode;
}): [number, number][] => {
  const kept = new Set(trackedWithin([merged]).map((node) => node.info.id));
  const leftChain = edgeChain(left, "end");
  return edgeChain(right, "start").flatMap((piece, level) => {
    const into = leftChain.at(level);
    return isTrackedWrapper(piece) &&
      into !== undefined &&
      isTrackedWrapper(into) &&
      piece.info.id !== into.info.id &&
      !kept.has(piece.info.id) &&
      kept.has(into.info.id)
      ? [[piece.info.id, into.info.id]]
      : [];
  });
};

/**
 * `nodes` once seam joins retired the ids `retired` maps to the ids that took
 * their place: a deferred seam naming a retired piece names the piece it
 * merged into, as the validation of a later resolution reads blockers. A
 * piece merged into a record that can hold no seam blocks none.
 */
const withRejoinedBlockers = ({
  nodes,
  retired,
  resolution,
}: {
  nodes: readonly InlineNode[];
  retired: ReadonlyMap<number, number>;
  resolution: Resolution;
}): InlineNode[] => {
  const successor = (id: number): number => {
    let current = id;
    for (let next = retired.get(current); next !== undefined;) {
      current = next;
      next = retired.get(current);
    }
    return current;
  };
  const holders = seamHolderIds(nodes);
  const renamed = (
    groups: readonly DeferredRemovalGroup[],
  ): DeferredRemovalGroup[] =>
    groups.flatMap((group) => {
      const blockers = [
        ...new Set(
          group.blockers.flatMap((id) => {
            const into = successor(id);
            return into === id || holders.has(into) ? [into] : [];
          }),
        ),
      ];
      return blockers.length === 0 ? [] : [{ ...group, blockers }];
    });
  const rewrite = (node: InlineNode): InlineNode => {
    const children = childNodes(node);
    const rewritten = children?.map(rewrite);
    let next =
      rewritten === undefined ||
      rewritten.every((child, index) => child === children?.at(index))
        ? node
        : rebuildNode(node, rewritten);
    if (isTrackedWrapper(next) && next.resolutionJoins !== undefined) {
      const own = next.resolutionJoins.deferredRemove ?? [];
      const held = canonicalDeferredGroups(renamed(own));
      if (!structurallyEqual(held, own)) {
        const joins = Object.assign({}, next.resolutionJoins);
        delete joins.deferredRemove;
        if (held.length > 0) joins.deferredRemove = held;
        next = Object.assign({}, next, { resolutionJoins: joins });
      }
    }
    if (next !== node) carryResolutionState(resolution, node, next);
    return next;
  };
  return nodes.map(rewrite);
};

type CutReach = {
  /** The cut's levels that stay once the resolution is done. */
  levels: number;
  /** Records each side's edge goes through to the deepest level that stays. */
  before: number;
  after: number;
};

/**
 * Two records meeting at a cut recorded `depth` levels deep, folded as
 * {@link foldAtSeam} folds an emptied piece into the piece across from it.
 * The cut split the containers on its facing edges down to that depth, so
 * one of them holding nothing is a piece a rejection emptied: this
 * resolution (one `emptied` holds) or an earlier one, which left it beside
 * the change that cut it. Read from the pieces themselves rather than from
 * the resolution that emptied them, the fold is the same whichever pass
 * emptied the piece, so rejecting a run's revisions at once and in any order
 * leave the same records.
 */
const foldAtCut = ({
  left,
  right,
  depth,
  emptied,
}: {
  left: InlineNode;
  right: InlineNode;
  depth: number;
  emptied: Emptied;
}): InlineNode[] => {
  const pieces = new Set<InlineNode>();
  const before = edgeChain(left, "end");
  const after = edgeChain(right, "start");
  for (let level = 0; level < depth; level += 1) {
    const near = before.at(level);
    const far = after.at(level);
    if (near === undefined || far === undefined || near.type === "run") break;
    const nearChildren = childNodes(near);
    const farChildren = childNodes(far);
    if (
      nearChildren === undefined ||
      farChildren === undefined ||
      !sameOwnFields(near, far)
    )
      break;
    if (nearChildren.length === 0) pieces.add(near);
    if (farChildren.length === 0) pieces.add(far);
  }
  return foldAtSeam(left, right, {
    has: (node) => emptied.has(node) || pieces.has(node),
  });
};

/**
 * Where a cut recorded `depth` levels deep reaches, once `resolution` is
 * done, into the records either side of it, given as the chains of records
 * at their edges facing the cut, as the list held them before the
 * resolution. The levels a cut counts are the containers it split: their
 * pieces stand on both sides of it and are alike, as the pieces a join
 * merges are, a tracked change's pieces sharing its kind, author and date.
 * The reach ends where the two sides differ. A tracked change on one side
 * only came
 * after the cut, or never held it, and is no level of it; nor is one that
 * also holds a piece of `spanning`, the content that made the cut, as a
 * deletion made over the cut does. Such a change this resolution removes
 * with its content is read past: the record beside it inside the same
 * container meets the cut once it is gone. So the reach is the same whichever
 * changes earlier resolutions removed and whichever later edits wrapped a
 * piece: a cut level the resolution resolves away leaves its content a level
 * up, and one whose content goes takes the position with it.
 */
const cutReach = ({
  before,
  after,
  depth,
  spanning,
  resolution,
}: {
  before: readonly InlineNode[];
  after: readonly InlineNode[];
  depth: number;
  spanning: readonly TrackedWrapper[];
  resolution: Resolution;
}): CutReach => {
  const reach: CutReach = { levels: depth, before: 0, after: 0 };
  const open = { before: true, after: true };
  const chains = { before: [...before], after: [...after] };
  /**
   * Reads one side's edge past a tracked change there that no record across
   * from it matches and this resolution removes with its content: the
   * record it stood in front of is what then meets the cut. Whether
   * resolved at once or before the cut's own change, the cut reaches the
   * same records.
   */
  const pastRemoved = (side: "before" | "after", index: number): boolean => {
    const node = chains[side][index];
    if (
      node === undefined ||
      !isTrackedWrapper(node) ||
      !resolution.ids.has(node.info.id) ||
      !removesContent(node, resolution.decision)
    )
      return false;
    const chain = edgeChainPastRemoved(
      chains[side],
      index,
      side === "before" ? "end" : "start",
    );
    if (chain === undefined) return false;
    chains[side] = chain;
    return true;
  };
  /** How a record on one side's edge counts once the resolution is done. */
  const pass = (
    side: "before" | "after",
    node: InlineNode,
  ): "stays" | "lifted" | "gone" => {
    if (isTrackedWrapper(node) && resolution.ids.has(node.info.id)) {
      if (!removesContent(node, resolution.decision)) return "lifted";
      open[side] = false;
      return "gone";
    }
    if (open[side]) reach[side] += 1;
    return "stays";
  };
  let nearIndex = 0;
  let farIndex = 0;
  let counted = 0;
  let staying = 0;
  while (counted < depth) {
    const near = chains.before[nearIndex];
    const far = chains.after[farIndex];
    if (near === undefined || far === undefined) break;
    const piecesOfOne =
      isTrackedWrapper(near) && isTrackedWrapper(far)
        ? sameStamp(near, far) &&
          !spanning.some((change) => sameStamp(change, near))
        : !isTrackedWrapper(near) &&
          !isTrackedWrapper(far) &&
          childNodes(near) !== undefined &&
          childNodes(far) !== undefined &&
          sameOwnFields(near, far);
    if (!piecesOfOne && isTrackedWrapper(near)) {
      if (pastRemoved("before", nearIndex)) continue;
      pass("before", near);
      nearIndex += 1;
      continue;
    }
    if (!piecesOfOne && isTrackedWrapper(far)) {
      if (pastRemoved("after", farIndex)) continue;
      pass("after", far);
      farIndex += 1;
      continue;
    }
    if (!piecesOfOne) break;
    counted += 1;
    const outcome = pass("before", near);
    pass("after", far);
    if (outcome === "gone") {
      reach.levels = staying;
      break;
    }
    if (outcome === "lifted") reach.levels -= 1;
    else staying += 1;
    nearIndex += 1;
    farIndex += 1;
  }
  return reach;
};

type ResolveListOptions = {
  nodes: readonly InlineNode[];
  resolution: Resolution;
  path: readonly number[];
};

/** Joins that record only where a piece meets its neighbours. */
const PLAIN_JOIN_FIELDS: ReadonlySet<string> = new Set([
  "before",
  "after",
  "remove",
]);

/**
 * `nodes` with the pieces of one rejected insertion met again where only
 * kept markers stand between them: two pieces of one change, each joined to
 * the other, that this resolution removes. The markers stand in the record
 * the pieces form (they never go inside a pending change, so they waited
 * beside its pieces); removing the pieces as one keeps them as that
 * record's own, so they land where rejecting it before the cut between its
 * pieces lands them.
 */
const withRemovedPiecesMet = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): readonly InlineNode[] => {
  if (resolution.decision !== REVISION_DECISIONS.REJECT) return nodes;
  const removedPiece = (node: InlineNode | undefined): node is TrackedWrapper =>
    node !== undefined &&
    isTrackedWrapper(node) &&
    isAddedRevision(node) &&
    resolution.ids.has(node.info.id) &&
    node.resolutionJoins !== undefined &&
    node.resolutionJoins.remove === 0 &&
    Object.keys(node.resolutionJoins).every((key) =>
      PLAIN_JOIN_FIELDS.has(key),
    ) &&
    // Positions inside it stay as recorded: it holds no other change.
    trackedWithin(node.content).length === 0;
  let out: InlineNode[] | undefined;
  let index = 0;
  const source = (): readonly InlineNode[] => out ?? nodes;
  while (index < source().length) {
    const first = source()[index];
    if (!removedPiece(first) || (first.resolutionJoins?.after ?? 0) === 0) {
      index += 1;
      continue;
    }
    const keptAt = (at: number): boolean => {
      const node = source()[at];
      return node !== undefined && standsInKeptCut(node);
    };
    let end = index + 1;
    while (keptAt(end)) end += 1;
    const second = source()[end];
    if (
      end === index + 1 ||
      !removedPiece(second) ||
      (second.resolutionJoins?.before ?? 0) === 0 ||
      !sameStamp(first, second)
    ) {
      index += 1;
      continue;
    }
    // The pieces' records meet as deep as each joins the other, the markers
    // inside the record they form.
    const depth =
      Math.min(
        first.resolutionJoins?.after ?? 0,
        second.resolutionJoins?.before ?? 0,
      ) - 1;
    const seam = withSeamBoundariesInside({
      left: first.content,
      right: [...source().slice(index + 1, end), ...second.content],
      depth,
    });
    const content =
      mergeLists(seam.left, seam.right, depth, { mode: "asFarAsAlike" }) ??
      panic("An as-far-as-alike merge always returns its records.");
    const met = Object.assign({}, rebuildNode(first, content), {
      resolutionJoins: {
        before: first.resolutionJoins?.before ?? 0,
        after: second.resolutionJoins?.after ?? 0,
        remove: 0,
      },
    });
    carryResolutionState(resolution, first, met);
    out ??= [...nodes];
    out.splice(index, end - index + 1, met);
  }
  return out ?? nodes;
};

const resolveList = ({
  nodes: listed,
  resolution,
  path,
}: ResolveListOptions): ResolvedList => {
  const nodes = withRemovedPiecesMet(listed, resolution);
  const out: SourceOccurrence[] = [];
  const seams: number[] = [];
  const plainSeams = new Set<number>();
  const exactSeams = new Map<number, number>();
  const sourceEdges = new Map<number, number>();
  const pending = [...nodes];
  const sourceOwners = nodes.map((): SourceOwner => ({
    state: { type: "source" },
  }));
  const appendSource = (
    node: InlineNode,
    index: number,
    payload = false,
  ): void => {
    const source = pending.at(index);
    out.push({
      node,
      owner:
        sourceOwners.at(index) ?? panic("A source occurrence has an owner."),
      ...(source !== undefined && { source }),
      ...(payload && { payload: true as const }),
    });
  };
  /**
   * How many levels of the cut `node` recorded for its removal remain once
   * this resolution is done, read against the records either side of it as
   * the list held them: a container it split that this resolution resolves
   * away is a level of it no more, and one whose content goes takes the cut
   * with it. The one reading of a recorded removal cut, for a change this
   * resolution removes and for one it leaves pending alike, so the cut joins
   * the same records whichever resolution removes the pieces it split.
   */
  const cutLevelsOnceResolved = (node: InlineNode, index: number): number => {
    if (!isTrackedWrapper(node)) return 0;
    const joins = node.resolutionJoins;
    const depth = joins?.remove ?? 0;
    if (depth === 0) return 0;
    // A change this resolution removes that holds no record either side
    // (its joins on keeping are 0) cut only the records beyond the changes
    // later edits put beside it; those go with this resolution too, so the
    // cut is read against what then meets at its seam.
    const reachesPast =
      resolution.ids.has(node.info.id) &&
      removesContent(node, resolution.decision) &&
      joins?.before === 0 &&
      joins.after === 0;
    const goes = (candidate: InlineNode | undefined): boolean =>
      reachesPast &&
      candidate !== undefined &&
      isTrackedWrapper(candidate) &&
      resolution.ids.has(candidate.info.id) &&
      removesContent(candidate, resolution.decision);
    let beforeIndex = index - 1;
    while (beforeIndex >= 0 && goes(pending.at(beforeIndex))) beforeIndex -= 1;
    let afterIndex = index + 1;
    while (afterIndex < pending.length && goes(pending.at(afterIndex)))
      afterIndex += 1;
    const before = beforeIndex >= 0 ? pending.at(beforeIndex) : undefined;
    const after = pending.at(afterIndex);
    // Every piece on one side goes with this resolution: no record of the
    // level it cut is left there to join.
    if (
      (before === undefined && beforeIndex < index - 1) ||
      (after === undefined && afterIndex > index + 1)
    )
      return 0;
    if (before === undefined || after === undefined) return depth;
    return cutReach({
      before: edgeChain(before, "end"),
      after: edgeChain(after, "start"),
      depth,
      spanning: trackedWithin(node.content),
      resolution,
    }).levels;
  };
  /**
   * The cut a pending insertion recorded, once this resolution is done (see
   * {@link cutLevelsOnceResolved}). Kept as it was, a later resolution of the
   * insertion would read the cut one level deeper than resolving both at once
   * does. `undefined` when the cut stays as recorded.
   */
  const remainingCut = (
    node: InlineNode,
    index: number,
  ): number | undefined => {
    if (!isAddedRevision(node) || resolution.ids.has(node.info.id))
      return undefined;
    const levels = cutLevelsOnceResolved(node, index);
    return levels < (node.resolutionJoins?.remove ?? 0) ? levels : undefined;
  };
  /**
   * The depth of the cut of the record a removed change at `index` was made
   * inside, when this resolution removes that record's pieces either side
   * of it too: the boundaries the change kept stand inside the record, and
   * go with its cut, as rejecting the change first and the record after
   * places them. `undefined` when no such record encloses it.
   */
  const enclosingRecordCut = (
    node: TrackedWrapper,
    index: number,
  ): number | undefined => {
    const before = index > 0 ? pending.at(index - 1) : undefined;
    const after = pending.at(index + 1);
    if (
      before === undefined ||
      after === undefined ||
      !isTrackedWrapper(before) ||
      !isTrackedWrapper(after) ||
      !sameStamp(before, after) ||
      ![before, after].every(
        (piece) =>
          resolution.ids.has(piece.info.id) &&
          removesContent(piece, resolution.decision),
      )
    )
      return undefined;
    if ((before.info.date ?? "") >= (node.info.date ?? "")) return undefined;
    const depth = Math.min(
      before.resolutionJoins?.remove ?? 0,
      after.resolutionJoins?.remove ?? 0,
    );
    return depth > 0 ? depth : undefined;
  };
  const appendUnowned = (items: readonly InlineNode[]): void => {
    out.push(...items.map((node) => ({ node, owner: undefined })));
  };
  /**
   * The kept markers `out` holds from `start` to `end` naming `piece` as the
   * side whose record is a piece of their cut: a pending change there took
   * their seam. A marker already naming a piece keeps it: the first change
   * to take its seam holds it. Whether any marker took the name.
   */
  const nameKeptPiece = (
    start: number,
    end: number,
    piece: KeptCutPiece,
  ): boolean => {
    let named = false;
    for (let index = start; index < end; index += 1) {
      const occurrence = out[index];
      if (occurrence === undefined) continue;
      const next = withKeptCutPieceRecorded(
        occurrence.node,
        piece,
        resolution.relocated.nodes,
      );
      if (next === occurrence.node) continue;
      out[index] = occurrenceOf(next, occurrence);
      named = true;
    }
    return named;
  };
  type PlaceOneSidedKeptOptions = {
    /** The removed insertion's place in `pending`. */
    index: number;
    piece: "before" | "after";
    depth: number;
    /** How many boundaries it kept, the last ones `out` holds. */
    kept: number;
    /** How many of them, nearest the piece, belong inside it. */
    inside: number;
    /** The removed content's record at the piece's edge. */
    edge: InlineNode | undefined;
  };
  /**
   * The boundaries a removed one-sided insertion kept, moved into the source
   * piece its join names: the record after it, not yet resolved, or the one
   * `out` holds before them. They stay standing where that record is a
   * tracked change or no piece of the record the edge one is. Their cuts go
   * with them, naming no piece: settling the list keeps a cut only while a
   * pending change still stands beside them, as placing them by accepting
   * that piece later does.
   */
  const placeOneSidedKept = ({
    index,
    piece,
    depth,
    kept,
    inside,
    edge,
  }: PlaceOneSidedKeptOptions): void => {
    const standsInside = (record: InlineNode | undefined): boolean =>
      record !== undefined &&
      edge !== undefined &&
      !isTrackedWrapper(record) &&
      childNodes(record) !== undefined &&
      sameOwnFields(edge, record);
    const first = out.length - kept;
    const boundaries = out
      .slice(
        piece === "after" ? out.length - inside : first,
        piece === "after" ? out.length : first + inside,
      )
      .map(({ node }) =>
        withKeptCutPieceConsumed(node, resolution.relocated.nodes),
      );
    if (piece === "after") {
      const next = pending.at(index + 1);
      if (next === undefined || !standsInside(next)) return;
      const placed = withBoundariesAtEdge(next, boundaries, depth, "start");
      if (placed === undefined) return;
      pending[index + 1] = placed;
      out.splice(out.length - inside, inside);
    } else {
      const previous = out.at(first - 1);
      if (previous === undefined || !standsInside(previous.node)) return;
      const placed = withBoundariesAtEdge(
        previous.node,
        boundaries,
        depth,
        "end",
      );
      if (placed === undefined) return;
      carryResolutionState(resolution, previous.node, placed);
      out[first - 1] = occurrenceOf(placed, previous);
      out.splice(first, inside);
    }
  };
  /**
   * Boundaries a rejection kept from new content joined to nothing, standing
   * beside a pending insertion whose join there cuts a record the content
   * holds or was made in: that join takes them, as of the revision they were
   * kept from. One made before the content cut the record the content was
   * made in, which holds every boundary the content kept; one made later,
   * inside the content's record at that edge, cut only that record, which
   * holds the boundaries kept from inside it, nearest the seam. They name
   * its side as their cut's piece, as a seam meeting such an insertion does
   * (see {@link handJoin}), so accepting the insertion later places them
   * inside that record, as accepting it before this rejection does.
   */
  const handKeptToFacingJoins = ({
    index,
    kept,
    standing,
    inEdgeRecords,
  }: {
    /** The removed content's place in `pending`. */
    index: number;
    /**
     * The date of the revision they were kept from, and how many boundaries
     * they stand for.
     */
    kept: KeptAtSeam;
    /** How many records the kept boundaries are, the last ones `out` holds. */
    standing: number;
    /**
     * How many of the kept boundaries nearest each edge of the removed
     * content its record at that edge held.
     */
    inEdgeRecords: Record<"first" | "last", number>;
  }): void => {
    /** The record a list's edge record holds at `edge`, past tracked changes. */
    const recordAt = (
      node: InlineNode | undefined,
      edge: "start" | "end",
    ): InlineNode | undefined => {
      let record = node;
      while (record !== undefined && isTrackedWrapper(record))
        record =
          edge === "start" ? record.content.at(0) : record.content.at(-1);
      return record;
    };
    const previousAt = out.length - standing - 1;
    const facing = (
      node: InlineNode | undefined,
      side: "before" | "after",
    ): { join: number; kept: KeptAtSeam } | undefined => {
      if (
        node === undefined ||
        !isTrackedWrapper(node) ||
        !isAddedRevision(node) ||
        resolution.ids.has(node.info.id)
      )
        return undefined;
      const join = node.resolutionJoins?.[side] ?? 0;
      const date = node.info.date ?? "";
      if (join === 0) return undefined;
      if (date < kept.from) return { join, kept };
      const count = inEdgeRecords[side === "before" ? "last" : "first"];
      // A piece of that record across the boundaries meets the insertion's
      // join around them: the seam places them, not the insertion alone. A
      // pending insertion there, joined to nothing, holds no such piece.
      const own = recordAt(node, side === "before" ? "start" : "end");
      const acrossNode =
        side === "before" ? out[previousAt]?.node : pending[index + 1];
      const across = recordAt(acrossNode, side === "before" ? "end" : "start");
      const pieceAcross =
        own !== undefined &&
        across !== undefined &&
        !(
          acrossNode !== undefined &&
          isAddedRevision(acrossNode) &&
          !resolution.ids.has(acrossNode.info.id) &&
          acrossNode.resolutionJoins === undefined
        ) &&
        childNodes(own) !== undefined &&
        childNodes(across) !== undefined &&
        sameOwnFields(own, across);
      return date > kept.from && count > 0 && !pieceAcross
        ? { join, kept: { from: kept.from, count } }
        : undefined;
    };
    /** How many records the `count` kept boundaries nearest an edge are. */
    const recordsFrom = (from: number, step: 1 | -1, count: number): number =>
      keptRun({ at: (at) => out[at]?.node, from, step, limit: count }).records;
    const next = pending[index + 1];
    const nextFacing = facing(next, "before");
    if (
      next !== undefined &&
      nextFacing !== undefined &&
      handJoin(next, "before", nextFacing.join).takesKept
    )
      nameKeptPiece(
        out.length - recordsFrom(out.length - 1, -1, nextFacing.kept.count),
        out.length,
        "after",
      );
    const previous = out[previousAt];
    const previousFacing = facing(previous?.node, "after");
    if (
      previous !== undefined &&
      previousFacing !== undefined &&
      handJoin(previous.node, "after", previousFacing.join).takesKept
    )
      nameKeptPiece(
        previousAt + 1,
        previousAt +
          1 +
          recordsFrom(previousAt + 1, 1, previousFacing.kept.count),
        "before",
      );
  };
  /**
   * Rejecting new content made inside a piece of an earlier pending
   * insertion (joined to it on that side) cancels that insertion's join
   * facing it. Accepting the insertion first leaves its join nothing to
   * take once the content goes; rejecting the content first must not leave
   * the join to reach past its place, so the result does not depend on the
   * order.
   */
  const retireFacingJoins = (index: number, node: TrackedWrapper): void => {
    const joins = node.resolutionJoins;
    if (joins === undefined) return;
    const facesNode = (
      neighbour: InlineNode | undefined,
      side: "before" | "after",
    ): neighbour is TrackedWrapper =>
      neighbour !== undefined &&
      isTrackedWrapper(neighbour) &&
      isAddedRevision(neighbour) &&
      !resolution.ids.has(neighbour.info.id) &&
      (neighbour.info.date ?? "") < (node.info.date ?? "") &&
      (neighbour.resolutionJoins?.[side] ?? 0) > 0;
    const previous = out.at(-1);
    if (joins.before > 0 && facesNode(previous?.node, "after")) {
      const retired = withJoinRetired(previous.node, "after");
      carryResolutionState(resolution, previous.node, retired);
      out[out.length - 1] = occurrenceOf(retired, previous);
    }
    const next = pending[index + 1];
    if (joins.after > 0 && facesNode(next, "before"))
      pending[index + 1] = withJoinRetired(next, "before");
  };
  const recordExactSeam = (index: number, depth: number): void => {
    if (depth > 0)
      exactSeams.set(index, Math.max(exactSeams.get(index) ?? 0, depth));
  };
  /** Seams a removed revision recorded between the source pieces around it. */
  const removalSeams = new Set<number>();
  /** Seams a kept revision's own joins recorded at its edges. */
  const keptJoinSeams = new Set<number>();
  /** The date of the earliest kept revision whose join recorded each seam. */
  const keptJoinDates = new Map<number, string>();
  const recordKeptJoin = (seam: number, node: TrackedWrapper): void => {
    keptJoinSeams.add(seam);
    const date = node.info.date ?? "";
    const known = keptJoinDates.get(seam);
    if (known === undefined || date < known) keptJoinDates.set(seam, date);
  };
  /**
   * Of those a rejected deletion or move source recorded, the seams with its
   * own content just before them (its join after) or just after them (its
   * join before). That content is authored, not boundaries an earlier
   * rejection kept at the seam. (An accepted insertion's join can be one a
   * seam handed over to it, which its content stands inside.)
   */
  const ownContentBefore = new Set<number>();
  const ownContentAfter = new Set<number>();
  /** Those of them a removed insertion or move destination recorded. */
  const addedRemovalSeams = new Set<number>();
  const recordSourceSeam = (index: number, depth: number): void => {
    recordExactSeam(index, depth);
    if (depth > 0)
      sourceEdges.set(index, Math.max(sourceEdges.get(index) ?? 0, depth));
  };
  /** Seams beside an emptied container, where only its fold happens. */
  const folds: number[] = [];
  /**
   * The earliest revision the kept markers `out` holds from `start` to `end`
   * were kept from, as their cuts record it.
   */
  const keptFromIn = (start: number, end: number): string | undefined =>
    out
      .slice(Math.max(0, start), Math.max(0, end))
      .flatMap(({ node }) => keptCutFrom(node) ?? [])
      .toSorted()
      .at(0);
  /** The revision a kept marker `out` holds at `index` was kept from. */
  const keptFromAt = (index: number): string | undefined =>
    index < 0 || !isStandingBoundary(out[index]?.node)
      ? undefined
      : keptFromIn(index, index + 1);
  /**
   * Boundaries an earlier rejection kept standing at a seam this resolution
   * joins, by the seam's position: the date of the revision they were kept
   * from, how many of those standing there they are, and, where a kept
   * revision's own join takes the seam, the side of it that revision's
   * content is on, a piece of the record the rejection cut.
   */
  const exactKept = new Map<
    number,
    { from: string; count?: number; piece?: "before" | "after" }
  >();
  /**
   * The kept markers standing in a row from `from`, stepping by `step`, that
   * name `piece` as the side taking them: the earliest revision they were
   * kept from and how many boundaries they stand for.
   */
  const keptByPiece = (
    at: (index: number) => InlineNode | undefined,
    from: number,
    step: 1 | -1,
    piece: KeptCutPiece,
  ): { from: string; count: number } | undefined => {
    const dates: string[] = [];
    let count = 0;
    for (let index = from; ; index += step) {
      const node = at(index);
      if (
        node === undefined ||
        !standsInKeptCut(node) ||
        keptCutPiece(node) !== piece
      )
        break;
      const date = keptCutFrom(node);
      if (date !== undefined) dates.push(date);
      count += keptBoundaryCount([node]);
    }
    const earliest = dates.toSorted().at(0);
    return earliest === undefined ? undefined : { from: earliest, count };
  };
  /**
   * The kept markers standing beside the removed change at `index` that name
   * its side, or both sides, as their cut's piece: those `out` holds before
   * `before` and those after it in `pending`. Its removal consumes the name;
   * the earliest revision they were kept from and how many boundaries they
   * stand for.
   */
  const takeKeptBeside = (
    index: number,
    before: number,
  ): { from: string; count: number } | undefined => {
    const dates: string[] = [];
    let count = 0;
    const take = (
      at: (position: number) => InlineNode | undefined,
      from: number,
      step: 1 | -1,
      piece: KeptCutPiece,
      put: (position: number, node: InlineNode) => void,
    ): void => {
      for (let position = from; ; position += step) {
        const node = at(position);
        if (node === undefined || !standsInKeptCut(node)) return;
        const named = keptCutPiece(node);
        if (named !== piece && named !== "both") return;
        const date = keptCutFrom(node);
        if (date !== undefined) dates.push(date);
        count += keptBoundaryCount([node]);
        put(
          position,
          withKeptCutPieceConsumed(node, resolution.relocated.nodes),
        );
      }
    };
    take(
      (position) => out[position]?.node,
      before - 1,
      -1,
      "after",
      (position, node) => {
        const occurrence = out[position];
        if (occurrence !== undefined)
          out[position] = occurrenceOf(node, occurrence);
      },
    );
    take(
      (position) => pending[position],
      index + 1,
      1,
      "before",
      (position, node) => {
        pending[position] = node;
      },
    );
    const earliest = dates.toSorted().at(0);
    return earliest === undefined ? undefined : { from: earliest, count };
  };
  /**
   * Boundaries kept from removed content just before a seam, and its depth.
   * They float as of the earliest revision their cuts record they were kept
   * from.
   */
  const floating = new Map<
    number,
    {
      count: number;
      depth: number;
      before: SourceOwner | undefined;
      after: SourceOwner | undefined;
      /** The tracked changes the removed content held. */
      spanning: readonly TrackedWrapper[];
      /**
       * The side whose record is a piece of the record the removed content
       * cut by its own join: it takes them across from any record.
       */
      piece?: "before" | "after";
      /**
       * Set where the removal's cut came only from a seam deferred on it:
       * that seam joins the pieces either side only where they are alike,
       * and only then do they take the boundaries, as resolving the
       * deferring change after the removal leaves them standing.
       */
      cut?: "deferred";
    }
  >();
  let changed = false;
  const enclosing: EnclosingHandoff[] = [];
  for (const [index, node] of pending.entries()) {
    if (isTrackedWrapper(node) && resolution.ids.has(node.info.id)) {
      changed = true;
      if (removesContent(node, resolution.decision)) {
        const removedOwner =
          sourceOwners.at(index) ??
          panic("A removed source occurrence has an owner.");
        const joins = node.resolutionJoins;
        if (
          resolution.decision === REVISION_DECISIONS.REJECT &&
          isAddedRevision(node)
        )
          retireFacingJoins(index, node);
        const removedDepth = joins?.remove ?? 0;
        // A one-sided prefix or suffix still belongs to the piece its join
        // names, even when removing it joins no records across its place.
        let beforeDepth = removedDepth;
        let afterDepth = removedDepth;
        if (removedDepth === 0) {
          beforeDepth = joins?.after === 0 ? joins.before : 0;
          afterDepth = joins?.before === 0 ? joins.after : 0;
        }
        if (
          resolution.decision === REVISION_DECISIONS.REJECT &&
          isAddedRevision(node) &&
          (beforeDepth > 0 || afterDepth > 0)
        ) {
          const before =
            index === 0
              ? undefined
              : (sourceOwners.at(index - 1) ??
                panic("A preceding source occurrence has an owner."));
          const after = sourceOwners.at(index + 1);
          removedOwner.state = {
            type: "removed",
            adjacent: { before, after },
            routes: {
              before:
                before === undefined || beforeDepth === 0
                  ? undefined
                  : { owner: before, depth: beforeDepth },
              after:
                after === undefined || afterDepth === 0
                  ? undefined
                  : { owner: after, depth: afterDepth },
            },
          };
        } else if (
          resolution.decision === REVISION_DECISIONS.REJECT &&
          isAddedRevision(node) &&
          joins !== undefined &&
          joins.before === 0 &&
          joins.after === 0 &&
          joins.remove === 0
        ) {
          // New content joined to nothing either side is no piece of a
          // record: removing it, the records either side meet across its
          // place, so a seam beyond it reaches them as deep as it goes.
          const before = index === 0 ? undefined : sourceOwners.at(index - 1);
          const after = sourceOwners.at(index + 1);
          const through = (owner: SourceOwner | undefined) =>
            owner === undefined
              ? undefined
              : { owner, depth: Number.POSITIVE_INFINITY };
          removedOwner.state = {
            type: "removed",
            adjacent: { before, after },
            routes: { before: through(before), after: through(after) },
          };
        }
        const transfers = node.resolutionJoins?.retainedAfter ?? [];
        if (transfers.length > 0) {
          // Restore before resolving later wrappers, so their own source
          // references follow the fragment whose identity returned.
          const prefixLength = out.length;
          const restored = restoreRetainedIdentities({
            nodes: out
              .map(({ node: previous }) => previous)
              .concat(pending.slice(index + 1)),
            transfers,
            resolution,
          });
          resolution.outstanding.push(...restored.unmatched);
          // Restoring source IDs can make a previously unselected blocker
          // selected; recompute deferred facts against those actual identities.
          const prepared = prepareDeferredRemovals(restored.nodes, resolution);
          out.splice(
            0,
            out.length,
            ...prepared.slice(0, prefixLength).map((next, prefixIndex) => {
              const previous =
                out.at(prefixIndex) ??
                panic(
                  "Restoring source identities preserves every output occurrence.",
                );
              const restoredOccurrence: SourceOccurrence = {
                node: next,
                owner: previous.owner,
              };
              if (previous.source !== undefined)
                restoredOccurrence.source = previous.source;
              if (previous.payload === true) restoredOccurrence.payload = true;
              return restoredOccurrence;
            }),
          );
          pending.splice(
            index + 1,
            pending.length - index - 1,
            ...prepared.slice(prefixLength),
          );
        }
        // Removing a prefix that was all a cut record held hands the record's
        // identities to the piece cut off after it.
        for (const entry of node.resolutionJoins?.retainedEnclosing ?? [])
          enclosing.push({
            levels: entry.levels,
            transfer: { depth: 0, source: entry.source, target: entry.target },
          });
        const removalContent =
          resolution.decision === REVISION_DECISIONS.REJECT &&
          isAddedRevision(node)
            ? rejectCopyOriginAncestorsInNodes(node.content, 1)
            : node.content;
        const keptOf = (removed: readonly InlineNode[]) =>
          removedContentBoundaries({
            nodes: removed,
            rejecting: resolution.decision === REVISION_DECISIONS.REJECT,
            // Endpoints follow keptPaths; a range deletion removes an enclosed
            // comment reference as deleting that range directly does.
            pairedComments:
              node.type === "deletion" && node.resolutionJoins !== undefined
                ? NO_COMMENT_RANGES
                : resolution.pairedComments,
            relocated: resolution.relocated,
            revisionDate: node.info.date ?? "",
            leavesPending: (revisionId) => !resolution.ids.has(revisionId),
            keptPaths: resolution.keptDeletionPaths,
            path: [...path, index],
            keepAutomaticMark: (run) =>
              automaticNoteMarkRuns([run]).map((mark) =>
                resolveRun(mark, resolution),
              ),
          });
        // Each boundary kept from the removed content records that it stands
        // in the cut the removal made, as of the removed revision: a later
        // join takes it, where an authored or accepted one blocks the join.
        // Placing it inside a piece clears the record unless a pending change
        // meets it there, as does a resolution that leaves none beside it.
        const cut = {
          depth: Math.max(1, joins?.remove ?? 0),
          keptFrom: node.info.date ?? "",
        };
        const kept = keptOf(removalContent).map((boundary) =>
          isKeptBoundaryCarrier(boundary)
            ? withKeptCutsRecorded(boundary, cut, resolution.relocated.nodes)
            : boundary,
        );
        const keptStart = out.length;
        appendUnowned(kept);
        // Kept boundaries float at the seam until the pieces either side are
        // resolved; the join then places them inside the piece it joins.
        // With the cut on one side only, the boundaries that go into that
        // piece are those before the removed content's first record past it,
        // that record's own included: one beyond that record stood outside
        // the record the piece continues.
        // A cut deferred by another edit can join the surrounding content.
        // It cannot give this removal's own boundaries solely to one side.
        const piece =
          resolution.decision === REVISION_DECISIONS.REJECT &&
          joins !== undefined &&
          !resolution.deferredOnlyRemovals.has(joins)
            ? oneSidedKeptPiece(joins)
            : undefined;
        const records = removalContent.flatMap((child, at) =>
          childNodes(child) === undefined ? [] : [at],
        );
        const firstRecord = records.at(0);
        const lastRecord = records.at(-1);
        let inside = kept.length;
        if (piece === "before" && firstRecord !== undefined)
          inside = keptOf(removalContent.slice(0, firstRecord + 1)).length;
        else if (piece === "after" && lastRecord !== undefined)
          inside =
            kept.length - keptOf(removalContent.slice(0, lastRecord)).length;
        if (
          inside > 0 &&
          resolution.decision === REVISION_DECISIONS.REJECT &&
          (node.resolutionJoins?.remove ?? 0) > 0
        ) {
          // Those it held stand first when the piece is before, last when after.
          const at =
            piece === "before"
              ? out.length - (kept.length - inside)
              : out.length;
          const own = node.resolutionJoins?.remove ?? 0;
          const enclosingCut = enclosingRecordCut(node, index);
          // A side meeting no piece of the cut has no record to take them.
          const closedOn = (side: SourceOwnerSide): boolean =>
            node.resolutionJoins?.plainSides?.includes(side) === true;
          floating.set(at, {
            count: inside,
            depth: Math.max(own, enclosingCut ?? 0),
            before:
              index > 0 && !closedOn("before")
                ? sourceOwners.at(index - 1)
                : undefined,
            after: closedOn("after") ? undefined : sourceOwners.at(index + 1),
            spanning: trackedWithin(node.content),
            // Accepting content beside this removed piece can leave its
            // recorded container join on one side only. The other side's
            // authored record is not part of that cut container.
            ...(piece === undefined ? {} : { piece }),
            ...(joins !== undefined &&
              resolution.deferredOnlyRemovals.has(joins) && {
                cut: "deferred" as const,
              }),
          });
        }
        // A rejected insertion joined on one side only, made inside a piece of
        // the record there, cuts nothing across its place: the boundaries it
        // kept inside that record go into the piece beside it now, as deep as
        // its join, as they do when its content goes with a cut it holds.
        if (
          inside > 0 &&
          piece !== undefined &&
          (joins?.remove ?? 0) === 0 &&
          resolution.decision === REVISION_DECISIONS.REJECT &&
          isAddedRevision(node)
        )
          placeOneSidedKept({
            index,
            piece,
            depth:
              piece === "after" ? (joins?.after ?? 0) : (joins?.before ?? 0),
            kept: kept.length,
            inside,
            edge: removalContent[
              (piece === "after" ? lastRecord : firstRecord) ?? -1
            ],
          });
        if (
          kept.length > 0 &&
          joins !== undefined &&
          joins.before === 0 &&
          joins.after === 0 &&
          joins.remove === 0 &&
          resolution.decision === REVISION_DECISIONS.REJECT &&
          isAddedRevision(node)
        )
          handKeptToFacingJoins({
            index,
            kept: {
              from: node.info.date ?? "",
              count: keptBoundaryCount(kept),
            },
            standing: kept.length,
            inEdgeRecords: {
              first:
                firstRecord === undefined ||
                keptOf(removalContent.slice(0, firstRecord)).length > 0
                  ? 0
                  : keptBoundaryCount(
                      keptOf(
                        removalContent.slice(firstRecord, firstRecord + 1),
                      ),
                    ),
              last:
                lastRecord === undefined ||
                keptOf(removalContent.slice(lastRecord + 1)).length > 0
                  ? 0
                  : keptBoundaryCount(
                      keptOf(removalContent.slice(lastRecord, lastRecord + 1)),
                    ),
            },
          });
        // Those beyond the record that piece continues stand at the removed
        // content's other edge: they go into the piece there where it is a
        // piece of the record the removal's seam cut, whose other piece
        // stands on the near side.
        const beyond = kept.length - inside;
        const far = piece === "before" ? "after" : "before";
        const farPiece = nodes[far === "after" ? index + 1 : index - 1];
        const cutAcross =
          farPiece !== undefined &&
          !isTrackedWrapper(farPiece) &&
          !isStandingBoundary(farPiece) &&
          (far === "after"
            ? nodes
                .slice(0, index)
                .some((near) => alikeDepth(near, farPiece) > 0)
            : nodes
                .slice(index + 1)
                .some((near) => alikeDepth(farPiece, near) > 0));
        if (
          piece !== undefined &&
          beyond > 0 &&
          cutAcross &&
          (node.resolutionJoins?.remove ?? 0) > 0
        ) {
          floating.set(far === "after" ? out.length : out.length - inside, {
            count: beyond,
            depth: node.resolutionJoins?.remove ?? 0,
            before: index > 0 ? sourceOwners.at(index - 1) : undefined,
            after: sourceOwners.at(index + 1),
            spanning: trackedWithin(node.content),
            piece: far,
          });
        }
        if (node.resolutionJoins === undefined) seams.push(out.length);
        else if ((node.resolutionJoins.plainSides?.length ?? 0) > 0) {
          plainSeams.add(out.length);
          // A cut closed on one side joins nothing across, but the
          // boundaries it kept still go into the piece on its other side.
          for (let at = out.length - kept.length; at <= out.length; at += 1)
            if (floating.has(at)) folds.push(at);
        } else {
          // Pieces of the cut this resolution also removes take that level
          // of it with them: the cut joins only what remains of what it split.
          const removalCut = cutLevelsOnceResolved(node, index);
          recordSourceSeam(out.length, removalCut);
          if (removalCut > 0) removalSeams.add(out.length);
          if (removalCut > 0 && isAddedRevision(node))
            addedRemovalSeams.add(out.length);
          // The kept markers beside it naming its side (or both) as their
          // cut's piece stood at a seam it took: they are kept there, and
          // its removal consumes the name.
          const taken =
            removalCut > 0 ? takeKeptBeside(index, keptStart) : undefined;
          if (taken !== undefined) exactKept.set(out.length, taken);
        }
        continue;
      }
      // A rejected deletion holding only boundaries a seam waits on (see
      // holdSeamOnKeptRemovals) leaves them standing at that seam: they float
      // there and the seam joins the pieces around them, as it does when the
      // removal that kept them is resolved after the deletion.
      const held = node.resolutionJoins?.remove ?? 0;
      if (
        node.type === "deletion" &&
        held > 0 &&
        resolution.decision === REVISION_DECISIONS.REJECT &&
        isKeptBoundaryCarrier(node)
      ) {
        const content = resolveList({
          nodes: node.content,
          resolution,
          path: [...path, index],
        }).nodes;
        // They float as of the revision their cut names they were kept from,
        // still standing in it once the deletion holding them goes.
        const keptFrom = keptCutFrom(node);
        appendUnowned(
          keptFrom === undefined
            ? content
            : content.map((child) =>
                withKeptCutsRecorded(
                  child,
                  { depth: held, keptFrom },
                  resolution.relocated.nodes,
                ),
              ),
        );
        recordSourceSeam(out.length, held);
        removalSeams.add(out.length);
        floating.set(out.length, {
          count: content.length,
          depth: held,
          before: index > 0 ? sourceOwners.at(index - 1) : undefined,
          after: sourceOwners.at(index + 1),
          spanning: [],
        });
        continue;
      }
      if (node.resolutionJoins === undefined) seams.push(out.length);
      else if (node.resolutionJoins.plainSides?.includes("before") === true)
        plainSeams.add(out.length);
      else {
        recordSourceSeam(out.length, node.resolutionJoins.before);
        if (node.resolutionJoins.before > 0) {
          recordKeptJoin(out.length, node);
          if (isRemovedRevisionNode(node)) ownContentAfter.add(out.length);
        }
        // Kept markers standing before it whose cut names this side take
        // the seam its join records: this insertion's content is the piece.
        const keptBefore =
          node.resolutionJoins.before > 0
            ? keptByPiece((at) => out[at]?.node, out.length - 1, -1, "after")
            : undefined;
        if (keptBefore !== undefined)
          exactKept.set(out.length, { ...keptBefore, piece: "after" });
      }
      const resolved = resolveList({
        nodes: node.content,
        resolution,
        path: [...path, index],
      });
      const content =
        isAddedRevision(node) &&
        resolution.decision === REVISION_DECISIONS.ACCEPT
          ? withPayloadEdgesClosed(
              acceptCopyOriginAncestor(resolved.nodes),
              node.resolutionJoins,
              resolution,
            )
          : resolved.nodes;
      if (resolved.edges?.first !== undefined)
        recordSourceSeam(out.length, resolved.edges.first);
      if (
        isAddedRevision(node) &&
        resolution.decision === REVISION_DECISIONS.ACCEPT &&
        node.resolutionJoins !== undefined
      ) {
        const first = content.at(0);
        const last = content.at(-1);
        if (node.resolutionJoins.after === 0 && first !== undefined)
          resolution.acceptedClosed.first.add(first);
        if (node.resolutionJoins.before === 0 && last !== undefined)
          resolution.acceptedClosed.last.add(last);
      }
      if (
        resolution.decision === REVISION_DECISIONS.ACCEPT &&
        node.resolutionJoins?.acceptance ===
          INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS
      ) {
        plainSeams.add(out.length);
        plainSeams.add(out.length + content.length);
      }
      const payload =
        isAddedRevision(node) &&
        resolution.decision === REVISION_DECISIONS.ACCEPT;
      for (const child of content) appendSource(child, index, payload);
      if (resolved.edges?.last !== undefined)
        recordSourceSeam(out.length, resolved.edges.last);
      if (node.resolutionJoins === undefined) seams.push(out.length);
      else if (node.resolutionJoins.plainSides?.includes("after") === true)
        plainSeams.add(out.length);
      else {
        recordSourceSeam(out.length, node.resolutionJoins.after);
        if (node.resolutionJoins.after > 0) {
          recordKeptJoin(out.length, node);
          if (isRemovedRevisionNode(node)) ownContentBefore.add(out.length);
        }
        const keptAfter =
          node.resolutionJoins.after > 0
            ? keptByPiece((at) => pending[at], index + 1, 1, "before")
            : undefined;
        if (keptAfter !== undefined)
          exactKept.set(out.length, { ...keptAfter, piece: "before" });
      }
      continue;
    }
    if (node.type === "run") {
      const run = resolveRun(node, resolution);
      if (run !== node) {
        changed = true;
        const selected = (node.propertyChanges ?? []).filter((change) =>
          resolution.ids.has(change.info.id),
        );
        if (
          selected.some(
            (change) =>
              change.boundaryJoins === undefined ||
              change.boundaryJoins.includes("before"),
          )
        )
          seams.push(out.length);
        if (
          selected.some(
            (change) =>
              change.boundaryJoins === undefined ||
              change.boundaryJoins.includes("after"),
          )
        )
          seams.push(out.length + 1);
      }
      appendSource(run, index);
      continue;
    }
    const children = childNodes(node);
    const inner =
      children === undefined
        ? undefined
        : resolveList({ nodes: children, resolution, path: [...path, index] });
    for (const handoff of inner?.enclosing ?? []) {
      if (handoff.levels > 1) {
        enclosing.push({
          levels: handoff.levels - 1,
          transfer: handoff.transfer,
        });
        continue;
      }
      if (
        inner !== undefined &&
        inner.nodes.length === 0 &&
        (children?.length ?? 0) > 0 &&
        hasSlots(node, handoff.transfer.source)
      )
        resolution.outstanding.push(handoff.transfer);
    }
    if (resolution.outstanding.length > 0) {
      // A prefix removed inside this record hands its identities on to the
      // pieces after it.
      const following = restoreOutstandingIdentities(
        pending.slice(index + 1),
        resolution,
      );
      pending.splice(
        index + 1,
        pending.length - index - 1,
        ...prepareDeferredRemovals(following, resolution),
      );
    }
    const cut = remainingCut(node, index);
    if (cut !== undefined) changed = true;
    const recut = (record: InlineNode): InlineNode =>
      cut === undefined ||
      !isTrackedWrapper(record) ||
      record.resolutionJoins === undefined
        ? record
        : Object.assign({}, record, {
            resolutionJoins: Object.assign({}, record.resolutionJoins, {
              remove: cut,
            }),
          });
    if (children === undefined || inner === undefined || !inner.changed) {
      const leaf =
        node.type === "simpleField" || node.type === "complexField"
          ? resolveFieldRuns(node, resolution)
          : node;
      if (leaf !== node) changed = true;
      appendSource(recut(leaf), index);
      continue;
    }
    changed = true;
    // A tracked change the resolution emptied goes. So does any record an
    // acceptance empties: it removes content as a direct deletion does, which
    // leaves no record with nothing. Kept, the record could not tell a later
    // resolution it was emptied rather than empty before, and would stay where
    // accepting the same changes at once folds it into an alike neighbour.
    // So does a record a rejection empties of insertions an acceptance left
    // as all it held: accepting that content last would have emptied it.
    if (
      (isTrackedWrapper(node) ||
        resolution.decision === REVISION_DECISIONS.ACCEPT ||
        children.some(
          (child) =>
            isAddedRevision(child) &&
            child.resolutionJoins?.emptiedHolder === true,
        )) &&
      inner.nodes.length === 0 &&
      children.length > 0
    ) {
      // Removing an unselected parent establishes no cut between its neighbours.
      // Fold only fragments this resolution emptied, preserving authored siblings.
      folds.push(out.length);
      // Accepted new content beside it, as an edit made directly would be,
      // merges with the plain runs it now meets.
      if (
        isTrackedWrapper(node) &&
        (node.resolutionJoins?.plainSides?.length ?? 0) > 0
      )
        plainSeams.add(out.length);
      continue;
    }
    let rebuilt = recut(
      rebuildNode(
        node,
        withEmptiedHolder(node, children, inner.nodes, resolution),
      ),
    );
    if (
      isTrackedWrapper(rebuilt) &&
      rebuilt.resolutionJoins?.retainedAfter !== undefined
    ) {
      const sourceKeys = new Set(identityKeysIn(inner.nodes));
      const retainedAfter = rebuilt.resolutionJoins.retainedAfter.flatMap(
        (entry) => {
          const source: RetainedIdentity["source"][number][] = [];
          const target: RetainedIdentity["target"][number][] = [];
          for (const [slotIndex, slot] of entry.source.entries()) {
            if (!sourceKeys.has(slotKey(slot))) continue;
            const bound = entry.target.at(slotIndex);
            if (bound === undefined)
              panic(
                "Validated cut provenance has paired source and target slots.",
              );
            source.push(slot);
            target.push(bound);
          }
          return source.length === 0
            ? []
            : [{ depth: entry.depth, source, target }];
        },
      );
      rebuilt = Object.assign({}, rebuilt, {
        resolutionJoins: Object.assign({}, rebuilt.resolutionJoins, {
          retainedAfter,
        }),
      });
    }
    // A container the resolution emptied, or one with such a container at an
    // edge, meets its neighbours there: otherwise an emptied piece of a cut
    // container, with no change resolved between it and the other piece,
    // would stay beside it rather than fold back in.
    const ends =
      inner.nodes.length === 0
        ? { first: true, last: true }
        : emptiedEndsOf(inner.nodes, resolution.emptied);
    if (inner.nodes.length === 0) resolution.emptied.add(rebuilt);
    if (ends.first) folds.push(out.length);
    if (ends.last) folds.push(out.length + 1);
    if (inner.edges !== undefined) {
      const edges: CutEdges = {};
      if (inner.edges.first !== undefined) edges.first = inner.edges.first + 1;
      if (inner.edges.last !== undefined) edges.last = inner.edges.last + 1;
      resolution.cutEdges.set(rebuilt, edges);
    }
    appendSource(rebuilt, index);
  }
  // An inner cut can meet the edge of a container while its other fragment
  // remains beyond a pending insertion. Preserve that exact seam on its
  // actual blocker; accepting the payload cancels it, rejecting restores it.
  for (const [index, { node, owner }] of out.entries()) {
    if (!isAddedRevision(node) || node.resolutionJoins === undefined) continue;
    if (node.resolutionJoins.remove === 0) continue;
    const left = out.at(index - 1)?.node;
    const right = out.at(index + 1)?.node;
    const depths = [
      index > 0 && left !== undefined
        ? resolution.cutEdges.get(left)?.last
        : undefined,
      right === undefined ? undefined : resolution.cutEdges.get(right)?.first,
    ];
    const groups = [...(node.resolutionJoins.deferredRemove ?? [])];
    for (const depth of depths) {
      if (depth !== undefined) groups.push({ depth, blockers: [node.info.id] });
    }
    if (groups.length === (node.resolutionJoins.deferredRemove?.length ?? 0))
      continue;
    const deferred = Object.assign({}, node, {
      resolutionJoins: Object.assign({}, node.resolutionJoins, {
        deferredRemove: canonicalDeferredGroups(groups),
      }),
    });
    out[index] = { node: deferred, owner };
    changed = true;
  }
  if (!changed) {
    return { nodes: out.map(({ node }) => node), changed };
  }
  // A container this resolution rebuilt can meet the seam with an inner cut
  // at its edge whose other piece lies beyond the seam: a removed insertion
  // at that edge cut deeper than the seam does. Once the seam joins the
  // container's level, that cut continues through it, as the seam deferred
  // on a pending insertion there joins when the insertion is rejected later.
  for (const seam of addedRemovalSeams) {
    const depth = exactSeams.get(seam);
    if (depth === undefined) continue;
    const standing = floating.get(seam)?.count ?? 0;
    // Boundaries this resolution kept beside the seam are no record of it:
    // the seam takes them inside the alike pieces it joins, so the
    // container beyond them meets it, inner cut and all, as it does when
    // their removal is resolved first.
    const keptStanding = (index: number): boolean =>
      keptFromAt(index) !== undefined;
    let leftAt = seam - standing - 1;
    while (leftAt >= 0 && keptStanding(leftAt)) leftAt -= 1;
    let rightAt = seam;
    while (rightAt < out.length && keptStanding(rightAt)) rightAt += 1;
    const pastKept =
      leftAt >= 0 &&
      rightAt < out.length &&
      sameOwnFields(
        out[leftAt]?.node ?? panic("A list position holds a record."),
        out[rightAt]?.node ?? panic("A list position holds a record."),
      );
    if (!pastKept) {
      leftAt = seam - standing - 1;
      rightAt = seam;
    }
    const left = out.at(leftAt)?.node;
    const right = out.at(rightAt)?.node;
    const reach = Math.max(
      depth,
      leftAt >= 0 && left !== undefined
        ? (resolution.cutEdges.get(left)?.last ?? 0)
        : 0,
      right === undefined ? 0 : (resolution.cutEdges.get(right)?.first ?? 0),
    );
    if (reach === depth) continue;
    exactSeams.set(seam, reach);
    sourceEdges.set(seam, Math.max(sourceEdges.get(seam) ?? 0, reach));
    const kept = floating.get(seam);
    if (kept !== undefined) floating.set(seam, { ...kept, depth: reach });
  }
  const edges: CutEdges = {};
  // Boundaries standing between the list's edge and a seam hold no piece of
  // the record it cut: the seam is still the edge's, and a join at that edge
  // takes them inside the piece it rejoins (see withSeamBoundariesInside).
  const edgeDepth = (
    at: number,
    step: 1 | -1,
    limit: number,
  ): number | undefined => {
    let depth = sourceEdges.get(at);
    for (
      let index = at;
      index !== limit &&
      isStandingBoundary(out[step === 1 ? index : index - 1]?.node);
      index += step
    ) {
      const next = sourceEdges.get(index + step);
      if (next !== undefined) depth = Math.max(depth ?? 0, next);
    }
    return depth;
  };
  const firstDepth = edgeDepth(0, 1, out.length);
  const lastDepth = edgeDepth(out.length, -1, 0);
  if (firstDepth !== undefined) edges.first = firstDepth;
  if (lastDepth !== undefined) edges.last = lastDepth;
  // Accepted new content stands beside a pending removal: the removal's
  // seam there no longer joins source pieces. Accepting the removal later
  // merges plain runs across it, as accepting both at once does, and joins
  // nothing else.
  const pendingRemoval = (node: InlineNode | undefined): boolean =>
    node !== undefined &&
    (((node.type === "deletion" || node.type === "moveFrom") &&
      !resolution.ids.has(node.info.id) &&
      node.resolutionJoins !== undefined) ||
      // A cut piece holding only content on its way out goes with it.
      (isTrackedWrapper(node) &&
        !resolution.ids.has(node.info.id) &&
        node.content.length > 0 &&
        node.content.every((child) => pendingRemoval(child))));
  const acceptAsPlain = (index: number, facing: "before" | "after"): void => {
    const occurrence = out[index];
    const node = occurrence?.node;
    if (
      occurrence === undefined ||
      node === undefined ||
      !isTrackedWrapper(node)
    )
      return;
    out[index] = occurrenceOf(
      Object.assign({}, node, {
        resolutionJoins: Object.assign(
          { before: 0, after: 0, remove: 0 },
          node.resolutionJoins,
          {
            plainSides: [
              ...new Set([...(node.resolutionJoins?.plainSides ?? []), facing]),
            ],
          },
        ),
      }),
      occurrence,
    );
    changed = true;
  };
  for (const seam of plainSeams) {
    const left = out.at(seam - 1)?.node;
    const right = out.at(seam)?.node;
    if (
      seam > 0 &&
      left?.type === "run" &&
      right?.type === "run" &&
      runsMergeable(left, right)
    )
      seams.push(seam);
    else if (pendingRemoval(right)) acceptAsPlain(seam, "before");
    else if (seam > 0 && pendingRemoval(left)) acceptAsPlain(seam - 1, "after");
  }
  // Boundaries a removal kept can stand at its seam held in a pending
  // deletion that no piece there can take (a link holds no tracked change).
  // The pieces meet once that deletion goes, either way: the seam waits on
  // it, so rejecting it later joins the pieces around the boundaries and
  // accepting it joins them across its place, as resolving both at once
  // does. Whether the seam now waits on such a deletion.
  const holdSeamOnKeptRemovals = ({
    seam,
    count,
    depth,
  }: {
    seam: number;
    count: number;
    depth: number;
  }): boolean => {
    const before = out[seam - count - 1]?.node;
    const after = out[seam]?.node;
    if (
      depth <= 0 ||
      before === undefined ||
      after === undefined ||
      isTrackedWrapper(before) ||
      isTrackedWrapper(after) ||
      childNodes(before) === undefined ||
      !sameOwnFields(before, after)
    )
      return false;
    let held = false;
    for (let index = seam - count; index < seam; index += 1) {
      const occurrence = out[index];
      const node = occurrence?.node;
      if (
        occurrence === undefined ||
        node?.type !== "deletion" ||
        resolution.ids.has(node.info.id) ||
        !isKeptBoundaryCarrier(node)
      )
        continue;
      const joins = node.resolutionJoins ?? { before: 0, after: 0, remove: 0 };
      if (joins.remove >= depth) {
        held = true;
        continue;
      }
      const next = Object.assign({}, node, {
        resolutionJoins: Object.assign({}, joins, { remove: depth }),
      });
      carryResolutionState(resolution, node, next);
      out[index] = occurrenceOf(next, occurrence);
      changed = true;
      held = true;
    }
    return held;
  };
  // Below the records a seam joins, pending insertions can stand between its
  // pieces, new content made inside one of them after the cut. Joined to
  // nothing on the side facing the seam, they are no piece of the record it
  // cut: the rest of the seam waits on them, as it does where such
  // insertions meet the seam itself, so rejecting them later joins the
  // pieces, as rejecting everything at once does, and accepting them keeps
  // the pieces apart. `undefined` where no such insertion stands there.
  const deferOnNestedInsertions = (
    left: InlineNode,
    right: InlineNode,
    leftChildren: readonly InlineNode[],
    rightChildren: readonly InlineNode[],
    depth: number,
  ): [InlineNode, InlineNode] | undefined => {
    const blocking = (
      node: InlineNode | undefined,
      facing: "before" | "after",
    ): boolean =>
      node !== undefined &&
      isTrackedWrapper(node) &&
      isAddedRevision(node) &&
      node.resolutionJoins !== undefined &&
      node.resolutionJoins[facing] === 0 &&
      !resolution.ids.has(node.info.id);
    // Kept markers standing among them wait on them too: they stand in the
    // cut, beside the insertions, until the pieces meet.
    const waits = (
      node: InlineNode | undefined,
      facing: "before" | "after",
    ): boolean =>
      blocking(node, facing) || (node !== undefined && standsInKeptCut(node));
    let leftEnd = leftChildren.length - 1;
    while (leftEnd >= 0 && waits(leftChildren[leftEnd], "after")) leftEnd -= 1;
    let rightStart = 0;
    while (
      rightStart < rightChildren.length &&
      waits(rightChildren[rightStart], "before")
    )
      rightStart += 1;
    const leftPiece = leftChildren.at(leftEnd);
    const rightPiece = rightChildren.at(rightStart);
    if (
      (leftEnd === leftChildren.length - 1 && rightStart === 0) ||
      leftEnd < 0 ||
      leftPiece === undefined ||
      rightPiece === undefined ||
      isTrackedWrapper(leftPiece) ||
      isTrackedWrapper(rightPiece) ||
      !sameOwnFields(leftPiece, rightPiece)
    )
      return undefined;
    const blockerNodes = [
      ...leftChildren.slice(leftEnd + 1),
      ...rightChildren.slice(0, rightStart),
    ];
    const blockers = blockerNodes.flatMap((node) =>
      isTrackedWrapper(node) && !standsInKeptCut(node) ? [node.info.id] : [],
    );
    if (blockers.length === 0) return undefined;
    const deferred = (node: InlineNode): InlineNode => {
      if (standsInKeptCut(node)) return node;
      if (!isTrackedWrapper(node) || node.resolutionJoins === undefined)
        return panic(
          "A deferring seam blocker is a pending change with joins.",
        );
      const joins = node.resolutionJoins;
      const next = Object.assign({}, node, {
        resolutionJoins: Object.assign({}, joins, {
          deferredRemove: canonicalDeferredGroups([
            ...(joins.deferredRemove ?? []),
            { depth: depth - 1, blockers },
          ]),
        }),
      });
      carryResolutionState(resolution, node, next);
      return next;
    };
    return [
      leftEnd === leftChildren.length - 1
        ? left
        : rebuildNode(left, [
            ...leftChildren.slice(0, leftEnd + 1),
            ...leftChildren.slice(leftEnd + 1).map(deferred),
          ]),
      rightStart === 0
        ? right
        : rebuildNode(right, [
            ...rightChildren.slice(0, rightStart).map(deferred),
            ...rightChildren.slice(rightStart),
          ]),
    ];
  };
  // Below the records a seam joins, its pieces can still be held in pending
  // deletions or move sources with boundaries kept standing between them:
  // the pieces stand only once those are rejected, so the rest of the seam
  // waits on them, as it does where such a change meets the seam itself.
  // Dropped, rejecting them later would leave the pieces apart where
  // rejecting everything at once joins them around the boundaries.
  const deferNestedRemovalSeam = (
    left: InlineNode,
    right: InlineNode,
    depth: number,
  ): [InlineNode, InlineNode] => {
    const leftChildren = childNodes(left);
    const rightChildren = childNodes(right);
    if (
      depth <= 1 ||
      leftChildren === undefined ||
      rightChildren === undefined ||
      !sameOwnFields(left, right)
    )
      return [left, right];
    let leftEnd = leftChildren.length - 1;
    while (leftEnd >= 0 && isStandingBoundary(leftChildren[leftEnd]))
      leftEnd -= 1;
    let rightStart = 0;
    while (
      rightStart < rightChildren.length &&
      isStandingBoundary(rightChildren[rightStart])
    )
      rightStart += 1;
    const leftHolder = leftChildren.at(leftEnd);
    const rightHolder = rightChildren.at(rightStart);
    if (leftHolder === undefined || rightHolder === undefined)
      return [left, right];
    const standingNodes = [
      ...leftChildren.slice(leftEnd + 1),
      ...rightChildren.slice(0, rightStart),
    ];
    const standing = standingNodes.length > 0;
    // Kept markers standing in the cut wait with the insertions beside them.
    if (standingNodes.every(standsInKeptCut)) {
      const blocked = deferOnNestedInsertions(
        left,
        right,
        leftChildren,
        rightChildren,
        depth,
      );
      if (blocked !== undefined) return blocked;
    }
    if (!standing) {
      const inner = deferNestedRemovalSeam(leftHolder, rightHolder, depth - 1);
      return [
        inner[0] === leftHolder
          ? left
          : rebuildNode(left, [...leftChildren.slice(0, -1), inner[0]]),
        inner[1] === rightHolder
          ? right
          : rebuildNode(right, [inner[1], ...rightChildren.slice(1)]),
      ];
    }
    const pendingHolder = (
      node: InlineNode,
    ): node is Extract<InlineNode, { type: "deletion" | "moveFrom" }> =>
      (node.type === "deletion" || node.type === "moveFrom") &&
      !resolution.ids.has(node.info.id) &&
      node.resolutionJoins !== undefined;
    if (
      !pendingHolder(leftHolder) ||
      !pendingHolder(rightHolder) ||
      !sameOwnFields(leftHolder, rightHolder)
    )
      return [left, right];
    const leftPiece = leftHolder.content.at(-1);
    const rightPiece = rightHolder.content.at(0);
    if (
      leftPiece === undefined ||
      rightPiece === undefined ||
      !sameOwnFields(leftPiece, rightPiece)
    )
      return [left, right];
    const blockers = [leftHolder.info.id, rightHolder.info.id];
    const withGroup = (
      holder: Extract<InlineNode, { type: "deletion" | "moveFrom" }>,
      edge: "start" | "end",
    ): InlineNode => {
      const joins =
        holder.resolutionJoins ?? panic("A seam holder records its joins.");
      const next = Object.assign({}, holder, {
        resolutionJoins: Object.assign({}, joins, {
          deferredRemove: canonicalDeferredGroups([
            ...(joins.deferredRemove ?? []),
            { depth: depth - 1, blockers, edge },
          ]),
        }),
      });
      carryResolutionState(resolution, holder, next);
      return next;
    };
    return [
      rebuildNode(left, [
        ...leftChildren.slice(0, leftEnd),
        withGroup(leftHolder, "end"),
        ...leftChildren.slice(leftEnd + 1),
      ]),
      rebuildNode(right, [
        ...rightChildren.slice(0, rightStart),
        withGroup(rightHolder, "start"),
        ...rightChildren.slice(rightStart + 1),
      ]),
    ];
  };
  // A recorded seam a pending deletion or move source stands at, holding
  // the piece on its side, joins once that change's content is kept: the
  // change carries the seam on that side, as deep as it was recorded.
  // Dropped, rejecting the change later would leave the pieces apart where
  // rejecting both at once joins them.
  const carrySeam = (
    node: InlineNode,
    side: "before" | "after",
    depth: number | undefined,
  ): InlineNode => {
    if (
      depth === undefined ||
      depth === 0 ||
      !isTrackedWrapper(node) ||
      isAddedRevision(node) ||
      resolution.ids.has(node.info.id) ||
      node.resolutionJoins === undefined ||
      node.resolutionJoins.plainSides?.includes(side) === true ||
      node.resolutionJoins[side] >= depth
    )
      return node;
    const carried = Object.assign({}, node, {
      resolutionJoins: Object.assign({}, node.resolutionJoins, {
        [side]: depth,
      }),
    });
    carryResolutionState(resolution, node, carried);
    return carried;
  };
  // Boundaries kept from a later insertion this resolution removes, standing
  // just before the boundaries a cut's removal floats, stood inside that cut
  // when a piece of it stands beyond them: they float with them, as
  // rejecting the cut after the insertion places them. With no piece there
  // the cut has one side, and they stood beside it.
  const pieceAt = (
    index: number,
    edge: "start" | "end",
  ): InlineNode | undefined => {
    let node: InlineNode | undefined = out[index]?.node;
    while (node !== undefined && isTrackedWrapper(node)) {
      const children: readonly InlineNode[] = childNodes(node) ?? [];
      node = edge === "start" ? children.at(0) : children.at(-1);
    }
    return node;
  };
  const floated = new Set(
    [...floating].flatMap(([at, { count }]) =>
      Array.from({ length: count }, (_, offset) => at - count + offset),
    ),
  );
  for (const [at, entry] of [...floating]) {
    const keptFrom = keptFromIn(at - entry.count, at);
    if (keptFrom === undefined) continue;
    let count = entry.count;
    for (;;) {
      const index = at - count - 1;
      const date = keptFromAt(index);
      if (date === undefined || date <= keptFrom || floated.has(index)) break;
      floated.add(index);
      count += 1;
    }
    if (
      count !== entry.count &&
      alikeDepth(pieceAt(at - count - 1, "end"), pieceAt(at, "start")) > 0
    )
      floating.set(at, { ...entry, count });
    else
      for (let index = at - count; index < at - entry.count; index += 1)
        floated.delete(index);
  }
  /**
   * Hand the join at a seam to the pending insertion at `holderAt` holding the
   * other piece of the record the plain record at `piece` was cut from;
   * whether it did.
   */
  const passesJoinOn = (
    holderAt: number,
    piece: number,
    depth: number,
  ): boolean => {
    const occurrence = out[holderAt];
    const holder = occurrence?.node;
    const near = out[piece]?.node;
    if (
      occurrence === undefined ||
      holder === undefined ||
      near === undefined ||
      isTrackedWrapper(near) ||
      !isTrackedWrapper(holder) ||
      !isAddedRevision(holder) ||
      resolution.ids.has(holder.info.id)
    )
      return false;
    const side = holderAt > piece ? "before" : "after";
    const edge =
      side === "before" ? holder.content.at(0) : holder.content.at(-1);
    if (
      edge === undefined ||
      (side === "before" ? alikeDepth(near, edge) : alikeDepth(edge, near)) <
        Math.min(depth, 1)
    )
      return false;
    const { node: joined } = handJoin(
      holder.resolutionJoins === undefined
        ? Object.assign({}, holder, {
            resolutionJoins: { before: 0, after: 0, remove: 0 },
          })
        : holder,
      side,
      depth,
    );
    carryResolutionState(resolution, holder, joined);
    out[holderAt] = occurrenceOf(joined, occurrence);
    changed = true;
    return true;
  };
  // Boundaries an earlier rejection kept at a seam still stand between the
  // pieces it joins: they float there as the ones this resolution kept do,
  // so rejecting first and joining later places them as the other order.
  for (const [recorded, depth] of [...exactSeams]) {
    if (floating.has(recorded)) {
      // A deferred seam's kept boundaries stand just after the ones this
      // removal floats there: one group, kept as of the earlier revision.
      const entry = floating.get(recorded);
      const deferred = exactKept.get(recorded);
      if (
        entry === undefined ||
        deferred?.count === undefined ||
        deferred.from === keptFromIn(recorded - entry.count, recorded)
      )
        continue;
      const after = keptRun({
        at: (index) => out[index]?.node,
        from: recorded,
        step: 1,
        limit: deferred.count,
      }).records;
      if (after === 0) continue;
      floating.delete(recorded);
      exactSeams.delete(recorded);
      exactSeams.set(
        recorded + after,
        Math.max(exactSeams.get(recorded + after) ?? 0, depth),
      );
      floating.set(recorded + after, { ...entry, count: entry.count + after });
      continue;
    }
    const leftPiece = out[recorded - 1]?.node;
    const rightPiece = out[recorded]?.node;
    // A recorded source cut already meeting its matching pieces rejoins
    // them here. A marker-only piece is an authored record, not a boundary
    // kept outside those records by an earlier rejection. Explicit kept
    // provenance still owns its boundaries, even beside matching pieces.
    if (
      !exactKept.has(recorded) &&
      leftPiece !== undefined &&
      rightPiece !== undefined &&
      alikeDepth(leftPiece, rightPiece) >= depth
    )
      continue;
    // Boundaries another seam already floats belong to that seam's removal:
    // this seam's kept ones are the others beside it.
    const claimed = new Set(
      [...floating].flatMap(([at, { count }]) =>
        Array.from({ length: count }, (_, offset) => at - count + offset),
      ),
    );
    const kept = exactKept.get(recorded);
    /**
     * Unclaimed kept records from `from`, for at most `limit` boundaries.
     * Without provenance on the seam, only a marker that records the cut it
     * stands in was kept there: an unmarked one is authored or accepted
     * content and blocks the join.
     */
    const standingAt = (from: number, step: 1 | -1, limit: number): KeptRun =>
      keptRun({
        at: (index) => {
          const node = claimed.has(index) ? undefined : out[index]?.node;
          return node === undefined ||
            (kept === undefined && !standsInKeptCut(node))
            ? undefined
            : node;
        },
        from,
        step,
        limit,
      });
    // Only those the rejection kept, nearest the seam: boundaries another
    // removal kept beside them stand where they are, as resolving both at
    // once leaves them.
    const limit = kept?.count ?? Number.POSITIVE_INFINITY;
    // A kept revision's own content beside the seam its join recorded is
    // authored content, not boundaries kept there: it stays where it is.
    const afterRun = ownContentAfter.has(recorded)
      ? NO_KEPT_RUN
      : standingAt(recorded, 1, limit);
    const after = afterRun.records;
    // A counted provenance can name boundaries on both sides: the seam's own
    // and those of a seam deferred into it, kept from the same revision.
    const before =
      ownContentBefore.has(recorded) ||
      (afterRun.boundaries > 0 &&
        (kept?.count === undefined || afterRun.boundaries >= kept.count))
        ? 0
        : standingAt(recorded - 1, -1, limit - afterRun.boundaries).records;
    const seam = recorded + after;
    const count = after + before;
    const left = out[seam - count - 1];
    const right = out[seam];
    // A side with no record still takes them when the seam's provenance
    // names the other side's record as the piece of the cut they stood in.
    const pieceOnly =
      kept?.piece !== undefined &&
      (kept.piece === "after" ? right : left) !== undefined;
    if (
      count === 0 ||
      ((left === undefined || right === undefined) && !pieceOnly)
    )
      continue;
    // With no kept provenance on the seam, only boundaries between two
    // pieces it joins can be ones a rejection kept there. Beside records it
    // does not join they are the document's own and stay where they are.
    // An accepted insertion whose content stands as boundaries only holds no
    // piece of the record its join cut: the join reaches the record beyond
    // them. A pending insertion holding that record's other piece takes the
    // join, so accepting it later joins the pieces around the boundaries, as
    // accepting both at once does.
    if (
      kept === undefined &&
      resolution.decision === REVISION_DECISIONS.ACCEPT &&
      keptJoinSeams.has(recorded) &&
      passesJoinOn(seam, seam - count - 1, depth)
    )
      continue;
    if (
      kept === undefined &&
      (left === undefined ||
        right === undefined ||
        !sameOwnFields(left.node, right.node))
    )
      continue;
    exactSeams.delete(recorded);
    exactSeams.set(seam, Math.max(exactSeams.get(seam) ?? 0, depth));
    // Boundaries this resolution kept there already float with the removal
    // they came from, which knows the records it cut and the changes it held;
    // a seam recorded beside it, as one a rejected deletion leaves, joins
    // with them rather than replacing what they float with.
    if (floating.has(seam)) continue;
    floating.set(seam, {
      count,
      depth,
      before: left?.owner,
      after: right?.owner,
      spanning: [],
      ...(kept?.piece === undefined ? {} : { piece: kept.piece }),
    });
  }
  const merges = new Set(seams);
  /** Ids seam joins retired, each with the id of the piece it merged into. */
  const retired = new Map<number, number>();
  for (const recorded of [
    ...new Set([...seams, ...exactSeams.keys(), ...folds]),
  ].toSorted((left, right) => right - left)) {
    // Boundaries kept just before this seam go inside the resolved piece it
    // joins (after it, else before it), so the pieces still meet.
    let seam = recorded;
    const kept = floating.get(recorded);
    /** Kept boundaries left standing before the seam, held by no piece. */
    let standingKept = kept?.count ?? 0;
    if (kept !== undefined) {
      const beforeOccurrence = out[seam - kept.count - 1];
      const afterOccurrence = out[seam];
      const before = beforeOccurrence?.node;
      const after = afterOccurrence?.node;
      const standing = out
        .slice(seam - kept.count, seam)
        .map(({ node }) => node);
      const keptFrom = keptFromIn(seam - kept.count, seam);
      /**
       * The boundaries as placed at `piece`'s `edge`. Inside it, they still
       * record their cut while a pending change meets them there, inside the
       * piece or across from it: resolving that change can still let the
       * records around them meet.
       */
      const placedAt = (
        piece: InlineNode,
        edge: "start" | "end",
      ): InlineNode[] => {
        const inside = childNodes(piece);
        const waiting =
          pendingAtEdge(
            edge === "start" ? inside?.at(0) : inside?.at(-1),
            edge,
          ) ||
          (edge === "start"
            ? pendingAtEdge(before, "end")
            : pendingAtEdge(after, "start"));
        return waiting
          ? standing
          : standing.map((node) =>
              withoutKeptCutsRecorded(node, resolution.relocated.nodes),
            );
      };
      // The depth counts the containers the removed content cut, read on
      // the records either side as the list held them: a container this
      // resolution resolves away is a level no more, and a tracked change on
      // one side only was never one. A side with no source record is read
      // against itself, as is a piece by its own join: what stands across
      // from it is no piece of the record the removed content cut, and its
      // join already says how deep in that record it was made.
      const ownerDepth = ({
        occurrence,
        source,
        side,
      }: {
        occurrence: SourceOccurrence | undefined;
        source: SourceOwner | undefined;
        side: SourceOwnerSide;
      }): number | undefined => {
        const depth = sourceOwnerDepthAt({
          owner: occurrence?.owner,
          source,
          side,
          depth: kept.depth,
        });
        if (depth === undefined) return undefined;
        const near = occurrence?.source;
        if (near === undefined) return depth;
        const opposite = (
          side === "before" ? afterOccurrence : beforeOccurrence
        )?.source;
        // A standing boundary is a source leaf, not a container the cut split.
        const far =
          kept.piece === side || isStandingBoundary(opposite)
            ? undefined
            : opposite;
        // A tracked change this resolution lifts recorded its join from its
        // content: the cut's levels start inside it.
        const lifted = (chain: InlineNode[]): InlineNode[] => {
          let start = 0;
          while (start < chain.length - 1) {
            const node = chain[start];
            if (
              node === undefined ||
              !isTrackedWrapper(node) ||
              !resolution.ids.has(node.info.id) ||
              removesContent(node, resolution.decision)
            )
              break;
            start += 1;
          }
          return start === 0 ? chain : chain.slice(start);
        };
        const nearChain = lifted(
          edgeChain(near, side === "before" ? "end" : "start"),
        );
        const farChain =
          far === undefined
            ? nearChain
            : lifted(edgeChain(far, side === "before" ? "start" : "end"));
        return cutReach({
          before: side === "before" ? nearChain : farChain,
          after: side === "before" ? farChain : nearChain,
          depth,
          spanning: far === undefined ? [] : kept.spanning,
          resolution,
        })[side];
      };
      const beforeDepth = ownerDepth({
        occurrence: beforeOccurrence,
        source: kept.before,
        side: "before",
      });
      const afterDepth = ownerDepth({
        occurrence: afterOccurrence,
        source: kept.after,
        side: "after",
      });
      // The records either side that own the cut take the boundaries, as
      // their source owners say. Two never do: a pending insertion made
      // after the revision they were kept from, which is no piece of what
      // it cut, and any record while the other side is a pending deletion or
      // move source, whose piece stands only once that is rejected. Nor does
      // a record across from such a later insertion, unless its own join
      // makes it a piece of the record the removed content cut: whichever
      // way the insertion is resolved, that piece takes them. Otherwise the
      // boundaries stand at the seam until its pieces meet.
      const heldInRemoval = (node: InlineNode | undefined): boolean =>
        node !== undefined &&
        (node.type === "deletion" || node.type === "moveFrom") &&
        !resolution.ids.has(node.info.id);
      const laterInsertion = (node: InlineNode | undefined): boolean =>
        node !== undefined &&
        isTrackedWrapper(node) &&
        isAddedRevision(node) &&
        !resolution.ids.has(node.info.id) &&
        keptFrom !== undefined &&
        (node.info.date ?? "") > keptFrom;
      const unjoined =
        kept.cut === "deferred" &&
        (before === undefined ||
          after === undefined ||
          !sameOwnFields(before, after));
      const takes = (side: "before" | "after"): boolean =>
        !unjoined &&
        !heldInRemoval(before) &&
        !heldInRemoval(after) &&
        !laterInsertion(side === "before" ? before : after) &&
        (kept.piece === side ||
          !laterInsertion(side === "before" ? after : before));
      const intoAfter =
        !takes("after") || after === undefined || afterDepth === undefined
          ? undefined
          : withBoundariesAtEdge(
              after,
              placedAt(after, "start"),
              afterDepth,
              "start",
            );
      const intoBefore =
        !takes("before") ||
        intoAfter !== undefined ||
        before === undefined ||
        beforeDepth === undefined
          ? undefined
          : withBoundariesAtEdge(
              before,
              placedAt(before, "end"),
              beforeDepth,
              "end",
            );
      if (intoAfter !== undefined && after !== undefined) {
        carryResolutionState(resolution, after, intoAfter);
        out.splice(
          seam - kept.count,
          kept.count + 1,
          occurrenceOf(intoAfter, afterOccurrence),
        );
        seam -= kept.count;
        standingKept = 0;
      } else if (intoBefore !== undefined && before !== undefined) {
        carryResolutionState(resolution, before, intoBefore);
        out.splice(
          seam - kept.count - 1,
          kept.count + 1,
          occurrenceOf(intoBefore, beforeOccurrence),
        );
        seam -= kept.count;
        standingKept = 0;
      } else {
        const depth = exactSeams.get(recorded) ?? kept.depth;
        // No piece takes them while a pending change stands at the seam:
        // each records the cut it stands in, to meet the pieces later.
        const pendingAt = (node: InlineNode | undefined): boolean =>
          node !== undefined &&
          isTrackedWrapper(node) &&
          !resolution.ids.has(node.info.id);
        const waiting = [
          out[seam - kept.count - 1]?.node,
          out[seam]?.node,
          ...out.slice(seam - kept.count, seam).map(({ node }) => node),
        ].some(pendingAt);
        if (waiting && keptFrom !== undefined && depth > 0)
          for (let index = seam - kept.count; index < seam; index += 1) {
            const occurrence = out[index];
            if (occurrence === undefined) continue;
            const marked = withKeptCuts(occurrence.node, { depth, keptFrom });
            if (marked === occurrence.node) continue;
            out[index] = occurrenceOf(marked, occurrence);
            changed = true;
          }
        if (holdSeamOnKeptRemovals({ seam, count: kept.count, depth }))
          continue;
      }
    }
    const leftOccurrence = out[seam - 1];
    const rightOccurrence = out[seam];
    const left = leftOccurrence?.node;
    const right = rightOccurrence?.node;
    if (left === undefined || right === undefined) continue;
    let exactDepth = exactSeams.get(recorded);
    if (exactDepth !== undefined) {
      // Extend only an already recorded outer cut with an unconsumed inner
      // source cut. Equal authored containers alone never establish a seam.
      exactDepth = Math.max(
        exactDepth,
        resolution.cutEdges.get(left)?.last ?? 0,
        resolution.cutEdges.get(right)?.first ?? 0,
      );
    }
    // A pending insertion can stand between the recorded seam's source
    // pieces, as a tracked replacement's payload stands before its deletion.
    // Keep the seam on that insertion, at any depth, as a split seam is kept
    // on the insertions between its paragraphs: rejecting it later joins
    // the pieces it separated, accepting it keeps them apart.
    // Boundaries floating at the seam that no piece took stand there, each
    // recording its cut and the revision it was kept from, so they land in
    // the pieces as they would had the insertion gone first.
    // Pending insertions either side, alike down to the seam's depth, are the
    // pieces the removed content cut, not insertions standing between them:
    // the seam joins them now. Kept as blockers, removing them later would
    // join the records beyond them, which the removed content never cut.
    const cutPieces =
      exactDepth !== undefined &&
      isTrackedWrapper(left) &&
      isTrackedWrapper(right) &&
      alikeDepth(left, right) >= exactDepth;
    if (
      exactDepth !== undefined &&
      exactDepth > 0 &&
      (kept === undefined || standingKept > 0) &&
      !cutPieces
    ) {
      const blocking = (index: number): boolean => {
        const node = out[index]?.node;
        return (
          node !== undefined &&
          isAddedRevision(node) &&
          isTrackedWrapper(node) &&
          node.resolutionJoins !== undefined &&
          !resolution.ids.has(node.info.id)
        );
      };
      // Boundaries this seam kept that no piece took stand just before it;
      // the pieces lie beyond them and beyond any insertion standing there.
      const keptEnd = seam - 1 - standingKept;
      let leftEnd = keptEnd;
      while (leftEnd >= 0 && blocking(leftEnd)) leftEnd -= 1;
      let rightStart = seam;
      while (rightStart < out.length && blocking(rightStart)) rightStart += 1;
      // A kept revision's join that meets a pending insertion made inside
      // the record it cut is that insertion's join on the facing side too:
      // accepting the insertion later joins its content with the piece, as
      // accepting both at once does, and rejecting it joins the pieces beyond
      // it, as the seam deferred below does. A removal's seam passes to a
      // pending insertion facing it only where that insertion is a piece of
      // the record the removed content cut (see {@link seamHandsOver}):
      // boundaries it kept standing at the seam go with the join, which
      // later places them inside the record, even to an insertion whose join
      // there is already as deep, made inside that record.
      const handedOver = new Set<number>();
      const piecesBeyond = leftEnd >= 0 && rightStart < out.length;
      const keptJoinHandsOver =
        keptJoinSeams.has(recorded) && standingKept === 0;
      if (keptJoinHandsOver || removalSeams.has(recorded)) {
        const keptStand =
          standingKept > 0 &&
          keptFromIn(seam - standingKept, seam) !== undefined;
        const facingJoinAt = (index: number, side: "before" | "after") => {
          const node = out[index]?.node;
          return node !== undefined && isTrackedWrapper(node)
            ? (node.resolutionJoins?.[side] ?? 0)
            : 0;
        };
        // An insertion joined to nothing on the seam's side takes the join
        // only as the edge of the pieces: where the insertion facing it
        // across the seam is joined there, made inside the record, that one
        // takes the kept markers, as it does once the first is resolved.
        const joinedFacing = {
          after: leftEnd !== keptEnd && facingJoinAt(keptEnd, "after") > 0,
          before: rightStart !== seam && facingJoinAt(seam, "before") > 0,
        };
        const facing = (index: number, side: "before" | "after"): void => {
          const occurrence = out[index];
          if (occurrence === undefined)
            panic("A blocking insertion stands in the list.");
          const facingJoin = facingJoinAt(index, side);
          const yieldsKept =
            facingJoin === 0 &&
            joinedFacing[side === "after" ? "before" : "after"];
          if (
            !seamHandsOver({
              seam: keptJoinHandsOver ? "keptJoin" : "removal",
              piecesBeyond,
              facingJoin,
            })
          )
            return;
          const { node, takesKept } = handJoin(
            occurrence.node,
            side,
            exactDepth,
          );
          const tookKept =
            keptStand &&
            takesKept &&
            !yieldsKept &&
            nameKeptPiece(
              seam - standingKept,
              seam,
              side === "after" ? "before" : "after",
            );
          if (node === occurrence.node && !tookKept) return;
          out[index] = occurrenceOf(node, occurrence);
          handedOver.add(index);
        };
        if (leftEnd !== keptEnd) facing(keptEnd, "after");
        if (rightStart !== seam) facing(seam, "before");
      }
      // A piece the seam joins can still be held in a pending deletion or
      // move source at its edge: it stands only once that is rejected, and
      // goes if it is accepted, so the seam waits on it too.
      const holdsPiece = (index: number, edge: "start" | "end"): boolean => {
        const node = out[index]?.node;
        const other = out[edge === "end" ? rightStart : leftEnd]?.node;
        if (
          node === undefined ||
          other === undefined ||
          (node.type !== "deletion" && node.type !== "moveFrom") ||
          node.resolutionJoins === undefined ||
          resolution.ids.has(node.info.id)
        )
          return false;
        const piece = edge === "end" ? node.content.at(-1) : node.content.at(0);
        // Facing it, the other piece can be held in a pending removal too.
        let facing: InlineNode | undefined = other;
        if (
          (other.type === "deletion" || other.type === "moveFrom") &&
          !resolution.ids.has(other.info.id)
        )
          facing = edge === "end" ? other.content.at(0) : other.content.at(-1);
        return (
          piece !== undefined &&
          facing !== undefined &&
          sameOwnFields(piece, facing)
        );
      };
      const leftHolder =
        leftEnd >= 0 && rightStart < out.length && holdsPiece(leftEnd, "end")
          ? leftEnd
          : undefined;
      const rightHolder =
        leftEnd >= 0 &&
        rightStart < out.length &&
        holdsPiece(rightStart, "start")
          ? rightStart
          : undefined;
      // Insertions whose own removal joins nothing become a join only by
      // waiting on the seam: they wait only with records beyond them that
      // are pieces of the record it cut, alike, as rejecting them joins
      // those. Boundaries standing there are none; with no alike pieces
      // beyond them the facing insertion holds the join. Insertions cutting
      // a record themselves wait to take the seam's depth into that cut.
      const pieceBeyond = (from: number, step: 1 | -1) => {
        let index = from;
        while (isStandingBoundary(out[index]?.node)) index += step;
        return out[index]?.node;
      };
      const beyondLeft = leftEnd < 0 ? undefined : pieceBeyond(leftEnd, -1);
      const beyondRight = pieceBeyond(rightStart, 1);
      const cutsOwn = [
        ...Array.from(
          { length: keptEnd - leftEnd },
          (_, offset) => leftEnd + 1 + offset,
        ),
        ...Array.from(
          { length: rightStart - seam },
          (_, offset) => seam + offset,
        ),
      ].every((index) => {
        const node = out[index]?.node;
        return (
          node !== undefined &&
          isTrackedWrapper(node) &&
          (node.resolutionJoins?.remove ?? 0) > 0
        );
      });
      const mayWait =
        cutsOwn ||
        leftHolder !== undefined ||
        rightHolder !== undefined ||
        (beyondLeft !== undefined &&
          beyondRight !== undefined &&
          !isTrackedWrapper(beyondLeft) &&
          !isTrackedWrapper(beyondRight) &&
          sameOwnFields(beyondLeft, beyondRight));
      // A kept revision's join belongs to the record that revision cut,
      // counting that record's own tracked wrapper as its first level.
      // Insertions made before the revision there are no changes made inside
      // that record but pieces of it: rejecting them takes that wrapper away,
      // so the seam they defer joins the pieces beyond them one level less
      // deep, and not at all when that wrapper was all the revision cut.
      const keptJoinFrom = removalSeams.has(recorded)
        ? undefined
        : keptJoinDates.get(recorded);
      const earlier = (index: number): boolean => {
        const node = out[index]?.node;
        return (
          keptJoinFrom !== undefined &&
          node !== undefined &&
          isTrackedWrapper(node) &&
          (node.info.date ?? "") < keptJoinFrom
        );
      };
      const piecesOnly =
        leftHolder === undefined &&
        rightHolder === undefined &&
        Array.from(
          { length: keptEnd - leftEnd },
          (_, offset) => leftEnd + 1 + offset,
        ).every(earlier) &&
        Array.from(
          { length: rightStart - seam },
          (_, offset) => seam + offset,
        ).every(earlier);
      const deferredDepth = piecesOnly ? exactDepth - 1 : exactDepth;
      if (
        (leftEnd !== keptEnd ||
          rightStart !== seam ||
          leftHolder !== undefined ||
          rightHolder !== undefined) &&
        mayWait &&
        deferredDepth > 0 &&
        leftEnd >= 0 &&
        rightStart < out.length
      ) {
        const blockerIndices = [
          ...(leftHolder === undefined ? [] : [leftHolder]),
          ...Array.from(
            { length: keptEnd - leftEnd },
            (_, offset) => leftEnd + 1 + offset,
          ),
          ...Array.from(
            { length: rightStart - seam },
            (_, offset) => seam + offset,
          ),
          ...(rightHolder === undefined ? [] : [rightHolder]),
        ];
        const blockers = blockerIndices.flatMap((index) => {
          const node = out[index]?.node;
          return node !== undefined && isTrackedWrapper(node)
            ? [node.info.id]
            : [];
        });
        for (const index of blockerIndices) {
          const occurrence = out[index];
          const node = occurrence?.node;
          if (
            occurrence === undefined ||
            node === undefined ||
            !isTrackedWrapper(node) ||
            node.resolutionJoins === undefined
          )
            panic("A deferring seam blocker is a pending change with joins.");
          const groups = [...(node.resolutionJoins.deferredRemove ?? [])];
          let edge: { edge?: "start" | "end" } = {};
          if (index === leftHolder) edge = { edge: "end" };
          if (index === rightHolder) edge = { edge: "start" };
          groups.push({ depth: deferredDepth, blockers, ...edge });
          out[index] = occurrenceOf(
            Object.assign({}, node, {
              resolutionJoins: Object.assign({}, node.resolutionJoins, {
                deferredRemove: canonicalDeferredGroups(groups),
              }),
            }),
            occurrence,
          );
        }
        // The blockers take the kept boundaries standing at the seam: those
        // name the blockers' side, or both, as their cut's piece.
        const blockedBefore = leftEnd !== keptEnd || leftHolder !== undefined;
        const blockedAfter = rightStart !== seam || rightHolder !== undefined;
        let blockedSide: KeptCutPiece = blockedBefore ? "before" : "after";
        if (blockedBefore && blockedAfter) blockedSide = "both";
        if (standingKept > 0)
          nameKeptPiece(seam - standingKept, seam, blockedSide);
        continue;
      }
      // With no piece beyond the pending insertions to defer on, nothing
      // joins now: the insertion facing the kept piece holds its join.
      if (handedOver.size > 0) continue;
    }
    // Later independent edits may make a recorded seam non-alike. Join
    // only the recorded depth and matching fields, preserving those edits.
    let met: InlineNode[];
    // A removal's seam joins the source pieces either side of it. Accepted
    // new content standing there is no such piece: it merges with a
    // neighbour only as far as they are alike, as a direct replacement does.
    const besidePayload =
      leftOccurrence?.payload === true || rightOccurrence?.payload === true;
    if (
      merges.has(recorded) ||
      (exactDepth !== undefined && removalSeams.has(recorded) && besidePayload)
    ) {
      // Accepted new content merges with a plain neighbour it is alike, and
      // records no seam on a neighbour's pending review: it is no piece of it.
      met =
        besidePayload && alikeDepth(left, right) === 0
          ? [left, right]
          : mergeAtSeam(left, right, resolution.emptied);
    } else if (exactDepth === undefined) {
      met = foldAtSeam(left, right, resolution.emptied);
    } else {
      // A resolution, this one or an earlier, can empty an original cut
      // fragment. Its surviving identity folds back before measuring the
      // remaining depth.
      const folded = foldAtCut({
        left,
        right,
        depth: exactDepth,
        emptied: resolution.emptied,
      });
      if (folded.length === 1) met = folded;
      else {
        const held = deferNestedRemovalSeam(left, right, exactDepth);
        const deferred = deferRunReviewSeam({
          left: held[0],
          right: held[1],
          depth: exactDepth,
        });
        const deferredLeft =
          deferred.at(0) ?? panic("A recorded seam retains its left node.");
        const deferredRight =
          deferred.at(1) ?? panic("A recorded seam retains its right node.");
        met =
          mergeLists([deferredLeft], [deferredRight], exactDepth, {
            mode: "asFarAsAlike",
          }) ?? panic("An as-far-as-alike merge always returns its records.");
      }
    }
    if (met.length === 1) {
      const merged = met.at(0) ?? panic("A single merged record exists.");
      for (const [id, into] of rejoinedPieceIds({ left, right, merged }))
        retired.set(id, into);
      const outerEdges: CutEdges = {};
      const first = resolution.cutEdges.get(left)?.first;
      const last = resolution.cutEdges.get(right)?.last;
      if (first !== undefined) outerEdges.first = first;
      if (last !== undefined) outerEdges.last = last;
      resolution.cutEdges.set(merged, outerEdges);
      const leftOwner = leftOccurrence?.owner;
      const rightOwner = rightOccurrence?.owner;
      if (leftOwner !== undefined && rightOwner !== undefined) {
        const leftRoot = sourceOwnerRoot(leftOwner);
        const rightRoot = sourceOwnerRoot(rightOwner);
        if (leftRoot !== rightRoot) {
          if (
            leftRoot.state.type !== "source" ||
            rightRoot.state.type !== "source"
          )
            panic("Only source owners can merge.");
          rightRoot.state = { type: "merged", into: leftRoot };
        }
      }
      out.splice(
        seam - 1,
        2,
        occurrenceOf(
          merged,
          leftOwner === undefined ? rightOccurrence : leftOccurrence,
        ),
      );
    } else {
      if (met.length !== 2)
        panic("A seam retains two roots or joins them into one.");
      const metLeft =
        met.at(0) ?? panic("An unjoined seam retains its left root.");
      const metRight =
        met.at(1) ?? panic("An unjoined seam retains its right root.");
      out.splice(
        seam - 1,
        2,
        occurrenceOf(carrySeam(metLeft, "after", exactDepth), leftOccurrence),
        occurrenceOf(
          carrySeam(metRight, "before", exactDepth),
          rightOccurrence,
        ),
      );
    }
  }
  if (retired.size > 0) {
    const rejoined = withRejoinedBlockers({
      nodes: out.map(({ node }) => node),
      retired,
      resolution,
    });
    for (const [index, node] of rejoined.entries())
      out[index] = occurrenceOf(node, out[index]);
  }
  // Boundaries standing at the list's edge hold no piece of a nested cut: the
  // edge is the first piece's beyond them, as the source edges above are.
  let firstAt = 0;
  while (firstAt < out.length - 1 && isStandingBoundary(out[firstAt]?.node))
    firstAt += 1;
  let lastAt = out.length - 1;
  while (lastAt > 0 && isStandingBoundary(out[lastAt]?.node)) lastAt -= 1;
  const first = out.at(firstAt)?.node;
  const last = out.at(lastAt)?.node;
  const nestedFirst =
    first === undefined ? undefined : resolution.cutEdges.get(first)?.first;
  const nestedLast =
    last === undefined ? undefined : resolution.cutEdges.get(last)?.last;
  if (nestedFirst !== undefined)
    edges.first = Math.max(edges.first ?? 0, nestedFirst);
  if (nestedLast !== undefined)
    edges.last = Math.max(edges.last ?? 0, nestedLast);
  const resolvedNodes = out.map(({ node }) => node);
  return {
    // A paragraph's own list is settled once the joins stand: one can put a
    // pending change beside a marker its paragraph leaves waiting on nothing.
    nodes:
      path.length === 0
        ? resolvedNodes
        : [
            ...settleKeptCuts(resolvedNodes, {
              known: resolution.relocated.nodes,
              rebuilt: (from, to) => carryResolutionState(resolution, from, to),
            }),
          ],
    changed,
    edges,
    ...(enclosing.length === 0 ? {} : { enclosing }),
  };
};

/** A paragraph's review fields once its property changes and mark are resolved; `undefined` when unchanged. */
const resolveParagraphReview = (
  paragraph: Paragraph,
  { ids, decision }: Resolution,
  keepsBreak: boolean,
): ParagraphReviewFields | undefined => {
  const review = reviewFieldsOf(paragraph);
  let changed = false;
  const changes = paragraph.propertyChanges ?? [];
  const selected = (change: (typeof changes)[number]): boolean =>
    ids.has(change.info.id);
  if (changes.some(selected)) {
    changed = true;
    let remaining = changes.filter((change) => !selected(change));
    if (decision === REVISION_DECISIONS.REJECT) {
      const removed = removePropertyChanges(changes, selected);
      if (removed.kind !== "unchanged") {
        remaining = removed.remaining;
      }
      if (removed.kind === "restore") {
        // A property change records paragraph properties: the mark's run properties stay.
        const restored = withMarkFormatting(
          removed.previous,
          paragraph.formatting,
        );
        if (restored === undefined) {
          delete review.formatting;
        } else {
          review.formatting = restored;
        }
      }
    }
    if (remaining.length > 0) {
      review.propertyChanges = remaining;
    } else {
      delete review.propertyChanges;
    }
  }
  if (keepsBreak) {
    changed = true;
    delete review.pPrMark;
  }
  return changed ? review : undefined;
};

/** Revision ids resolution reaches: tracked changes and run property changes in content, and a paragraph's own. */
const reachableIds = (paragraph: Paragraph): Set<number> => {
  const out = new Set<number>();
  for (const change of paragraph.propertyChanges ?? []) out.add(change.info.id);
  if (paragraph.pPrMark !== undefined) out.add(paragraph.pPrMark.info.id);
  if (paragraph.pPrMark?.deletion !== undefined)
    out.add(paragraph.pPrMark.deletion.info.id);
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node)) out.add(node.info.id);
      for (const change of ownRunPropertyChanges(node)) out.add(change.info.id);
      visit(childNodes(node) ?? []);
    }
  };
  visit(paragraph.content);
  return out;
};

type JoinPlan = {
  op: ResolveRevisionOp;
  story: OpStory;
  paraId: string;
  revisionId: number;
  added: boolean;
  depth: number | undefined;
  /** When the paragraph mark the join resolves was made. */
  splitDate: string | undefined;
};

/**
 * Whether the record a paragraph's content starts with, and the one it ends
 * with, are containers resolution emptied. The record is the innermost one at
 * that edge, down the chain of first (or last) children: a join merges alike
 * containers down that chain, so it is the one a join meets.
 */
type EmptiedEnds = { first: boolean; last: boolean };

const NO_EMPTIED_ENDS: EmptiedEnds = { first: false, last: false };

/** The innermost record at one edge of a list, down the chain of first or last children. */
const edgeRecord = (
  nodes: readonly InlineNode[],
  edge: "first" | "last",
): InlineNode | undefined => {
  const at = (list: readonly InlineNode[]): InlineNode | undefined =>
    edge === "first" ? list.at(0) : list.at(-1);
  let node = at(nodes);
  while (node !== undefined) {
    const children = childNodes(node);
    const inner = children === undefined ? undefined : at(children);
    if (inner === undefined) return node;
    node = inner;
  }
  return node;
};

/** Which edges of a list end in a record `emptied` holds. */
const emptiedEndsOf = (
  nodes: readonly InlineNode[],
  emptied: Emptied,
): EmptiedEnds => {
  const first = edgeRecord(nodes, "first");
  const last = edgeRecord(nodes, "last");
  return {
    first: first !== undefined && emptied.has(first),
    last: last !== undefined && emptied.has(last),
  };
};

/**
 * The emptied ends of each paragraph resolution changed, by paragraph id.
 * Staging the inline changes copies their records, so a join finds them by
 * position rather than by identity.
 */
type EmptiedEndsById = Map<string, EmptiedEnds>;

type DeferredSplitSeamOptions = {
  left: readonly InlineNode[];
  right: readonly InlineNode[];
  depth: number;
};
/** Preserve the recorded source seam when later pending insertions stand at it. */
const deferSplitSeam = ({
  left,
  right,
  depth,
}: DeferredSplitSeamOptions): {
  left: readonly InlineNode[];
  right: readonly InlineNode[];
} => {
  if (depth === 0) return { left, right };
  const last = left.at(-1);
  const first = right.at(0);
  const lastChildren = last === undefined ? undefined : childNodes(last);
  const firstChildren = first === undefined ? undefined : childNodes(first);
  if (
    last !== undefined &&
    first !== undefined &&
    lastChildren !== undefined &&
    firstChildren !== undefined &&
    sameOwnFields(last, first)
  ) {
    const inner = deferSplitSeam({
      left: lastChildren,
      right: firstChildren,
      depth: depth - 1,
    });
    if (inner.left === lastChildren && inner.right === firstChildren)
      return { left, right };
    return {
      left: [...left.slice(0, -1), rebuildNode(last, inner.left)],
      right: [rebuildNode(first, inner.right), ...right.slice(1)],
    };
  }
  const blocking = (node: InlineNode | undefined): boolean =>
    node !== undefined &&
    isAddedRevision(node) &&
    isTrackedWrapper(node) &&
    node.resolutionJoins !== undefined;
  let leftEnd = left.length - 1;
  let rightStart = 0;
  while (leftEnd >= 0 && blocking(left.at(leftEnd))) leftEnd -= 1;
  while (rightStart < right.length && blocking(right.at(rightStart)))
    rightStart += 1;
  // Pending revisions can themselves be the two source fragments. Keep
  // those outside the blocker group; only insertions between them defer it.
  let sourcePair = false;
  for (
    let leftIndex = left.length - 1;
    leftIndex > leftEnd && !sourcePair;
    leftIndex -= 1
  ) {
    const leftSource = left.at(leftIndex);
    if (leftSource === undefined) panic("A scanned source fragment exists.");
    for (let rightIndex = 0; rightIndex < rightStart; rightIndex += 1) {
      const rightSource = right.at(rightIndex);
      if (rightSource === undefined) panic("A scanned source fragment exists.");
      if (!sameOwnFields(leftSource, rightSource)) continue;
      leftEnd = leftIndex;
      rightStart = rightIndex;
      sourcePair = true;
      break;
    }
  }
  if (leftEnd !== left.length - 1 || rightStart !== 0) {
    const blockers = [
      ...left.slice(leftEnd + 1),
      ...right.slice(0, rightStart),
    ].flatMap((node) => (isTrackedWrapper(node) ? [node.info.id] : []));
    const deferred = (node: InlineNode): InlineNode => {
      if (!isTrackedWrapper(node) || node.resolutionJoins === undefined)
        panic(
          "Only a provenance-bearing pending insertion blocks this split seam.",
        );
      const groups = [...(node.resolutionJoins.deferredRemove ?? [])];
      if (
        !groups.some(
          (group) =>
            group.depth === depth &&
            group.blockers.length === blockers.length &&
            group.blockers.every((id, index) => id === blockers.at(index)),
        )
      )
        groups.push({ depth, blockers });
      return Object.assign({}, node, {
        resolutionJoins: Object.assign({}, node.resolutionJoins, {
          deferredRemove: groups,
        }),
      });
    };
    return {
      left: left.map((node, index) =>
        index > leftEnd ? deferred(node) : node,
      ),
      right: right.map((node, index) =>
        index < rightStart ? deferred(node) : node,
      ),
    };
  }
  return { left, right };
};

type ResolutionReceiverEffect = {
  type:
    | "paragraph-join"
    | "table-run-in"
    | "block-run-in"
    | "retained-range-boundaries"
    | "section-transfer";
  receiver: Paragraph;
};

type ParagraphMarkPlan = {
  ops: readonly DocumentOp[];
  receiverEffects: readonly ResolutionReceiverEffect[];
  /**
   * The paragraph, ending a block content control, whose join takes an
   * empty paragraph after the control once every other join is done.
   */
  endsControl?: string;
};

/**
 * The bare empty paragraph right after the block content control that the
 * paragraph at `location` ends, with the list holding it: resolving away
 * that paragraph's mark joins it with this one, which has nothing to bring
 * into the control and goes. `undefined` where the paragraph ends no
 * control, or what follows the control holds anything.
 */
const emptyParagraphAfterControl = (
  body: DocumentBody,
  { list, index }: Pick<ParagraphLocation, "list" | "index">,
):
  | { blocks: readonly BlockContent[]; index: number; follower: Paragraph }
  | undefined => {
  let steps = list;
  let position = index;
  for (;;) {
    if (position !== blockListAt(body.content, steps).length - 1)
      return undefined;
    const last = steps.at(-1);
    if (last?.kind !== "blockSdt") return undefined;
    steps = steps.slice(0, -1);
    position = last.block;
    const blocks = blockListAt(body.content, steps);
    const follower = blocks[position + 1];
    if (follower === undefined) continue;
    return follower.type === "paragraph" &&
      follower.content.length === 0 &&
      follower.pPrMark === undefined &&
      follower.sectionProperties === undefined &&
      (follower.propertyChanges?.length ?? 0) === 0
      ? { blocks, index: position + 1, follower }
      : undefined;
  }
};

/**
 * The paragraph a resolved mark's join left its content in: the paragraph
 * itself where it stands (a join keeping the first), else the one its
 * content ran into.
 */
const joinSurvivor = (
  edit: DocumentEdit,
  paraId: string,
  effects: readonly ResolutionReceiverEffect[],
): string | undefined => {
  const removed = new Set(edit.touched.removed.map(idKey));
  const candidates = [
    paraId,
    ...effects.flatMap(({ type, receiver }) =>
      type === "paragraph-join" ||
      type === "table-run-in" ||
      type === "block-run-in"
        ? [receiver.paraId ?? ""]
        : [],
    ),
  ];
  return candidates.find((id) => id !== "" && !removed.has(idKey(id)));
};

/** `edit` reporting every paragraph it removed as joined into `survivor`. */
const withAbsorptions = (
  edit: DocumentEdit,
  survivor: string | undefined,
): DocumentEdit =>
  survivor === undefined || edit.touched.removed.length === 0
    ? edit
    : {
        ...edit,
        touched: {
          ...edit.touched,
          absorbed: [
            ...(edit.touched.absorbed ?? []),
            ...edit.touched.removed
              .filter((id) => idKey(id) !== idKey(survivor))
              .map((absorbed) => ({ absorbed, survivor })),
          ],
        },
      };

const paragraphMarkPlan = (
  ops: DocumentOp[],
  receiverEffects: ResolutionReceiverEffect[] = [],
  endsControl?: string,
) =>
  Result.ok<ParagraphMarkPlan>({
    ops,
    receiverEffects,
    ...(endsControl === undefined ? {} : { endsControl }),
  });

const joinOps = (
  document: Document,
  { op, story, paraId, revisionId, added, depth, splitDate }: JoinPlan,
  ends: EmptiedEndsById,
  /** Table receivers a later paragraph of this resolution already ran into. */
  ranInto: ReadonlySet<string> = new Set(),
  /** Paragraphs whose marks this resolution resolves after this one's. */
  joinsBefore: ReadonlySet<string> = new Set(),
): Result<ParagraphMarkPlan, DocumentOpRefusal> => {
  const body = storyBody(document, story);
  const location = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(paraId),
  );
  if (location === undefined) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        `No paragraph is ${paraId}.`,
      ),
    );
  }
  const { paragraph } = location;
  const blocks = blockListAt(body.content, location.list);
  const previous = blocks[location.index - 1];
  // Range boundaries the block list holds between this paragraph and the
  // next one go into the joined paragraph where the two meet: what stood
  // before them stays before, what stood after stays after.
  let followerIndex = location.index + 1;
  while (
    paragraph.sectionProperties === undefined &&
    blocks[followerIndex] !== undefined &&
    isBlockBoundary(blocks[followerIndex] ?? panic("A block is listed."))
  )
    followerIndex += 1;
  const follower = blocks[followerIndex];
  const standing =
    follower?.type === "paragraph"
      ? blocks.slice(location.index + 1, followerIndex).filter(isBlockBoundary)
      : [];
  const next =
    standing.length > 0 && follower?.type === "paragraph"
      ? { ...follower, content: [...standing, ...follower.content] }
      : blocks[location.index + 1];
  if (next?.type === "paragraph") {
    const outcome = paragraphMarkOutcome({
      follows: "paragraph",
      holdsContent: paragraphHoldsContent(paragraph),
      markWasAdded: added,
      containerStillEndsWithParagraph: true,
      followingTableCanReceiveRunIn: false,
      removedSectionProperties:
        added && paragraph.sectionProperties !== undefined,
      precedingSectionOwner: precedingSectionOwner(previous, paragraph),
      followingHasSectionProperties: next.sectionProperties !== undefined,
    });
    if (outcome.kind !== "join")
      return panic("A following paragraph always receives a join.");
    // The mark that goes takes the paragraph's properties with it: the next
    // paragraph is left, whole, with the first's content before its own.
    const firstEnds = ends.get(idKey(paraId)) ?? NO_EMPTIED_ENDS;
    const nextId = idKey(next.paraId ?? "");
    const nextEnds = ends.get(nextId) ?? NO_EMPTIED_ENDS;
    const emptied = new Set<InlineNode>();
    const marked: [readonly InlineNode[], "first" | "last", boolean][] = [
      [paragraph.content, "first", firstEnds.first],
      [paragraph.content, "last", firstEnds.last],
      [next.content, "first", nextEnds.first],
      [next.content, "last", nextEnds.last],
    ];
    for (const [content, edge, isEmptied] of marked) {
      const record = edgeRecord(content, edge);
      if (isEmptied && record !== undefined) emptied.add(record);
    }
    let merged: InlineNode[];
    if (depth === undefined) {
      merged = added
        ? mergedAtSeam(paragraph.content, next.content, emptied)
        : joinParagraphSeam(paragraph.content, next.content);
    } else {
      const seam = deferSplitSeam(
        depth > 0
          ? {
              ...withSeamBoundariesInside({
                left: paragraph.content,
                right: next.content,
                depth,
              }),
              depth,
            }
          : { left: paragraph.content, right: next.content, depth },
      );
      // Markers kept from content made after the split stood in its seam:
      // they stand in the join's cut, as when the join was resolved first.
      if (depth > 0 && splitDate !== undefined) {
        const cut = { depth, after: splitDate };
        seam.left = withEdgeKeptCutsWidened(seam.left, "end", cut);
        seam.right = withEdgeKeptCutsWidened(seam.right, "start", cut);
      }
      const left = seam.left.at(-1);
      const right = seam.right.at(0);
      if (depth > 0 && left !== undefined && right !== undefined) {
        const folded = foldAtCut({ left, right, depth, emptied });
        const met =
          folded.length === 1
            ? folded
            : (mergeLists([left], [right], depth, { mode: "asFarAsAlike" }) ??
              panic("An as-far-as-alike merge always returns its records."));
        merged = [...seam.left.slice(0, -1), ...met, ...seam.right.slice(1)];
      } else {
        merged = [...seam.left, ...seam.right];
      }
    }
    const keepsFirst = paragraph.pPrMark?.survivor === "first";
    // A record left as it was keeps its identity through the merge; an emptied
    // one folded into its neighbour is rebuilt, and is no longer empty. The
    // joined paragraph is found again under the id it keeps.
    ends.set(
      keepsFirst ? idKey(paraId) : nextId,
      emptiedEndsOf(merged, emptied),
    );
    const mergedContent = asParagraphContent(merged);
    const receivesSection = outcome.sectionDestination === "following";
    let survivor: Paragraph;
    if (keepsFirst)
      survivor = joinedKeepingFirst(
        paragraph,
        next,
        mergedContent,
        receivesSection,
      );
    // An unchanged follower that receives nothing keeps its identity.
    else if (!receivesSection && structurallyEqual(mergedContent, next.content))
      survivor = next;
    else
      survivor = {
        ...next,
        content: mergedContent,
        ...(receivesSection
          ? { sectionProperties: paragraph.sectionProperties }
          : {}),
      };
    const transfersToPrevious =
      outcome.sectionDestination === "preceding" &&
      previous?.type === "paragraph";
    const receiver =
      standing.length > 0 && follower?.type === "paragraph" ? follower : next;
    const receiverEffects: ResolutionReceiverEffect[] = [
      { type: "paragraph-join", receiver },
    ];
    if (transfersToPrevious)
      receiverEffects.push({ type: "section-transfer", receiver: previous });
    if (receivesSection)
      receiverEffects.push({ type: "section-transfer", receiver });
    if (standing.length > 0) {
      const start = transfersToPrevious ? location.index - 1 : location.index;
      return paragraphMarkPlan(
        [
          {
            type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
            story,
            // The joined paragraph stays in the list to anchor the inverse.
            blockId: survivor.paraId ?? paraId,
            expected: blocks,
            blocks: [
              ...blocks.slice(0, start),
              ...(transfersToPrevious
                ? [
                    {
                      ...previous,
                      sectionProperties: paragraph.sectionProperties,
                    },
                  ]
                : []),
              survivor,
              ...blocks.slice(followerIndex + 1),
            ],
          },
        ],
        receiverEffects,
      );
    }
    return paragraphMarkPlan(
      [
        {
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story,
          expected: transfersToPrevious
            ? [previous, paragraph, next]
            : [paragraph, next],
          blocks: transfersToPrevious
            ? [
                { ...previous, sectionProperties: paragraph.sectionProperties },
                survivor,
              ]
            : [survivor],
          ...(paragraph.sectionProperties === undefined
            ? {}
            : { sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE }),
        },
      ],
      receiverEffects,
    );
  }
  // Fail closed: content resolved across a range boundary standing between
  // blocks would move into or out of its range.
  if (next !== undefined && isBlockBoundary(next)) {
    return Result.err(
      new DocumentOpRefusal({
        opType: op.type,
        reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        message: `A range boundary stands between ${paraId} and the next block.`,
        blockedParagraphMarkRevisionId: revisionId,
      }),
    );
  }
  // No paragraph to join: a table or another block follows, or the
  // paragraph ends its container.
  const followsBlock = isOpaqueFollower(next);
  const containerFinal = endsItsContainer(body, location);
  const tableBoundary =
    next?.type === "table"
      ? paragraphTableRunInBoundary(paragraph, next)
      : undefined;
  const tableTarget =
    tableBoundary?.type === "receiver" ? tableBoundary.target : undefined;
  const blockTarget =
    next !== undefined && followsBlock
      ? paragraphBlockRunInTarget(paragraph, next)
      : undefined;
  const runInReceiver = tableTarget?.receiver ?? blockTarget?.receiver;
  const remains = paragraphRemains(paragraph.content);
  let follows: "table" | "block" | "nothing" = "nothing";
  if (next?.type === "table") follows = "table";
  else if (followsBlock) follows = "block";
  // Before a block, what stays decides; resolving the marks one at a time
  // from the first joins the paragraphs before this one, whose marks this
  // resolution resolves too, into it first, so their content counts.
  const joining: Paragraph[] = [];
  for (let index = location.index - 1; follows === "block"; index -= 1) {
    const before = blocks[index];
    if (
      before?.type !== "paragraph" ||
      !joinsBefore.has(idKey(before.paraId ?? ""))
    )
      break;
    joining.unshift(before);
  }
  const chained: Paragraph =
    joining.length === 0
      ? paragraph
      : {
          ...paragraph,
          content: [...joining, paragraph].flatMap(({ content }) => content),
        };
  const outcome = paragraphMarkOutcome({
    follows,
    holdsContent: paragraphHoldsContent(chained),
    markWasAdded: added,
    containerStillEndsWithParagraph:
      follows !== "nothing"
        ? blocks.at(-1)?.type === "paragraph"
        : // Removing the final paragraph leaves the block before it last: in
          // a cell, that paragraph's mark becomes the cell's terminator.
          containerFinal && blocks.at(location.index - 1)?.type === "paragraph",
    followingTableCanReceiveRunIn: tableTarget !== undefined,
    holdsContentOnceAccepted: paragraphHoldsContentOnceAccepted(chained),
    followingBlockCanReceiveRunIn: blockTarget !== undefined,
    needsFollowingCommentHost:
      next?.type === "table" &&
      previous?.type !== "paragraph" &&
      remains.length > 0,
    tableRemovalIsTrackable:
      next?.type !== "table" ||
      (tableBoundary !== undefined &&
        (markBeforeTableResolves({ blocks, boundary: tableBoundary }) ||
          // A story another producer ended with a table already: resolving
          // this mark leaves its end as it was, only the run-in crosses.
          (!location.list.some((step) => step.kind === "tableCell") &&
            location.index > 0 &&
            blocks.at(-1)?.type === "table" &&
            tableBoundary.type === "receiver"))),
    emptyInsertedTableBoundary:
      tableBoundary?.type === "blocked" && tableBoundary.reason === "review"
        ? "pendingReview"
        : "unavailable",
    precededByBlock: location.index > 0,
    removedSectionProperties:
      added && paragraph.sectionProperties !== undefined,
    precedingSectionOwner: precedingSectionOwner(previous, paragraph),
    followingHasSectionProperties: false,
  });
  if (outcome.kind === "refuse") {
    return Result.err(
      refusal(
        op,
        outcome.reason,
        `Resolving the mark of ${paraId} would leave its container without a final paragraph.`,
      ),
    );
  }
  // The paragraph after this one ran into the table or block in this
  // resolution, so this one follows it in: resolving the marks one at a time
  // from the first joins this paragraph to that one before it runs in.
  // Removing it as a paragraph with no content of its own would drop what it
  // holds.
  const continuesRunIn =
    outcome.kind === "remove" &&
    runInReceiver !== undefined &&
    paragraph.content.length > 0 &&
    ranInto.has(idKey(runInReceiver.paraId ?? ""));
  const tableOutcome = continuesRunIn ? "run-in" : outcome.kind;
  // An added endpoint may retire under the preceding mark, including when
  // that mark already states the same section. Explicit replacement owns the
  // endpoint removal; structural table inverses must preserve every section.
  const precedingKeepsSection =
    outcome.sectionDestination === "preceding" ||
    (added &&
      paragraph.sectionProperties !== undefined &&
      precedingSectionOwner(previous, paragraph) === "same");
  if (
    (tableOutcome === "remove" || tableOutcome === "run-in") &&
    !(tableOutcome === "remove" && precedingKeepsSection) &&
    blocks.length > 1 &&
    (next?.type === "table" || followsBlock)
  ) {
    const anchor = blocks.find(
      (block, index): block is Paragraph =>
        index !== location.index && block.type === "paragraph",
    );
    if (anchor === undefined) {
      return Result.err(
        refusal(
          op,
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          `Resolving the mark of ${paraId} leaves no structural inverse anchor.`,
        ),
      );
    }
    if (tableOutcome === "remove") {
      const removal = planParagraphBoundaryRemoval({
        blocks,
        index: location.index,
      });
      if (removal.type === "removed")
        return paragraphMarkPlan(
          [
            {
              type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
              story,
              blockId: anchor.paraId ?? "",
              expected: blocks,
              blocks: removal.blocks,
            },
          ],
          [...removal.effects],
        );
    } else if (blockTarget !== undefined) {
      // The paragraph goes from around the block, and the receiver, where it
      // stands inside the block, takes what it held.
      const { receiver } = blockTarget;
      return paragraphMarkPlan(
        [
          {
            type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
            story,
            blockId: anchor.paraId ?? "",
            expected: blocks,
            blocks: blocks.filter((_, index) => index !== location.index),
          },
          {
            type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
            story,
            expected: [receiver],
            blocks: [
              {
                ...receiver,
                content: joinParagraphSeam(paragraph.content, receiver.content),
              },
            ],
          },
        ],
        [{ type: "block-run-in", receiver }],
      );
    } else {
      if (tableTarget === undefined)
        panic("An eligible table run-in has a validated receiver.");
      const replacement = runParagraphIntoTable(paragraph, tableTarget);
      return paragraphMarkPlan(
        [
          {
            type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
            story,
            blockId: anchor.paraId ?? "",
            expected: blocks,
            blocks: blocks.flatMap((block, index) => {
              if (index === location.index) return [];
              return index === location.index + 1 ? [replacement] : [block];
            }),
          },
        ],
        [{ type: "table-run-in", receiver: tableTarget.receiver }],
      );
    }
  }
  if (outcome.kind === "remove" && blocks.length > 1) {
    const anchor = blocks.find(
      (block, index): block is Paragraph =>
        index !== location.index && block.type === "paragraph",
    );
    const removal = planParagraphBoundaryRemoval({
      blocks,
      index: location.index,
    });
    if (anchor !== undefined && removal.type === "removed") {
      const replacementBlocks = removal.blocks;
      const preceding = replacementBlocks[location.index - 1];
      const receiverEffects: ResolutionReceiverEffect[] = [...removal.effects];
      if (precedingKeepsSection && preceding?.type === "paragraph") {
        const survivor = {
          ...preceding,
          sectionProperties: paragraph.sectionProperties,
        };
        const receiver = blocks[location.index - 1];
        if (receiver?.type !== "paragraph")
          panic("A section transfer retains its preceding paragraph.");
        receiverEffects.push({ type: "section-transfer", receiver });
        return paragraphMarkPlan(
          [
            {
              type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
              story,
              expected: [receiver, paragraph],
              blocks: [survivor],
              sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
            },
          ],
          receiverEffects,
        );
      }
      return paragraphMarkPlan(
        [
          {
            type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
            story,
            blockId: anchor.paraId ?? "",
            expected: blocks,
            blocks: replacementBlocks,
          },
        ],
        receiverEffects,
      );
    }
  }
  const review = reviewFieldsOf(paragraph);
  delete review.pPrMark;
  const markOps: DocumentOp[] = [
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
      story,
      blockId: paraId,
      expected: reviewFieldsOf(paragraph),
      review,
    },
  ];
  // Ending a block content control, the paragraph joins the one after the
  // control: an empty one has nothing to bring in, and goes, once the other
  // joins have settled what follows the control.
  return paragraphMarkPlan(
    markOps,
    [],
    next === undefined &&
      emptyParagraphAfterControl(body, location) !== undefined
      ? paraId
      : undefined,
  );
};

type AbsorbControlFollowerOptions = {
  document: Document;
  story: OpStory;
  paraId: string;
  applyOps: ApplyOps;
};

type ControlFollowerAbsorption = { edit: DocumentEdit; receiver: Paragraph };

/**
 * Remove the empty paragraph after the control the paragraph `paraId` ends,
 * where one still follows once every join of the resolution is done.
 */
const absorbControlFollower = ({
  document,
  story,
  paraId,
  applyOps,
}: AbsorbControlFollowerOptions): Result<
  ControlFollowerAbsorption | undefined,
  DocumentOpRefusal
> => {
  const body = storyBody(document, story);
  const location = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(paraId),
  );
  if (location === undefined) return Result.ok(undefined);
  const absorbed = emptyParagraphAfterControl(body, location);
  const anchor = absorbed?.blocks.find(
    (block, index): block is Paragraph =>
      index !== absorbed.index && block.type === "paragraph",
  );
  if (absorbed === undefined || anchor === undefined)
    return Result.ok(undefined);
  const removed = applyOps(document, [
    {
      type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
      story,
      blockId: anchor.paraId ?? "",
      expected: absorbed.blocks,
      blocks: absorbed.blocks.filter((_, index) => index !== absorbed.index),
    },
  ]);
  return removed.isErr()
    ? Result.err(removed.error)
    : Result.ok({ edit: removed.value, receiver: absorbed.follower });
};

/** The main story's own content: the body's blocks, without its comments or section view. */
const storyIdentityKeys = (document: Document, story: OpStory): Set<string> =>
  new Set(identityKeysIn(storyBody(document, story).content));

type AcceptedClosedBoundaryOptions = {
  nodes: readonly InlineNode[];
  edge: "first" | "last";
  resolution: Resolution;
};
/** A closed accepted payload ends the original cut at its containing depth. */
const acceptedClosedBoundary = ({
  nodes,
  edge,
  resolution,
}: AcceptedClosedBoundaryOptions): number | undefined => {
  if (resolution.decision !== REVISION_DECISIONS.ACCEPT) return undefined;
  let current = nodes;
  let depth = 0;
  while (current.length > 0) {
    const node = edge === "first" ? current.at(0) : current.at(-1);
    if (node === undefined) return undefined;
    if (resolution.acceptedClosed[edge].has(node)) return depth;
    if (
      isTrackedWrapper(node) &&
      isAddedRevision(node) &&
      resolution.ids.has(node.info.id) &&
      node.resolutionJoins !== undefined
    ) {
      const fitting =
        edge === "first"
          ? node.resolutionJoins.after
          : node.resolutionJoins.before;
      if (fitting === 0) return depth;
    }
    const children = childNodes(node);
    if (children === undefined) return undefined;
    current = children;
    depth += 1;
  }
  return undefined;
};

/**
 * The depth at which a rejection removes the record standing at a
 * paragraph's edge, where a split's cut left a piece. An insertion standing
 * inside the seam itself (one that cut records of its own) is passed over:
 * removing it exposes the seam rather than the piece. `undefined` when the
 * edge keeps a record down to its leaves.
 */
/**
 * How many of the `depth` levels a split cut, read down the records at its
 * seam, are pieces of one tracked change on both sides that this resolution
 * unwraps, keeping their content: an accepted insertion or move destination,
 * a rejected deletion or move source.
 */
const unwrappedCutLevels = ({
  left,
  right,
  depth,
  resolution,
  splitDate,
}: {
  left: readonly InlineNode[];
  right: readonly InlineNode[];
  depth: number;
  resolution: Resolution;
  /** The date of the split whose cut this reads, when it has one. */
  splitDate: string | undefined;
}): number => {
  const unwraps = (node: InlineNode): node is TrackedWrapper =>
    isTrackedWrapper(node) &&
    resolution.ids.has(node.info.id) &&
    !removesContent(node, resolution.decision);
  // Content this resolution removes that was inserted after the split, or
  // into the seam itself, is no piece the split left (see
  // {@link removedClosedBoundary}).
  const insideSeam = (node: InlineNode): boolean =>
    isTrackedWrapper(node) &&
    isAddedRevision(node) &&
    resolution.ids.has(node.info.id) &&
    removesContent(node, resolution.decision) &&
    ((node.resolutionJoins?.remove ?? 0) > 0 ||
      (splitDate !== undefined && (node.info.date ?? "") > splitDate));
  const edgeOf = (
    nodes: readonly InlineNode[] | undefined,
    edge: "first" | "last",
  ): InlineNode | undefined =>
    (edge === "first" ? nodes : nodes?.toReversed())?.find(
      (node) => !insideSeam(node),
    );
  let levels = 0;
  let before = edgeOf(left, "last");
  let after = edgeOf(right, "first");
  for (let level = 0; level < depth; level += 1) {
    if (before === undefined || after === undefined) break;
    if (unwraps(before) && unwraps(after) && sameStamp(before, after))
      levels += 1;
    before = edgeOf(childNodes(before), "last");
    after = edgeOf(childNodes(after), "first");
  }
  return levels;
};

/**
 * `nodes` with the cut sides of the tracked changes an inserted break split
 * closed, where this resolution unwraps them. A split inside a change leaves
 * a piece of it either side of the new mark, each carrying the change's
 * joins, so the two still merge back as one record when the break goes. The
 * join a piece holds on the side facing the mark is a copy of the one at the
 * change's far edge: the change's content met nothing there. Unwrapped, the
 * piece records no seam on that side, as unwrapping the change after its
 * pieces met records none inside it. `edge` is the side of `nodes` facing
 * the mark, `facing` the content across it, `depth` the levels the split cut.
 */
const closedSplitCutSides = ({
  nodes,
  facing,
  edge,
  depth,
  resolution,
  splitDate,
}: {
  nodes: readonly InlineNode[];
  facing: readonly InlineNode[];
  edge: "first" | "last";
  depth: number;
  resolution: Resolution;
  splitDate: string;
}): readonly InlineNode[] => {
  // Content this resolution removes that was inserted after the split, or
  // into the seam itself, is no piece the split left (as in
  // {@link unwrappedCutLevels}).
  const insideSeam = (node: InlineNode): boolean =>
    isTrackedWrapper(node) &&
    isAddedRevision(node) &&
    resolution.ids.has(node.info.id) &&
    removesContent(node, resolution.decision) &&
    ((node.resolutionJoins?.remove ?? 0) > 0 ||
      (node.info.date ?? "") > splitDate);
  const edgeAt = (
    list: readonly InlineNode[],
    side: "first" | "last",
  ): number =>
    side === "first"
      ? list.findIndex((node) => !insideSeam(node))
      : list.findLastIndex((node) => !insideSeam(node));
  const side = edge === "last" ? "after" : "before";
  const close = (
    list: readonly InlineNode[],
    across: readonly InlineNode[],
    level: number,
  ): readonly InlineNode[] => {
    if (level >= depth) return list;
    const at = edgeAt(list, edge);
    const node = list.at(at);
    const other = across.at(edgeAt(across, edge === "last" ? "first" : "last"));
    if (at < 0 || node === undefined || other === undefined) return list;
    const tracked = isTrackedWrapper(node) && isTrackedWrapper(other);
    const pieces = tracked
      ? sameStamp(node, other) && (node.info.date ?? "") < splitDate
      : !isTrackedWrapper(node) &&
        !isTrackedWrapper(other) &&
        sameOwnFields(node, other);
    if (!pieces) return list;
    const children = childNodes(node);
    const inner =
      children === undefined
        ? undefined
        : close(children, childNodes(other) ?? [], level + 1);
    let next =
      inner === undefined || inner === children
        ? node
        : rebuildNode(node, inner);
    const joins = isTrackedWrapper(next) ? next.resolutionJoins : undefined;
    if (
      isTrackedWrapper(next) &&
      joins !== undefined &&
      joins[side] > 0 &&
      joins.plainSides?.includes(side) !== true &&
      resolution.ids.has(next.info.id) &&
      !removesContent(next, resolution.decision)
    )
      next = {
        ...next,
        resolutionJoins:
          side === "after" ? { ...joins, after: 0 } : { ...joins, before: 0 },
      };
    return next === node ? list : list.with(at, next);
  };
  return close(nodes, facing, 0);
};

const removedClosedBoundary = ({
  nodes,
  edge,
  resolution,
  splitDate,
}: AcceptedClosedBoundaryOptions & {
  /** The date of the split whose cut this reads, when it has one. */
  splitDate: string | undefined;
}): number | undefined => {
  if (resolution.decision !== REVISION_DECISIONS.REJECT) return undefined;
  const removed = (node: InlineNode): boolean =>
    isTrackedWrapper(node) &&
    isAddedRevision(node) &&
    resolution.ids.has(node.info.id);
  // Content inserted after the split, or into the seam itself, is no piece
  // the split left: removing it exposes the seam rather than the piece.
  const insideSeam = (node: InlineNode): boolean =>
    removed(node) &&
    isTrackedWrapper(node) &&
    ((node.resolutionJoins?.remove ?? 0) > 0 ||
      (splitDate !== undefined && (node.info.date ?? "") > splitDate));
  let current = nodes;
  let depth = 0;
  while (current.length > 0) {
    const ordered = edge === "first" ? current : current.toReversed();
    const node = ordered.find((candidate) => !insideSeam(candidate));
    if (node === undefined) return undefined;
    if (removed(node)) return depth;
    const children = childNodes(node);
    if (children === undefined) return undefined;
    current = children;
    depth += 1;
  }
  return undefined;
};

type RevisionResolutionPlan = {
  edit: DocumentEdit;
  receiverEffects: readonly ResolutionReceiverEffect[];
};

type RevisionResolutionPlanOptions = {
  document: Document;
  op: ResolveRevisionOp;
  applyOps: ApplyOps;
};

/**
 * A mark holding an insertion and a later deletion of it, settled to the one
 * record resolution then decides, with the revision ids left to resolve:
 * - accepting the deletion takes the break, whatever becomes of the
 *   insertion: the mark is the deletion alone, which the resolution accepts;
 * - accepting the insertion alone leaves the deletion pending;
 * - rejecting the insertion takes the break with its deletion: the mark is
 *   the insertion alone, which the resolution rejects;
 * - rejecting the deletion alone leaves the insertion pending.
 */
const settleTwoRecordMarks = (
  document: Document,
  op: ResolveRevisionOp,
): { ops: DocumentOp[]; ids: number[] } => {
  const ids = new Set(op.revisionIds);
  const ops: DocumentOp[] = [];
  for (const { paragraph } of storyParagraphs(storyBody(document, op.story))) {
    const mark = paragraph.pPrMark;
    if (mark?.deletion === undefined) continue;
    const insertion = mark.info.id;
    const deletion = mark.deletion.info.id;
    if (!ids.has(insertion) && !ids.has(deletion)) continue;
    const review = reviewFieldsOf(paragraph);
    if (op.decision === REVISION_DECISIONS.ACCEPT) {
      review.pPrMark = { kind: "del", info: mark.deletion.info };
      ids.delete(insertion);
    } else {
      const { deletion: _deleted, ...added } = mark;
      review.pPrMark = added;
      ids.delete(deletion);
    }
    ops.push({
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
      story: op.story,
      blockId: paragraph.paraId ?? "",
      expected: reviewFieldsOf(paragraph),
      review,
    });
  }
  return { ops, ids: op.revisionIds.filter((id) => ids.has(id)) };
};

/**
 * Resolve tracked changes by revision id, applied as the primitive
 * operations it expands to. An id no record in the story carries is
 * skipped; one carried by a record resolution does not reach yet (a table,
 * cell or section change, a record inside a field) is refused.
 */
export const planRevisionResolution = (
  options: RevisionResolutionPlanOptions,
): Result<RevisionResolutionPlan, DocumentOpRefusal> => {
  const { document, op, applyOps } = options;
  const sections = settleSectionRecords(document, op);
  if (sections.ops.length === 0) return planMarkResolution(options);
  const settled = applyOps(document, sections.ops);
  if (settled.isErr()) return Result.err(settled.error);
  const rest = planMarkResolution({
    document: settled.value.document,
    op: { ...op, revisionIds: sections.ids },
    applyOps,
  });
  if (rest.isErr()) return Result.err(rest.error);
  return Result.ok({
    edit: combineEdits(document, [settled.value, rest.value.edit]),
    receiverEffects: rest.value.receiverEffects,
  });
};

/**
 * The records of the body's sections the ids select, resolved through the
 * section calculation every reviewer applies, with the revision ids left to
 * resolve. A section's records travel on the paragraph that ends it; the
 * final section's are resolved by section review.
 */
const settleSectionRecords = (
  document: Document,
  op: ResolveRevisionOp,
): { ops: DocumentOp[]; ids: number[] } => {
  if (
    op.story !== OP_STORIES.MAIN ||
    !Object.values(REVISION_DECISIONS).includes(op.decision)
  )
    return { ops: [], ids: [...op.revisionIds] };
  const ids = new Set(op.revisionIds);
  const ops: DocumentOp[] = [];
  for (const [sectionIndex, { properties }] of sectionEnds(
    document.package.document,
  ).entries()) {
    let resolved = properties;
    for (const { info } of properties.propertyChanges ?? []) {
      if (!ids.has(info.id)) continue;
      const next = resolveSectionChange({
        current: resolved,
        decision: op.decision,
        id: info.id,
      });
      if (next === null) continue;
      resolved = next;
      ids.delete(info.id);
    }
    if (resolved !== properties)
      ops.push({
        type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
        sectionIndex,
        patch: sectionResolutionPatch({ current: properties, next: resolved }),
      });
  }
  return { ops, ids: op.revisionIds.filter((id) => ids.has(id)) };
};

/** Paragraph marks, inline records, property records and table rows. */
const planMarkResolution = (
  options: RevisionResolutionPlanOptions,
): Result<RevisionResolutionPlan, DocumentOpRefusal> => {
  const { document, op, applyOps } = options;
  const settled = settleTwoRecordMarks(document, op);
  if (settled.ops.length === 0) return planSingleRecordResolution(options);
  const marks = applyOps(document, settled.ops);
  if (marks.isErr()) return Result.err(marks.error);
  const rest = planSingleRecordResolution({
    document: marks.value.document,
    op: { ...op, revisionIds: settled.ids },
    applyOps,
  });
  if (rest.isErr()) return Result.err(rest.error);
  return Result.ok({
    edit: combineEdits(document, [marks.value, rest.value.edit]),
    receiverEffects: rest.value.receiverEffects,
  });
};

const planSingleRecordResolution = ({
  document,
  op,
  applyOps,
}: RevisionResolutionPlanOptions): Result<
  RevisionResolutionPlan,
  DocumentOpRefusal
> => {
  if (!Object.values(REVISION_DECISIONS).includes(op.decision)) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "A decision accepts or rejects.",
      ),
    );
  }
  if (!op.revisionIds.every((id) => Number.isInteger(id))) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        "A revision id is an integer.",
      ),
    );
  }
  const finalSectionRevision = storyBody(
    document,
    op.story,
  ).finalSectionProperties?.propertyChanges?.find(({ info }) =>
    op.revisionIds.includes(info.id),
  );
  if (finalSectionRevision !== undefined)
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        `Final section revision ${String(finalSectionRevision.info.id)} requires section review.`,
      ),
    );
  const present = storyIdentityKeys(document, op.story);
  const ids = new Set(
    op.revisionIds.filter((id) =>
      present.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })),
    ),
  );
  if (ids.size === 0) {
    return Result.ok({
      edit: {
        document,
        inverse: [],
        touched: { modified: [], inserted: [], removed: [] },
      },
      receiverEffects: [],
    });
  }
  let resolution = createResolution(ids, op.decision);
  const rejectedNoteReferences =
    op.decision === REVISION_DECISIONS.REJECT
      ? removedNoteReferences(document, op, ids)
      : [];
  const paragraphs = storyParagraphs(storyBody(document, op.story));
  const reachable = new Set([
    ...paragraphs.flatMap(({ paragraph }) => [...reachableIds(paragraph)]),
    ...reachableRowIds(document, op.story),
  ]);
  const unreachable = [...ids].find((id) => !reachable.has(id));
  if (unreachable !== undefined) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        `Revision ${unreachable} is not an inline change, property change, row change or paragraph mark.`,
      ),
    );
  }

  const blockerIds = seamHolderIds(
    paragraphs.flatMap(({ paragraph }) => paragraph.content),
  );
  const validProvenance = (nodes: readonly InlineNode[]): boolean =>
    nodes.every(
      (node) =>
        validRetainedIdentities(node) &&
        (!isTrackedWrapper(node) ||
          node.resolutionJoins?.acceptance === undefined ||
          node.resolutionJoins.acceptance ===
            INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS) &&
        (!isTrackedWrapper(node) ||
          node.resolutionJoins?.deferredRemove === undefined ||
          node.resolutionJoins.deferredRemove.every((group) =>
            group.blockers.every((id) => blockerIds.has(id)),
          )) &&
        validProvenance(childNodes(node) ?? []),
    );
  if (
    paragraphs.some(({ paragraph }) => {
      const depth = paragraph.pPrMark?.resolutionJoin;
      return (
        !validProvenance(paragraph.content) ||
        (depth !== undefined &&
          (!Number.isInteger(depth) || depth < 0 || depth > MAX_REVISION_ID))
      );
    })
  ) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "Resolution provenance must name valid source identity slots and cut depths.",
      ),
    );
  }

  const inline: DocumentOp[] = [];
  const joins: JoinPlan[] = [];
  const receiverEffects: ResolutionReceiverEffect[] = [];
  const emptiedEnds: EmptiedEndsById = new Map();
  const resolvedContents = new Map<string, readonly InlineNode[]>();
  const resolvedCutEdges = new Map<string, CutEdges>();
  const prepareContents = () => {
    const prepared = paragraphs.map(({ paragraph }) => ({
      paragraph,
      content: prepareDeferredRemovals(paragraph.content, resolution),
    }));
    // The pieces an inserted break split a change into face it with copies
    // of the change's far joins: closed where this resolution unwraps them.
    const at = new Map(
      prepared.map((entry, index) => [entry.paragraph, index] as const),
    );
    const blocks = storyBody(document, op.story);
    for (const location of paragraphs) {
      const mark = location.paragraph.pPrMark;
      if (
        mark === undefined ||
        !markWasAdded(mark.kind) ||
        (mark.resolutionJoin ?? 0) === 0
      )
        continue;
      const next = blockListAt(blocks.content, location.list).at(
        location.index + 1,
      );
      if (next?.type !== "paragraph") continue;
      const left = prepared[at.get(location.paragraph) ?? -1];
      const right = prepared[at.get(next) ?? -1];
      if (left === undefined || right === undefined) continue;
      const shared = {
        depth: mark.resolutionJoin ?? 0,
        resolution,
        splitDate: mark.info.date ?? "",
      };
      const trailing = closedSplitCutSides({
        ...shared,
        nodes: left.content,
        facing: right.content,
        edge: "last",
      });
      const leading = closedSplitCutSides({
        ...shared,
        nodes: right.content,
        facing: left.content,
        edge: "first",
      });
      left.content = trailing;
      right.content = leading;
    }
    return prepared;
  };
  let preparedContents = prepareContents();
  const resolvedChanged = new Set<string>();
  type SelectionClosure =
    | { type: "initial" }
    | { type: "replaying"; domain: ReadonlySet<string>; remaining: number };
  let closure: SelectionClosure = { type: "initial" };
  for (;;) {
    const selectedBefore = ids.size;
    const keptDeletionPaths =
      op.decision === REVISION_DECISIONS.ACCEPT
        ? keptDeletionBoundaryPaths({
            body: storyBody(document, op.story),
            removes: ({ ancestors }) =>
              ancestors.some(
                (ancestor) =>
                  isTrackedWrapper(ancestor) &&
                  ids.has(ancestor.info.id) &&
                  removesContent(ancestor, op.decision),
              ),
          })
        : new Map<Paragraph, Set<string>>();
    resolution.pairedComments = pairedCommentRanges(
      preparedContents.flatMap(({ content }) => content),
    );
    // Acceptance removes selected occurrences through the boundary plan. An ID
    // can also name a standing pair, which an unrelated removed pair cannot erase.
    const removedCommentIds =
      op.decision === REVISION_DECISIONS.REJECT
        ? removedCommentRangeIds(
            preparedContents.flatMap(({ content }) => content),
            resolution,
          )
        : NO_COMMENT_RANGES;
    for (const { paragraph, content: source } of preparedContents) {
      const paraId = paragraph.paraId ?? "";
      // A replacement's paragraph break can stand between a removed prefix and
      // the piece that takes its identities: they arrive here.
      const handed = restoreOutstandingIdentities(source, resolution);
      const prepared =
        handed === source
          ? source
          : prepareDeferredRemovals(handed, resolution);
      resolution.keptDeletionPaths =
        keptDeletionPaths.get(paragraph) ?? NO_BOUNDARY_PATHS;
      const resolved = resolveList({ nodes: prepared, resolution, path: [] });
      resolvedContents.set(
        idKey(paraId),
        withoutCommentAnchors(resolved.nodes, removedCommentIds),
      );
      if (resolved.edges !== undefined)
        resolvedCutEdges.set(idKey(paraId), resolved.edges);
      if (resolved.changed || prepared !== paragraph.content)
        resolvedChanged.add(idKey(paraId));
    }
    if (op.decision === REVISION_DECISIONS.REJECT) break;
    if (ids.size < selectedBefore)
      panic("Revision selection is monotone across identity handoffs.");
    if (ids.size === selectedBefore) break;
    // The same handoff owner discovered more selected source identities.
    // Discard this whole pass, including boundary decisions made before that
    // handoff, and publish only a pass whose selection is already closed.
    switch (closure.type) {
      case "initial": {
        const domain = new Set([
          ...present,
          ...reservedIdentityKeysIn(storyBody(document, op.story)),
        ]);
        closure = {
          type: "replaying",
          domain,
          remaining: domain.size - selectedBefore,
        };
        break;
      }
      case "replaying":
        closure = {
          type: "replaying",
          domain: closure.domain,
          remaining: closure.remaining - 1,
        };
        break;
      default: {
        const exhaustive: never = closure;
        return exhaustive;
      }
    }
    // Every expanding pass selects at least one member of this finite source
    // domain. No identity is minted here, so at most domain.size - initialSize
    // expansions precede the final stable pass.
    const { domain, remaining } = closure;
    if (
      remaining <= 0 ||
      [...ids].some(
        (id) => !domain.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })),
      )
    )
      panic("Revision selection exceeded its finite source identity domain.");
    resolution = createResolution(ids, op.decision);
    resolvedContents.clear();
    resolvedCutEdges.clear();
    resolvedChanged.clear();
    preparedContents = prepareContents();
  }
  const noteReferences =
    op.decision === REVISION_DECISIONS.ACCEPT
      ? removedNoteReferences(document, op, ids)
      : rejectedNoteReferences;
  // Boundaries kept from removed content can leave a range ending before it starts.
  const normalizedRangeContents = wellFormedRanges({
    contents: resolvedContents,
    order: preparedContents.map(({ paragraph }) =>
      idKey(paragraph.paraId ?? ""),
    ),
    relocated: resolution.relocated,
  });
  for (const { paragraph } of preparedContents) {
    const key = idKey(paragraph.paraId ?? "");
    const content = normalizedRangeContents.get(key);
    if (content === undefined) continue;
    resolvedContents.set(key, content);
    resolvedChanged.add(key);
    receiverEffects.push({
      type: "retained-range-boundaries",
      receiver: paragraph,
    });
  }
  // A blocker the resolution emptied away, unnamed, leaves the story too.
  const standing = seamHolderIds([...resolvedContents.values()].flat());
  for (const [key, content] of resolvedContents) {
    const settled = settleDeferredRemovals(content, {
      fateOf: (id) => (standing.has(id) ? "pending" : "removed"),
    });
    if (settled === content) continue;
    resolvedContents.set(key, settled);
    resolvedChanged.add(key);
  }
  for (const { paragraph } of preparedContents) {
    const paraId = paragraph.paraId ?? "";
    const resolvedNodes =
      resolvedContents.get(idKey(paraId)) ??
      panic("Every source paragraph has a resolved content result.");
    if (resolvedChanged.has(idKey(paraId))) {
      emptiedEnds.set(
        idKey(paraId),
        emptiedEndsOf(resolvedNodes, resolution.emptied),
      );
      inline.push({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: op.story,
        blockId: paraId,
        expected: paragraph.content,
        content: asParagraphContent(resolvedNodes),
      });
    }
  }
  const retiredCutDepths = new Map<string, number>();
  const body = storyBody(document, op.story);
  type FollowingCuts = { leadingCut: number; chainCut: number };
  const followingCutsByList = new WeakMap<
    readonly BlockContent[],
    Map<number, FollowingCuts>
  >();
  const cutsFollowing = (
    list: readonly BlockContent[],
    start: number,
  ): FollowingCuts => {
    let known = followingCutsByList.get(list);
    if (known === undefined) {
      known = new Map();
      followingCutsByList.set(list, known);
    }
    const visited: { index: number; cuts: FollowingCuts }[] = [];
    let suffix: FollowingCuts = { leadingCut: 0, chainCut: 0 };
    for (let index = start; index < list.length; index += 1) {
      const cached = known.get(index);
      if (cached !== undefined) {
        suffix = cached;
        break;
      }
      const candidate = list.at(index);
      if (candidate?.type !== "paragraph") break;
      const key = idKey(candidate.paraId ?? "");
      const firstCut = resolvedCutEdges.get(key)?.first ?? 0;
      const mark = candidate.pPrMark;
      if (
        (resolvedContents.get(key)?.length ?? 0) !== 0 ||
        mark === undefined ||
        !markWasAdded(mark.kind)
      ) {
        suffix = { leadingCut: firstCut, chainCut: 0 };
        known.set(index, suffix);
        break;
      }
      // An empty inserted paragraph can carry part of a cut to a later join.
      const chainCut =
        op.decision === REVISION_DECISIONS.REJECT && ids.has(mark.info.id)
          ? (mark.resolutionJoin ?? 0)
          : 0;
      visited.push({
        index,
        cuts: { leadingCut: Math.max(firstCut, chainCut), chainCut },
      });
    }
    for (const { index, cuts } of visited.reverse()) {
      suffix = {
        leadingCut: Math.max(cuts.leadingCut, suffix.leadingCut),
        chainCut: Math.max(cuts.chainCut, suffix.chainCut),
      };
      known.set(index, suffix);
    }
    return suffix;
  };
  for (const location of paragraphs) {
    const paragraph = location.paragraph;
    const mark = paragraph.pPrMark;
    if (
      mark === undefined ||
      !markWasAdded(mark.kind) ||
      mark.resolutionJoin === undefined
    )
      continue;
    const trailing = acceptedClosedBoundary({
      nodes:
        resolvedContents.get(idKey(paragraph.paraId ?? "")) ??
        panic("Every source paragraph has a resolved content result."),
      edge: "last",
      resolution,
    });
    const list = blockListAt(body.content, location.list);
    const next = list.at(location.index + 1);
    const leading =
      next?.type === "paragraph"
        ? acceptedClosedBoundary({
            nodes:
              resolvedContents.get(idKey(next.paraId ?? "")) ??
              panic("Every following paragraph has a resolved content result."),
            edge: "first",
            resolution,
          })
        : undefined;
    const { leadingCut, chainCut } = cutsFollowing(list, location.index + 1);
    // A level of the cut held by pieces of one tracked change this
    // resolution unwraps is a level no more: their content meets uncut.
    const unwrapped =
      next?.type === "paragraph"
        ? unwrappedCutLevels({
            left: paragraph.content,
            right: next.content,
            depth: mark.resolutionJoin,
            resolution,
            splitDate: mark.info.date,
          })
        : 0;
    let depth = Math.max(
      mark.resolutionJoin - unwrapped,
      resolvedCutEdges.get(idKey(paragraph.paraId ?? ""))?.last ?? 0,
      leadingCut,
    );
    if (trailing !== undefined) depth = Math.min(depth, trailing);
    if (leading !== undefined) depth = Math.min(depth, leading);
    // Where this rejection removes the pieces the cut left at both edges, the
    // cut ends at the shallower of the two depths: no pieces stand on both
    // sides below it to join, as when they were removed earlier.
    const removedTrailing = removedClosedBoundary({
      nodes: paragraph.content,
      edge: "last",
      resolution,
      splitDate: mark.info.date,
    });
    const removedLeading =
      next?.type === "paragraph"
        ? removedClosedBoundary({
            nodes: next.content,
            edge: "first",
            resolution,
            splitDate: mark.info.date,
          })
        : undefined;
    if (removedTrailing !== undefined && removedLeading !== undefined)
      depth = Math.min(depth, removedTrailing, removedLeading);
    // The pieces the chain carries the cut to stand beyond the emptied
    // paragraph, so its removed edge does not end that cut.
    if (chainCut > depth) depth = Math.min(chainCut, trailing ?? chainCut);
    if (depth !== mark.resolutionJoin)
      retiredCutDepths.set(idKey(paragraph.paraId ?? ""), depth);
  }
  for (const location of paragraphs) {
    const { paragraph } = location;
    const paraId = paragraph.paraId ?? "";
    const mark = paragraph.pPrMark;
    const markResolved = mark !== undefined && ids.has(mark.info.id);
    const added = mark !== undefined && markWasAdded(mark.kind);
    const keepsBreak =
      markResolved && added === (op.decision === REVISION_DECISIONS.ACCEPT);
    if (markResolved && !keepsBreak) {
      let depth = retiredCutDepths.get(idKey(paraId)) ?? mark.resolutionJoin;
      if (
        added &&
        depth !== undefined &&
        op.decision === REVISION_DECISIONS.REJECT
      ) {
        const next = blockListAt(body.content, location.list).at(
          location.index + 1,
        );
        // A deletion can have cut the source run before a later paragraph
        // split. Rejecting its wrapper exposes that source seam at the edge;
        // carry the recorded depth across the removed inserted break.
        depth = Math.max(
          depth,
          resolvedCutEdges.get(idKey(paraId))?.last ?? 0,
          next?.type === "paragraph"
            ? (resolvedCutEdges.get(idKey(next.paraId ?? ""))?.first ?? 0)
            : 0,
        );
      }
      joins.push({
        op,
        story: op.story,
        paraId,
        revisionId: mark.info.id,
        added,
        depth,
        splitDate: mark.info.date,
      });
    }
    let review = resolveParagraphReview(paragraph, resolution, keepsBreak);
    const retiredDepth = retiredCutDepths.get(idKey(paraId));
    if (retiredDepth !== undefined) {
      const remainingMark =
        review === undefined ? paragraph.pPrMark : review.pPrMark;
      if (remainingMark !== undefined) {
        review ??= reviewFieldsOf(paragraph);
        review.pPrMark = Object.assign({}, remainingMark, {
          resolutionJoin: retiredDepth,
        });
      }
    }
    if (review !== undefined) {
      inline.push({
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
        story: op.story,
        blockId: paraId,
        expected: reviewFieldsOf(paragraph),
        review,
      });
    }
  }

  const staged = applyOps(document, inline);
  if (staged.isErr()) {
    return Result.err(staged.error);
  }
  const rows = resolveTableRows({
    document: staged.value.document,
    op:
      op.decision === REVISION_DECISIONS.ACCEPT
        ? { ...op, revisionIds: [...ids] }
        : op,
    applyOps,
  });
  if (rows.isErr()) return Result.err(rows.error);
  const removedParagraphs = new Set(rows.value.touched.removed.map(idKey));
  const edits: DocumentEdit[] = [staged.value, rows.value];
  let current = rows.value.document;
  const ranInto = new Set<string>();
  const joinsBefore = new Set(joins.map(({ paraId }) => idKey(paraId)));
  const controlEnds: string[] = [];
  for (const join of joins.toReversed()) {
    joinsBefore.delete(idKey(join.paraId));
    if (removedParagraphs.has(idKey(join.paraId))) continue;
    const planned = joinOps(current, join, emptiedEnds, ranInto, joinsBefore);
    if (planned.isErr()) {
      return Result.err(planned.error);
    }
    receiverEffects.push(...planned.value.receiverEffects);
    for (const effect of planned.value.receiverEffects)
      if (effect.type === "table-run-in" || effect.type === "block-run-in")
        ranInto.add(idKey(effect.receiver.paraId ?? ""));
    const joined = applyOps(current, planned.value.ops);
    if (joined.isErr()) {
      return Result.err(joined.error);
    }
    const absorbing = withAbsorptions(
      joined.value,
      joinSurvivor(joined.value, join.paraId, planned.value.receiverEffects),
    );
    edits.push(absorbing);
    current = joined.value.document;
    // A paragraph ending a control that an earlier one now absorbs (a join
    // keeping the first) hands the control's end to that survivor.
    for (const { absorbed, survivor } of absorbing.touched.absorbed ?? [])
      for (const [index, end] of controlEnds.entries())
        if (idKey(end) === idKey(absorbed)) controlEnds[index] = survivor;
    if (planned.value.endsControl !== undefined)
      controlEnds.push(planned.value.endsControl);
  }
  for (const paraId of controlEnds) {
    const absorbed = absorbControlFollower({
      document: current,
      story: op.story,
      paraId,
      applyOps,
    });
    if (absorbed.isErr()) return Result.err(absorbed.error);
    if (absorbed.value === undefined) continue;
    const { edit, receiver } = absorbed.value;
    receiverEffects.push({ type: "paragraph-join", receiver });
    edits.push(withAbsorptions(edit, paraId));
    current = edit.document;
  }
  // Kept cuts are settled where the joins leave each marker, so a cut waits
  // on the same pending changes whichever order brought the pieces together.
  if ([...resolvedContents.values()].some(holdsKeptCut)) {
    const settledCuts: DocumentOp[] = [];
    for (const { paragraph } of storyParagraphs(storyBody(current, op.story))) {
      const settled = settleKeptCuts(paragraph.content);
      if (settled === paragraph.content) continue;
      settledCuts.push({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: op.story,
        blockId: paragraph.paraId ?? "",
        expected: paragraph.content,
        content: asParagraphContent(settled),
      });
    }
    if (settledCuts.length > 0) {
      const settled = applyOps(current, settledCuts);
      if (settled.isErr()) return Result.err(settled.error);
      edits.push(settled.value);
      current = settled.value.document;
    }
  }
  // Notes whose last reference went are removed first: a comment anchored in
  // one of them is then exhausted too and retires below.
  const orphaned = noteReferences.filter(
    (reference) => !hasNoteReference(current, reference),
  );
  if (orphaned.length > 0) {
    const footnotes = new Set(
      orphaned.filter(({ kind }) => kind === "footnote").map(({ id }) => id),
    );
    const endnotes = new Set(
      orphaned.filter(({ kind }) => kind === "endnote").map(({ id }) => id),
    );
    const packageParts = { ...current.package };
    const removesAddedFootnote = orphaned.some(
      ({ kind, added }) => kind === "footnote" && added,
    );
    const removesAddedEndnote = orphaned.some(
      ({ kind, added }) => kind === "endnote" && added,
    );
    if (footnotes.size > 0 && packageParts.footnotes !== undefined) {
      packageParts.footnotes = packageParts.footnotes.filter(
        ({ id }) => !footnotes.has(id),
      );
      if (packageParts.footnotes.length === 0 && removesAddedFootnote)
        delete packageParts.footnotes;
    }
    if (endnotes.size > 0 && packageParts.endnotes !== undefined) {
      packageParts.endnotes = packageParts.endnotes.filter(
        ({ id }) => !endnotes.has(id),
      );
      if (packageParts.endnotes.length === 0 && removesAddedEndnote)
        delete packageParts.endnotes;
    }
    const removed = replaceNoteParts({
      document: current,
      parts: packageParts,
      op,
      applyOps,
    });
    if (removed.isErr()) return Result.err(removed.error);
    edits.push(removed.value);
    current = removed.value.document;
  }
  const exhausted = (document.package.document.comments ?? []).filter(
    ({ id, parentId }) =>
      parentId === undefined &&
      !isCommentExhausted({ type: "document", document, id }) &&
      isCommentExhausted({ type: "document", document: current, id }),
  );
  for (const { id } of exhausted) {
    const retired = applyOps(current, [
      { type: DOCUMENT_OP_TYPES.DELETE_COMMENT, id, scope: "thread" },
    ]);
    if (retired.isErr()) return Result.err(retired.error);
    edits.push(retired.value);
    current = retired.value.document;
  }
  const edit = combineEdits(document, edits);
  if (orphaned.length > 0) {
    const { absorbed } = edit.touched;
    edit.touched = {
      ...storyLifecycleTouched({ before: document, after: current }),
      ...(absorbed === undefined ? {} : { absorbed }),
    };
  }
  return Result.ok({ edit, receiverEffects });
};

export const resolveRevision = (
  document: Document,
  op: ResolveRevisionOp,
  applyOps: ApplyOps,
): Result<DocumentEdit, DocumentOpRefusal> =>
  planRevisionResolution({ document, op, applyOps }).map(({ edit }) => edit);
