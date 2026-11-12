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
  paragraphTableRunInTarget,
  type TableRunInTarget,
} from "./blockDeletion";
import {
  blockListAt,
  endsItsContainer,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { isCommentExhausted } from "./comments";
import {
  acceptCopyOriginAncestor,
  rejectCopyOriginAncestorsInNodes,
} from "./copyOriginAncestors";
import { combineEdits, type DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { canonicalDeferredGroups } from "./identity";
import { IDENTITY_SPACES, identityKeysIn, idKey, slotKey } from "./ids";
import { joinParagraphSeam } from "./inline";
import {
  alikeDepth,
  asParagraphContent,
  childNodes,
  type InlineNode,
  leafSpans,
  mergeAlike,
  mergeLists,
  rebuildNode,
  sameOwnFields,
} from "./leaves";
import {
  paragraphHoldsContent,
  paragraphHoldsContentOnceAccepted,
} from "./offsets";
import {
  isBlockBoundary,
  isBoundary,
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
import { reachableRowIds, resolveTableRows } from "./resolveTableRows";
import {
  isAddedRevision,
  isTrackedWrapper,
  reviewFieldsOf,
  type TrackedWrapper,
  withMarkFormatting,
} from "./review";
import { runsMergeable } from "./runMerge";
import { identitySlots, withInlineIdentity } from "./slots";
import { documentStories } from "./stories";
import { storyLifecycleEdit } from "./storyLifecycle";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
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
    for (const { node, ancestors } of leafSpans(paragraph.content)) {
      if (node.type !== "footnoteRef" && node.type !== "endnoteRef") continue;
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
      leafSpans(paragraph.content).some(
        ({ node }) =>
          (node.type === "footnoteRef" || node.type === "endnoteRef") &&
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
  ids: Set<number>;
  decision: RevisionDecision;
  /** The containers this resolution has emptied so far. */
  emptied: WeakSet<InlineNode>;
  /** Actual payload edges accepted after source identities were restored. */
  acceptedClosed: Record<"first" | "last", WeakSet<InlineNode>>;
  /** Unconsumed source-cut edges on the actual rebuilt container. */
  cutEdges: WeakMap<InlineNode, CutEdges>;
  /** Comment ranges with a start and an end in the paragraphs resolved. */
  pairedComments: ReadonlySet<number>;
  /** Boundaries this resolution kept from removed content. */
  relocated: RelocatedBoundaries;
  /**
   * Identity handoffs whose retained piece was not beside the removed prefix:
   * a replacement's new content, or a paragraph break it inserted, stands
   * between them. They reach that piece later in the same story.
   */
  outstanding: RetainedIdentity[];
};

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
            depth: group.depth,
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
    if (!Array.isArray(groups) || groups.length === 0 || !isAddedRevision(node))
      return false;
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
        !deferred.blockers.includes(node.info.id)
      )
        return false;
      const key = `${deferred.depth}:${deferred.blockers.join(",")}`;
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

/** A range boundary standing on its own in a list, held by no record. */
const isStandingBoundary = (node: InlineNode | undefined): boolean =>
  node !== undefined && isBoundary(node);

/** What has become of a deferred seam's blocker. */
type BlockerFate = "pending" | "accepted" | "removed";

/**
 * The one rule for a deferred seam whose blockers leave the story. An
 * accepted blocker retires the seam; a removed one, rejected or emptied by
 * resolving the changes it held, no longer stands between the pieces, and
 * once none stands the seam joins when its holder is removed.
 */
const settleDeferredRemovals = (
  nodes: readonly InlineNode[],
  fateOf: (blocker: number) => BlockerFate,
): readonly InlineNode[] => {
  let changed = false;
  const out: InlineNode[] = [];
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
          if (fates.includes("accepted")) continue;
          const remaining = group.blockers.filter(
            (_, index) => fates[index] === "pending",
          );
          if (remaining.length === 0)
            joins.remove = Math.max(joins.remove, group.depth);
          else
            remainingGroups.push({ depth: group.depth, blockers: remaining });
        }
        delete joins.deferredRemove;
        if (remainingGroups.length > 0)
          joins.deferredRemove = canonicalDeferredGroups(remainingGroups);
        next = Object.assign({}, node, { resolutionJoins: joins });
      }
    }
    const children = childNodes(next);
    const settled =
      children === undefined
        ? children
        : settleDeferredRemovals(children, fateOf);
    if (children !== undefined && settled !== children)
      next = rebuildNode(next, settled ?? []);
    changed ||= next !== node;
    out.push(next);
  }
  return changed ? out : nodes;
};

/** Update an explicit pending seam before removing any of its blocking revisions. */
const prepareDeferredRemovals = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): readonly InlineNode[] =>
  settleDeferredRemovals(nodes, (id) => {
    if (!resolution.ids.has(id)) return "pending";
    return resolution.decision === REVISION_DECISIONS.ACCEPT
      ? "accepted"
      : "removed";
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
 * The ids of the pending insertions in `nodes`, at any depth, that can hold a
 * deferred seam: the only records a deferred seam may name as its blockers.
 */
const seamHolderIds = (nodes: readonly InlineNode[]): Set<number> =>
  new Set(
    trackedWithin(nodes).flatMap((node) =>
      isAddedRevision(node) && node.resolutionJoins !== undefined
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
      return blockers.length === 0 ? [] : [{ depth: group.depth, blockers }];
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

/** `node` and its last (`end`) or first (`start`) child, theirs, and so on. */
const edgeChain = (node: InlineNode, edge: "start" | "end"): InlineNode[] => {
  const chain: InlineNode[] = [];
  let current: InlineNode | undefined = node;
  while (current !== undefined) {
    chain.push(current);
    const children = childNodes(current);
    current = edge === "start" ? children?.at(0) : children?.at(-1);
  }
  return chain;
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
 * deletion made over the cut does. So the reach is the same whichever
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
    const near = before[nearIndex];
    const far = after[farIndex];
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
      pass("before", near);
      nearIndex += 1;
      continue;
    }
    if (!piecesOfOne && isTrackedWrapper(far)) {
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

const resolveList = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): ResolvedList => {
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
  const appendUnowned = (items: readonly InlineNode[]): void => {
    out.push(...items.map((node) => ({ node, owner: undefined })));
  };
  const recordExactSeam = (index: number, depth: number): void => {
    if (depth > 0)
      exactSeams.set(index, Math.max(exactSeams.get(index) ?? 0, depth));
  };
  /** Seams a removed revision recorded between the source pieces around it. */
  const removalSeams = new Set<number>();
  const recordSourceSeam = (index: number, depth: number): void => {
    recordExactSeam(index, depth);
    if (depth > 0)
      sourceEdges.set(index, Math.max(sourceEdges.get(index) ?? 0, depth));
  };
  /** Seams beside an emptied container, where only its fold happens. */
  const folds: number[] = [];
  /** Boundaries kept from removed content just before a seam, and its depth. */
  const floating = new Map<
    number,
    {
      count: number;
      depth: number;
      before: SourceOwner | undefined;
      after: SourceOwner | undefined;
      /** The tracked changes the removed content held. */
      spanning: readonly TrackedWrapper[];
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
        const removedDepth = node.resolutionJoins?.remove ?? 0;
        if (
          resolution.decision === REVISION_DECISIONS.REJECT &&
          isAddedRevision(node) &&
          removedDepth > 0
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
                before === undefined
                  ? undefined
                  : { owner: before, depth: removedDepth },
              after:
                after === undefined
                  ? undefined
                  : { owner: after, depth: removedDepth },
            },
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
        const kept = removedContentBoundaries({
          nodes: removalContent,
          rejecting: resolution.decision === REVISION_DECISIONS.REJECT,
          pairedComments: resolution.pairedComments,
          relocated: resolution.relocated,
          revisionDate: node.info.date ?? "",
        });
        appendUnowned(kept);
        // Kept boundaries float at the seam until the pieces either side are
        // resolved; the join then places them inside the piece it joins.
        if (
          kept.length > 0 &&
          resolution.decision === REVISION_DECISIONS.REJECT &&
          (node.resolutionJoins?.remove ?? 0) > 0
        )
          floating.set(out.length, {
            count: kept.length,
            depth: node.resolutionJoins?.remove ?? 0,
            before: index > 0 ? sourceOwners.at(index - 1) : undefined,
            after: sourceOwners.at(index + 1),
            spanning: trackedWithin(node.content),
          });
        if (node.resolutionJoins === undefined) seams.push(out.length);
        else if ((node.resolutionJoins.plainSides?.length ?? 0) > 0)
          plainSeams.add(out.length);
        else {
          // Pieces of the cut this resolution also removes take that level
          // of it with them: the cut joins only what remains of what it split.
          const removalCut = cutLevelsOnceResolved(node, index);
          recordSourceSeam(out.length, removalCut);
          if (removalCut > 0) removalSeams.add(out.length);
        }
        continue;
      }
      if (node.resolutionJoins === undefined) seams.push(out.length);
      else if (node.resolutionJoins.plainSides?.includes("before") === true)
        plainSeams.add(out.length);
      else recordSourceSeam(out.length, node.resolutionJoins.before);
      const resolved = resolveList(node.content, resolution);
      const content =
        isAddedRevision(node) &&
        resolution.decision === REVISION_DECISIONS.ACCEPT
          ? acceptCopyOriginAncestor(resolved.nodes)
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
      else recordSourceSeam(out.length, node.resolutionJoins.after);
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
      children === undefined ? undefined : resolveList(children, resolution);
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
      appendSource(recut(node), index);
      continue;
    }
    changed = true;
    // A tracked change the resolution emptied goes. So does any record an
    // acceptance empties: it removes content as a direct deletion does, which
    // leaves no record with nothing. Kept, the record could not tell a later
    // resolution it was emptied rather than empty before, and would stay where
    // accepting the same changes at once folds it into an alike neighbour.
    if (
      (isTrackedWrapper(node) ||
        resolution.decision === REVISION_DECISIONS.ACCEPT) &&
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
    let rebuilt = recut(rebuildNode(node, inner.nodes));
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
  const edges: CutEdges = {};
  const firstDepth = sourceEdges.get(0);
  const lastDepth = sourceEdges.get(out.length);
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
  // Boundaries an earlier rejection kept at a seam still stand between the
  // pieces it joins: they float there as the ones this resolution kept do,
  // so rejecting first and joining later places them as the other order.
  for (const [recorded, depth] of [...exactSeams]) {
    if (floating.has(recorded)) continue;
    const standingAt = (from: number, step: 1 | -1): number => {
      let count = 0;
      while (isStandingBoundary(out[from + step * count]?.node)) count += 1;
      return count;
    };
    const after = standingAt(recorded, 1);
    const before = after > 0 ? 0 : standingAt(recorded - 1, -1);
    const seam = recorded + after;
    const count = after + before;
    const left = out[seam - count - 1];
    const right = out[seam];
    if (count === 0 || left === undefined || right === undefined) continue;
    exactSeams.delete(recorded);
    exactSeams.set(seam, Math.max(exactSeams.get(seam) ?? 0, depth));
    floating.set(seam, {
      count,
      depth,
      before: left.owner,
      after: right.owner,
      spanning: [],
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
    if (kept !== undefined) {
      const boundaries = out
        .slice(seam - kept.count, seam)
        .map(({ node }) => node);
      const beforeOccurrence = out[seam - kept.count - 1];
      const afterOccurrence = out[seam];
      const before = beforeOccurrence?.node;
      const after = afterOccurrence?.node;
      // The depth counts the containers the removed content cut, read on
      // the records either side as the list held them: a container this
      // resolution resolves away is a level no more, and a tracked change on
      // one side only was never one. A side with no source record is read
      // against itself.
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
        const far = (side === "before" ? afterOccurrence : beforeOccurrence)
          ?.source;
        const nearChain = edgeChain(near, side === "before" ? "end" : "start");
        const farChain =
          far === undefined
            ? nearChain
            : edgeChain(far, side === "before" ? "start" : "end");
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
      const intoAfter =
        after === undefined || afterDepth === undefined
          ? undefined
          : withBoundariesAtEdge(after, boundaries, afterDepth, "start");
      const intoBefore =
        intoAfter !== undefined ||
        before === undefined ||
        beforeDepth === undefined
          ? undefined
          : withBoundariesAtEdge(before, boundaries, beforeDepth, "end");
      if (intoAfter !== undefined && after !== undefined) {
        carryResolutionState(resolution, after, intoAfter);
        out.splice(
          seam - kept.count,
          kept.count + 1,
          occurrenceOf(intoAfter, afterOccurrence),
        );
        seam -= kept.count;
      } else if (intoBefore !== undefined && before !== undefined) {
        carryResolutionState(resolution, before, intoBefore);
        out.splice(
          seam - kept.count - 1,
          kept.count + 1,
          occurrenceOf(intoBefore, beforeOccurrence),
        );
        seam -= kept.count;
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
    // Boundaries floating at the seam are placed with its own join; a
    // deferred join would place them apart from the pieces they belong in.
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
      kept === undefined &&
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
      let leftEnd = seam - 1;
      while (leftEnd >= 0 && blocking(leftEnd)) leftEnd -= 1;
      let rightStart = seam;
      while (rightStart < out.length && blocking(rightStart)) rightStart += 1;
      if (
        (leftEnd !== seam - 1 || rightStart !== seam) &&
        leftEnd >= 0 &&
        rightStart < out.length
      ) {
        const blockerIndices = [
          ...Array.from(
            { length: seam - 1 - leftEnd },
            (_, offset) => leftEnd + 1 + offset,
          ),
          ...Array.from(
            { length: rightStart - seam },
            (_, offset) => seam + offset,
          ),
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
            panic("A deferring seam blocker is a pending insertion.");
          const groups = [...(node.resolutionJoins.deferredRemove ?? [])];
          groups.push({ depth: exactDepth, blockers });
          out[index] = occurrenceOf(
            Object.assign({}, node, {
              resolutionJoins: Object.assign({}, node.resolutionJoins, {
                deferredRemove: canonicalDeferredGroups(groups),
              }),
            }),
            occurrence,
          );
        }
        continue;
      }
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
        const deferred = deferRunReviewSeam({ left, right, depth: exactDepth });
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
  const first = out.at(0)?.node;
  const last = out.at(-1)?.node;
  const nestedFirst =
    first === undefined ? undefined : resolution.cutEdges.get(first)?.first;
  const nestedLast =
    last === undefined ? undefined : resolution.cutEdges.get(last)?.last;
  if (nestedFirst !== undefined)
    edges.first = Math.max(edges.first ?? 0, nestedFirst);
  if (nestedLast !== undefined)
    edges.last = Math.max(edges.last ?? 0, nestedLast);
  return {
    nodes: out.map(({ node }) => node),
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
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node)) out.add(node.info.id);
      if (node.type === "run") {
        for (const change of node.propertyChanges ?? [])
          out.add(change.info.id);
      }
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
  added: boolean;
  depth: number | undefined;
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
};

const paragraphMarkPlan = (
  ops: DocumentOp[],
  receiverEffects: ResolutionReceiverEffect[] = [],
) => Result.ok({ ops, receiverEffects });

const joinOps = (
  document: Document,
  { op, story, paraId, added, depth }: JoinPlan,
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
  const next = blocks[location.index + 1];
  const previous = blocks[location.index - 1];
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
      const left = seam.left.at(-1);
      const right = seam.right.at(0);
      if (globalThis.__folioJoinDebug === true)
        console.error(
          "DEBUGJOIN",
          JSON.stringify({ depth, added, paragraph: paragraph.content, next: next.content, seamLeft: seam.left, seamRight: seam.right, emptied: [...emptied] }),
        );
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
    if (globalThis.__folioJoinDebug === true)
      console.error("DEBUGMERGED", JSON.stringify({ depth, merged }));
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
    const receiverEffects: ResolutionReceiverEffect[] = [
      { type: "paragraph-join", receiver: next },
    ];
    if (transfersToPrevious)
      receiverEffects.push({ type: "section-transfer", receiver: previous });
    if (receivesSection)
      receiverEffects.push({ type: "section-transfer", receiver: next });
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
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        `A range boundary stands between ${paraId} and the next block.`,
      ),
    );
  }
  // No paragraph to join: a table or another block follows, or the
  // paragraph ends its container.
  const followsBlock = isOpaqueFollower(next);
  const containerFinal = endsItsContainer(body, location);
  // A cell-ending mark describes the terminator, not a paragraph to retire.
  const cellFinal =
    containerFinal && location.list.some((step) => step.kind === "tableCell");
  const tableTarget =
    next?.type === "table"
      ? paragraphTableRunInTarget(paragraph, next)
      : undefined;
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
        : // Removing the final paragraph leaves the block before it last.
          containerFinal &&
          !cellFinal &&
          blocks.at(location.index - 1)?.type === "paragraph",
    followingTableCanReceiveRunIn: tableTarget !== undefined,
    holdsContentOnceAccepted: paragraphHoldsContentOnceAccepted(chained),
    followingBlockCanReceiveRunIn: blockTarget !== undefined,
    needsFollowingCommentHost:
      next?.type === "table" &&
      previous?.type !== "paragraph" &&
      remains.blocks.length === 0 &&
      remains.comments.length > 0,
    tableRemovalIsTrackable:
      next?.type !== "table" ||
      markBeforeTableResolves({ blocks, paragraph, table: next }),
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
  if (
    (tableOutcome === "remove" || tableOutcome === "run-in") &&
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
        place: "blocks",
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
      place: "previous",
    });
    if (anchor !== undefined && removal.type === "removed") {
      const replacementBlocks = removal.blocks;
      const preceding = replacementBlocks[location.index - 1];
      const receiverEffects: ResolutionReceiverEffect[] = [...removal.effects];
      if (
        outcome.sectionDestination === "preceding" &&
        preceding?.type === "paragraph"
      ) {
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
  return paragraphMarkPlan([
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
      story,
      blockId: paraId,
      expected: reviewFieldsOf(paragraph),
      review,
    },
  ]);
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
 * Resolve tracked changes by revision id, applied as the primitive
 * operations it expands to. An id no record in the story carries is
 * skipped; one carried by a record resolution does not reach yet (a table,
 * cell or section change, a record inside a field) is refused.
 */
export const planRevisionResolution = ({
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
  const resolution: Resolution = {
    ids,
    decision: op.decision,
    emptied: new WeakSet(),
    acceptedClosed: { first: new WeakSet(), last: new WeakSet() },
    cutEdges: new WeakMap(),
    pairedComments: new Set(),
    relocated: noRelocatedBoundaries(),
    outstanding: [],
  };
  const noteReferences = removedNoteReferences(document, op, ids);
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
  const preparedContents = paragraphs.map(({ paragraph }) => ({
    paragraph,
    content: prepareDeferredRemovals(paragraph.content, resolution),
  }));
  resolution.pairedComments = pairedCommentRanges(
    preparedContents.flatMap(({ content }) => content),
  );
  const removedCommentIds = removedCommentRangeIds(
    preparedContents.flatMap(({ content }) => content),
    resolution,
  );
  const resolvedChanged = new Set<string>();
  for (const { paragraph, content: source } of preparedContents) {
    const paraId = paragraph.paraId ?? "";
    // A replacement's paragraph break can stand between a removed prefix and
    // the piece that takes its identities: they arrive here.
    const handed = restoreOutstandingIdentities(source, resolution);
    const prepared =
      handed === source ? source : prepareDeferredRemovals(handed, resolution);
    const resolved = resolveList(prepared, resolution);
    resolvedContents.set(
      idKey(paraId),
      withoutCommentAnchors(resolved.nodes, removedCommentIds),
    );
    if (resolved.edges !== undefined)
      resolvedCutEdges.set(idKey(paraId), resolved.edges);
    if (resolved.changed || prepared !== paragraph.content)
      resolvedChanged.add(idKey(paraId));
  }
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
    const settled = settleDeferredRemovals(content, (id) =>
      standing.has(id) ? "pending" : "removed",
    );
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
    const next = blockListAt(body.content, location.list).at(
      location.index + 1,
    );
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
    let leadingCut =
      next?.type === "paragraph"
        ? (resolvedCutEdges.get(idKey(next.paraId ?? ""))?.first ?? 0)
        : 0;
    const following = blockListAt(body.content, location.list);
    for (let index = location.index + 1; index < following.length; index += 1) {
      const candidate = following.at(index);
      if (candidate?.type !== "paragraph") break;
      leadingCut = Math.max(
        leadingCut,
        resolvedCutEdges.get(idKey(candidate.paraId ?? ""))?.first ?? 0,
      );
      if (
        (resolvedContents.get(idKey(candidate.paraId ?? ""))?.length ?? 0) !==
          0 ||
        candidate.pPrMark === undefined ||
        !markWasAdded(candidate.pPrMark.kind)
      )
        break;
      // An empty inserted paragraph can carry the source cut for a later
      // join. Rejecting its break must carry that cut across the entire chain.
      if (
        op.decision === REVISION_DECISIONS.REJECT &&
        ids.has(candidate.pPrMark.info.id)
      )
        leadingCut = Math.max(
          leadingCut,
          candidate.pPrMark.resolutionJoin ?? 0,
        );
    }
    let depth = Math.max(
      mark.resolutionJoin,
      resolvedCutEdges.get(idKey(paragraph.paraId ?? ""))?.last ?? 0,
      leadingCut,
    );
    if (trailing !== undefined) depth = Math.min(depth, trailing);
    if (leading !== undefined) depth = Math.min(depth, leading);
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
      joins.push({ op, story: op.story, paraId, added, depth });
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
    op,
    applyOps,
  });
  if (rows.isErr()) return Result.err(rows.error);
  const removedParagraphs = new Set(rows.value.touched.removed.map(idKey));
  const edits: DocumentEdit[] = [staged.value, rows.value];
  let current = rows.value.document;
  const ranInto = new Set<string>();
  const joinsBefore = new Set(joins.map(({ paraId }) => idKey(paraId)));
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
    edits.push(joined.value);
    current = joined.value.document;
  }
  // Notes whose last reference went are removed first: a comment anchored in
  // one of them is then exhausted too and retires below.
  let notesRemoved = false;
  if (noteReferences.length > 0) {
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
      current = { ...current, package: packageParts };
      notesRemoved = true;
    }
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
  // A removed note changes story parts; one lifecycle edit captures them and
  // the retired comments together.
  if (notesRemoved)
    return storyLifecycleEdit(document, current, op, true).map((edit) => ({
      edit,
      receiverEffects,
    }));
  return Result.ok({ edit: combineEdits(document, edits), receiverEffects });
};

export const resolveRevision = (
  document: Document,
  op: ResolveRevisionOp,
  applyOps: ApplyOps,
): Result<DocumentEdit, DocumentOpRefusal> =>
  planRevisionResolution({ document, op, applyOps }).map(({ edit }) => edit);
