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

import { Result } from "better-result";

import type {
  Document,
  Paragraph,
  ParagraphMarkChangeKind,
  ParagraphPropertyChange,
  Run,
  RunPropertyChange,
} from "../model/document";
import { blockListAt, endsItsContainer, storyBody, storyParagraphs } from "./blocks";
import { combineEdits, type DocumentEdit } from "./edits";
import { IDENTITY_SPACES, identityKeysIn, idKey, slotKey } from "./ids";
import {
  alikeDepth,
  asParagraphContent,
  childNodes,
  type InlineNode,
  mergeAlike,
  rebuildNode,
  sameOwnFields,
} from "./leaves";
import { paragraphLength } from "./offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { isAddedRevision, isTrackedWrapper, reviewFieldsOf, withMarkFormatting } from "./review";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type OpStory,
  type ParagraphReviewFields,
  REVISION_DECISIONS,
  type ResolveRevisionOp,
  type RevisionDecision,
} from "./types";

/** Applies operations atomically; `applyDocumentOps`, passed in to keep the modules acyclic. */
export type ApplyOps = (
  document: Document,
  ops: readonly DocumentOp[],
) => Result<DocumentEdit, DocumentOpRefusal>;

const refusal = (op: ResolveRevisionOp, reason: DocumentOpRefusal["reason"], message: string) =>
  new DocumentOpRefusal({ message, reason, opType: op.type });

/** Whether a decision removes the content of a tracked change of this kind. */
const removesContent = (node: InlineNode, decision: RevisionDecision): boolean =>
  decision === REVISION_DECISIONS.ACCEPT ? !isAddedRevision(node) : isAddedRevision(node);

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
  | { kind: "restore"; remaining: Change[]; previous: Change["previousFormatting"] | undefined };

const removePropertyChanges = <Change extends PropertyChange>(
  changes: readonly Change[],
  remove: (change: Change) => boolean,
): RemovedPropertyChanges<Change> => {
  const remaining: Change[] = [];
  let removedPrevious: Change["previousFormatting"] | undefined = undefined;
  let removedAny = false;
  let removing = false;
  for (const change of changes) {
    if (remove(change)) {
      if (!removing) {
        removedPrevious = change.previousFormatting;
        removing = true;
      }
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
    remaining.push(rebased);
    removedPrevious = undefined;
    removing = false;
  }
  if (!removedAny) return { kind: "unchanged" };
  if (!removing) return { kind: "keepLive", remaining };
  return { kind: "restore", remaining, previous: removedPrevious };
};

type Resolution = { ids: ReadonlySet<number>; decision: RevisionDecision };

const resolveRun = (run: Run, { ids, decision }: Resolution): Run => {
  const changes = run.propertyChanges ?? [];
  const selected = (change: (typeof changes)[number]): boolean => ids.has(change.info.id);
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
 * container of the same kind and fields, holding nothing, where the other
 * holds something.
 */
const emptyPieceOf = (empty: InlineNode, other: InlineNode): boolean => {
  const emptyChildren = childNodes(empty);
  const otherChildren = childNodes(other);
  return (
    empty.type !== "run" &&
    emptyChildren !== undefined &&
    otherChildren !== undefined &&
    emptyChildren.length === 0 &&
    otherChildren.length > 0 &&
    sameOwnFields(empty, other)
  );
};

/**
 * Two records meeting where a change was resolved, merged as far as they are
 * alike. A piece of a cut container that resolution emptied goes into the
 * piece it was cut from, which keeps the first one's ids.
 */
export const mergeAtSeam = (left: InlineNode, right: InlineNode): InlineNode[] => {
  if (alikeDepth(left, right) > 0) {
    return mergeAlike([left], [right]);
  }
  if (emptyPieceOf(left, right)) {
    return [rebuildNode(left, childNodes(right) ?? [])];
  }
  return emptyPieceOf(right, left) ? [left] : [left, right];
};

/** Two lists end to end, merged at the seam as {@link mergeAtSeam} merges. */
const mergedAtSeam = (left: readonly InlineNode[], right: readonly InlineNode[]): InlineNode[] => {
  const last = left.at(-1);
  const first = right.at(0);
  if (last === undefined || first === undefined) {
    return [...left, ...right];
  }
  return [...left.slice(0, -1), ...mergeAtSeam(last, first), ...right.slice(1)];
};

type ResolvedList = { nodes: InlineNode[]; changed: boolean };

/**
 * A list with its tracked changes resolved. Where a change was resolved,
 * the records left meeting are merged as far as they are alike.
 */
const resolveList = (nodes: readonly InlineNode[], resolution: Resolution): ResolvedList => {
  const out: InlineNode[] = [];
  const seams: number[] = [];
  let changed = false;
  for (const node of nodes) {
    if (isTrackedWrapper(node) && resolution.ids.has(node.info.id)) {
      changed = true;
      seams.push(out.length);
      if (removesContent(node, resolution.decision)) {
        continue;
      }
      out.push(...resolveList(node.content, resolution).nodes);
      seams.push(out.length);
      continue;
    }
    if (node.type === "run") {
      const run = resolveRun(node, resolution);
      if (run !== node) {
        changed = true;
        seams.push(out.length, out.length + 1);
      }
      out.push(run);
      continue;
    }
    const children = childNodes(node);
    const inner = children === undefined ? undefined : resolveList(children, resolution);
    if (children === undefined || inner === undefined || !inner.changed) {
      out.push(node);
      continue;
    }
    changed = true;
    // A tracked change the resolution emptied goes.
    if (isTrackedWrapper(node) && inner.nodes.length === 0 && children.length > 0) {
      seams.push(out.length);
      continue;
    }
    out.push(rebuildNode(node, inner.nodes));
  }
  if (!changed) {
    return { nodes: out, changed };
  }
  for (const seam of [...new Set(seams)].toSorted((left, right) => right - left)) {
    const left = out[seam - 1];
    const right = out[seam];
    if (left === undefined || right === undefined) continue;
    out.splice(seam - 1, 2, ...mergeAtSeam(left, right));
  }
  return { nodes: out, changed };
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
  const selected = (change: (typeof changes)[number]): boolean => ids.has(change.info.id);
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
        const restored = withMarkFormatting(removed.previous, paragraph.formatting);
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
        for (const change of node.propertyChanges ?? []) out.add(change.info.id);
      }
      visit(childNodes(node) ?? []);
    }
  };
  visit(paragraph.content);
  return out;
};

type JoinPlan = { op: ResolveRevisionOp; story: OpStory; paraId: string; added: boolean };

/** The operations that remove a paragraph's resolved mark, against the document as it stands. */
const joinOps = (
  document: Document,
  { op, story, paraId, added }: JoinPlan,
): Result<DocumentOp[], DocumentOpRefusal> => {
  const body = storyBody(document, story);
  const location = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(paraId),
  );
  if (location === undefined) {
    return Result.err(
      refusal(op, DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND, `No paragraph is ${paraId}.`),
    );
  }
  const { paragraph } = location;
  if (paragraph.sectionProperties !== undefined) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        `The mark of ${paraId} ends a section: joining it is a section operation.`,
      ),
    );
  }
  const blocks = blockListAt(body.content, location.list);
  const next = blocks[location.index + 1];
  const empty = paragraphLength(paragraph) === 0;
  if (next?.type === "paragraph") {
    // The mark that goes takes the paragraph's properties with it: the next
    // paragraph is left, whole, with the first's content before its own.
    const survivor: Paragraph = {
      ...next,
      content: asParagraphContent(mergedAtSeam(paragraph.content, next.content)),
    };
    return Result.ok([
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story,
        expected: [paragraph, next],
        blocks: [survivor],
      },
    ]);
  }
  if (next !== undefined && (next.type === "bookmarkStart" || next.type === "bookmarkEnd")) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        `A block-level bookmark stands between ${paraId} and the next paragraph.`,
      ),
    );
  }
  // No paragraph to join: a table follows, or the paragraph ends its container.
  const canGo = added || !endsItsContainer(body, location);
  if (empty && canGo && blocks.length > 1) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        `Resolving the mark of ${paraId} removes the paragraph, which is a block operation.`,
      ),
    );
  }
  const review = reviewFieldsOf(paragraph);
  delete review.pPrMark;
  return Result.ok([
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

/**
 * Resolve tracked changes by revision id, applied as the primitive
 * operations it expands to. An id no record in the story carries is
 * skipped; one carried by a record resolution does not reach yet (a table,
 * row or section change, a record inside a field) is refused.
 */
export const resolveRevision = (
  document: Document,
  op: ResolveRevisionOp,
  applyOps: ApplyOps,
): Result<DocumentEdit, DocumentOpRefusal> => {
  if (!Object.values(REVISION_DECISIONS).includes(op.decision)) {
    return Result.err(
      refusal(op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, "A decision accepts or rejects."),
    );
  }
  if (!op.revisionIds.every((id) => Number.isInteger(id))) {
    return Result.err(
      refusal(op, DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID, "A revision id is an integer."),
    );
  }
  const present = storyIdentityKeys(document, op.story);
  const ids = new Set(
    op.revisionIds.filter((id) => present.has(slotKey({ space: IDENTITY_SPACES.REVISION, id }))),
  );
  if (ids.size === 0) {
    return Result.ok({
      document,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  }
  const resolution: Resolution = { ids, decision: op.decision };
  const paragraphs = storyParagraphs(storyBody(document, op.story));
  const reachable = new Set(paragraphs.flatMap(({ paragraph }) => [...reachableIds(paragraph)]));
  const unreachable = [...ids].find((id) => !reachable.has(id));
  if (unreachable !== undefined) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        `Revision ${unreachable} is not an inline change, property change or paragraph mark.`,
      ),
    );
  }

  const inline: DocumentOp[] = [];
  const joins: JoinPlan[] = [];
  for (const { paragraph } of paragraphs) {
    const paraId = paragraph.paraId ?? "";
    const resolved = resolveList(paragraph.content, resolution);
    if (resolved.changed) {
      inline.push({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: op.story,
        blockId: paraId,
        expected: paragraph.content,
        content: asParagraphContent(resolved.nodes),
      });
    }
    const mark = paragraph.pPrMark;
    const markResolved = mark !== undefined && ids.has(mark.info.id);
    const added = mark !== undefined && markWasAdded(mark.kind);
    const keepsBreak = markResolved && added === (op.decision === REVISION_DECISIONS.ACCEPT);
    if (markResolved && !keepsBreak) {
      joins.push({ op, story: op.story, paraId, added });
    }
    const review = resolveParagraphReview(paragraph, resolution, keepsBreak);
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
  const edits: DocumentEdit[] = [staged.value];
  let current = staged.value.document;
  for (const join of joins.toReversed()) {
    const planned = joinOps(current, join);
    if (planned.isErr()) {
      return Result.err(planned.error);
    }
    const joined = applyOps(current, planned.value);
    if (joined.isErr()) {
      return Result.err(joined.error);
    }
    edits.push(joined.value);
    current = joined.value.document;
  }
  return Result.ok(combineEdits(document, edits));
};
