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

import { INSERTION_SEAM_POLICIES } from "../model/content";
import { runsMergeable } from "./runMerge";
import { panic, Result } from "better-result";

import { MAX_REVISION_ID } from "../model/document";
import type {
  Document,
  Deletion,
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
  mergeLists,
  rebuildNode,
  sameOwnFields,
} from "./leaves";
import { paragraphLength } from "./offsets";
import { identitySlots, withInlineIdentity } from "./slots";
import { canonicalDeferredGroups } from "./identity";
import { joinParagraphSeam } from "./inline";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { reachableRowIds, resolveTableRows } from "./resolveTableRows";
import { isAddedRevision, isTrackedWrapper, reviewFieldsOf, withMarkFormatting } from "./review";
import {
  DOCUMENT_OP_TYPES,
  SECTION_BOUNDARY_POLICIES,
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
            ...removedRunChanges.flatMap((removed) => removed.boundaryJoins ?? []),
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
};

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
 * container of the same kind and fields that the resolution emptied, where
 * the other holds something or was emptied too (two emptied pieces are one
 * piece, whichever emptied first). A container that held nothing before (an
 * empty content control or hyperlink) is markup of its own and is never merged.
 */
const emptyPieceOf = (empty: InlineNode, other: InlineNode, emptied: Emptied): boolean => {
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
  if (emptyPieceOf(left, right, emptied)) return rebuildNode(left, childNodes(right) ?? []);
  return undefined;
};

/**
 * Two records meeting beside a container resolution emptied, with no change
 * resolved between them: the emptied piece folds into its neighbour, down the
 * chain of alike containers at their facing edges when it sits inside one.
 * Nothing else merges there.
 */
const foldAtSeam = (left: InlineNode, right: InlineNode, emptied: Emptied): InlineNode[] => {
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
  return [rebuildNode(left, [...leftChildren.slice(0, -1), ...inner, ...rightChildren.slice(1)])];
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
  return folded === undefined ? [left, right] : [folded];
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
  return [...left.slice(0, -1), ...mergeAtSeam(last, first, emptied), ...right.slice(1)];
};

type RetainedIdentity = NonNullable<
  NonNullable<Deletion["resolutionJoins"]>["retainedAfter"]
>[number];

const hasSlots = (node: InlineNode, expected: RetainedIdentity["target"]): boolean => {
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
}: RestoreRetainedIdentitiesOptions): InlineNode[] => {
  const references = new Map<string, RetainedIdentity["source"][number]>();
  let out = [...nodes];
  for (const transfer of transfers) {
    const source = transfer.source.map((slot) => references.get(slotKey(slot)) ?? slot);
    let matched = false;
    const restore = (node: InlineNode): InlineNode => {
      const matches = hasSlots(node, transfer.target);
      matched ||= matches;
      let next = matches
        ? withInlineIdentity(
            node,
            identitySlots(node).map((slot) => {
              const index = transfer.target.findIndex(
                (target) => slot.space === target.space && slot.id === target.id,
              );
              return index < 0
                ? slot.id
                : (source.at(index)?.id ??
                    panic("A validated identity transfer has its source slot."));
            }),
          )
        : node;
      const children = childNodes(next);
      if (children !== undefined) next = rebuildNode(next, children.map(restore));
      return next;
    };
    out = out.map(restore);
    if (!matched) continue;
    for (const [index, target] of transfer.target.entries()) {
      const restored = source.at(index);
      if (restored === undefined || restored.space !== target.space)
        panic("Validated identity transfers retain slot order and space.");
      references.set(slotKey(target), restored);
      // Selection follows the actual retained record through a source-ID
      // transfer, rather than resolving an unrelated later use of its old ID.
      if (target.space === IDENTITY_SPACES.REVISION && resolution.ids.has(target.id))
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
          source: entry.source.map((slot) => references.get(slotKey(slot)) ?? slot),
          target: entry.target.map((slot) => references.get(slotKey(slot)) ?? slot),
        }));
      }
      if (joins.deferredRemove !== undefined) {
        joins.deferredRemove = canonicalDeferredGroups(
          joins.deferredRemove.map((group) => ({
            depth: group.depth,
            blockers: [
              ...new Set(
                group.blockers.map((id) => {
                  const restored = references.get(slotKey({ space: IDENTITY_SPACES.REVISION, id }));
                  return restored?.space === IDENTITY_SPACES.REVISION ? restored.id : id;
                }),
              ),
            ],
          })),
        );
      }
      next = Object.assign({}, node, { resolutionJoins: joins });
    }
    const children = childNodes(next);
    return children === undefined ? next : rebuildNode(next, children.map(rebase));
  };
  return out.map(rebase);
};

/** Journal-supplied provenance must name actual source slots at its recorded depth. */
const validRetainedIdentities = (node: InlineNode): boolean => {
  if (!isTrackedWrapper(node) || node.resolutionJoins === undefined) return true;
  const joins = node.resolutionJoins;
  if (
    typeof joins !== "object" ||
    joins === null ||
    ![joins.before, joins.after, joins.remove].every(
      (depth) => Number.isInteger(depth) && depth >= 0 && depth <= MAX_REVISION_ID,
    )
  )
    return false;
  if (joins.deferredRemove !== undefined) {
    const groups = joins.deferredRemove;
    if (!Array.isArray(groups) || groups.length === 0 || !isAddedRevision(node)) return false;
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
            typeof id === "number" && Number.isInteger(id) && id >= 0 && id <= MAX_REVISION_ID,
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
  const validSlot = (slot: unknown): slot is RetainedIdentity["source"][number] => {
    if (typeof slot !== "object" || slot === null || !("space" in slot) || !("id" in slot))
      return false;
    // Revision/control ids allow zero; note-record reserved ids do not apply here.
    const retainedId = slot.id;
    return (
      (slot.space === IDENTITY_SPACES.REVISION || slot.space === IDENTITY_SPACES.CONTROL) &&
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

/** Update an explicit pending seam before removing any of its blocking revisions. */
const prepareDeferredRemovals = (
  nodes: readonly InlineNode[],
  resolution: Resolution,
): readonly InlineNode[] => {
  let changed = false;
  const out: InlineNode[] = [];
  for (const node of nodes) {
    let next = node;
    if (isTrackedWrapper(node) && node.resolutionJoins?.deferredRemove !== undefined) {
      const groups = node.resolutionJoins.deferredRemove;
      if (groups.some((group) => group.blockers.some((id) => resolution.ids.has(id)))) {
        const joins = Object.assign({}, node.resolutionJoins);
        const remainingGroups: NonNullable<typeof joins.deferredRemove>[number][] = [];
        for (const group of groups) {
          if (!group.blockers.some((id) => resolution.ids.has(id))) {
            remainingGroups.push(group);
            continue;
          }
          if (resolution.decision === REVISION_DECISIONS.ACCEPT) continue;
          const remaining = group.blockers.filter((id) => !resolution.ids.has(id));
          if (remaining.length === 0) joins.remove = Math.max(joins.remove, group.depth);
          else remainingGroups.push({ depth: group.depth, blockers: remaining });
        }
        delete joins.deferredRemove;
        if (remainingGroups.length > 0)
          joins.deferredRemove = canonicalDeferredGroups(remainingGroups);
        next = Object.assign({}, node, { resolutionJoins: joins });
      }
    }
    const children = childNodes(next);
    const prepared =
      children === undefined ? children : prepareDeferredRemovals(children, resolution);
    if (children !== undefined && prepared !== children) next = rebuildNode(next, prepared ?? []);
    changed ||= next !== node;
    out.push(next);
  }
  return changed ? out : nodes;
};

type CutEdges = { first?: number; last?: number };
type ResolvedList = { nodes: InlineNode[]; changed: boolean; edges?: CutEdges };

/**
 * A list with its tracked changes resolved. Where a change was resolved,
 * the records left meeting are merged as far as they are alike.
 */
const resolveList = (nodes: readonly InlineNode[], resolution: Resolution): ResolvedList => {
  const out: InlineNode[] = [];
  const seams: number[] = [];
  const plainSeams = new Set<number>();
  const exactSeams = new Map<number, number>();
  const sourceEdges = new Map<number, number>();
  const pending = [...nodes];
  const recordExactSeam = (index: number, depth: number): void => {
    if (depth > 0) exactSeams.set(index, Math.max(exactSeams.get(index) ?? 0, depth));
  };
  const recordSourceSeam = (index: number, depth: number): void => {
    recordExactSeam(index, depth);
    if (depth > 0) sourceEdges.set(index, Math.max(sourceEdges.get(index) ?? 0, depth));
  };
  /** Seams beside an emptied container, where only its fold happens. */
  const folds: number[] = [];
  let changed = false;
  for (const [index, node] of pending.entries()) {
    if (isTrackedWrapper(node) && resolution.ids.has(node.info.id)) {
      changed = true;
      if (removesContent(node, resolution.decision)) {
        const transfers = node.resolutionJoins?.retainedAfter ?? [];
        if (transfers.length > 0) {
          // Restore before resolving later wrappers, so their own source
          // references follow the fragment whose identity returned.
          const prefixLength = out.length;
          const restored = restoreRetainedIdentities({
            nodes: out.concat(pending.slice(index + 1)),
            transfers,
            resolution,
          });
          // Restoring source IDs can make a previously unselected blocker
          // selected; recompute deferred facts against those actual identities.
          const prepared = prepareDeferredRemovals(restored, resolution);
          out.splice(0, out.length, ...prepared.slice(0, prefixLength));
          pending.splice(index + 1, pending.length - index - 1, ...prepared.slice(prefixLength));
        }
        if (node.resolutionJoins === undefined) seams.push(out.length);
        else recordSourceSeam(out.length, node.resolutionJoins.remove);
        continue;
      }
      if (node.resolutionJoins === undefined) seams.push(out.length);
      else recordSourceSeam(out.length, node.resolutionJoins.before);
      const resolved = resolveList(node.content, resolution);
      const content = resolved.nodes;
      if (resolved.edges?.first !== undefined) recordSourceSeam(out.length, resolved.edges.first);
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
        node.resolutionJoins?.acceptance === INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS
      ) {
        plainSeams.add(out.length);
        plainSeams.add(out.length + content.length);
      }
      out.push(...content);
      if (resolved.edges?.last !== undefined) recordSourceSeam(out.length, resolved.edges.last);
      if (node.resolutionJoins === undefined) seams.push(out.length);
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
              change.boundaryJoins === undefined || change.boundaryJoins.includes("before"),
          )
        )
          seams.push(out.length);
        if (
          selected.some(
            (change) =>
              change.boundaryJoins === undefined || change.boundaryJoins.includes("after"),
          )
        )
          seams.push(out.length + 1);
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
    // A tracked change the resolution emptied goes. So does any record an
    // acceptance empties: it removes content as a direct deletion does, which
    // leaves no record with nothing. Kept, the record could not tell a later
    // resolution it was emptied rather than empty before, and would stay where
    // accepting the same changes at once folds it into an alike neighbour.
    if (
      (isTrackedWrapper(node) || resolution.decision === REVISION_DECISIONS.ACCEPT) &&
      inner.nodes.length === 0 &&
      children.length > 0
    ) {
      seams.push(out.length);
      continue;
    }
    let rebuilt = rebuildNode(node, inner.nodes);
    if (isTrackedWrapper(rebuilt) && rebuilt.resolutionJoins?.retainedAfter !== undefined) {
      const sourceKeys = new Set(identityKeysIn(inner.nodes));
      const retainedAfter = rebuilt.resolutionJoins.retainedAfter.flatMap((entry) => {
        const source: RetainedIdentity["source"][number][] = [];
        const target: RetainedIdentity["target"][number][] = [];
        for (const [slotIndex, slot] of entry.source.entries()) {
          if (!sourceKeys.has(slotKey(slot))) continue;
          const bound = entry.target.at(slotIndex);
          if (bound === undefined)
            panic("Validated cut provenance has paired source and target slots.");
          source.push(slot);
          target.push(bound);
        }
        return source.length === 0 ? [] : [{ depth: entry.depth, source, target }];
      });
      rebuilt = Object.assign({}, rebuilt, {
        resolutionJoins: Object.assign({}, rebuilt.resolutionJoins, { retainedAfter }),
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
    out.push(rebuilt);
  }
  // An inner cut can meet the edge of a container while its other fragment
  // remains beyond a pending insertion. Preserve that exact seam on its
  // actual blocker; accepting the payload cancels it, rejecting restores it.
  for (const [index, node] of out.entries()) {
    if (!isAddedRevision(node) || node.resolutionJoins === undefined) continue;
    if (node.resolutionJoins.remove === 0) continue;
    const left = out.at(index - 1);
    const right = out.at(index + 1);
    const depths = [
      index > 0 && left !== undefined ? resolution.cutEdges.get(left)?.last : undefined,
      right === undefined ? undefined : resolution.cutEdges.get(right)?.first,
    ];
    const groups = [...(node.resolutionJoins.deferredRemove ?? [])];
    for (const depth of depths) {
      if (depth !== undefined) groups.push({ depth, blockers: [node.info.id] });
    }
    if (groups.length === (node.resolutionJoins.deferredRemove?.length ?? 0)) continue;
    out[index] = Object.assign({}, node, {
      resolutionJoins: Object.assign({}, node.resolutionJoins, {
        deferredRemove: canonicalDeferredGroups(groups),
      }),
    });
    changed = true;
  }
  if (!changed) {
    return { nodes: out, changed };
  }
  const edges: CutEdges = {};
  const firstDepth = sourceEdges.get(0);
  const lastDepth = sourceEdges.get(out.length);
  if (firstDepth !== undefined) edges.first = firstDepth;
  if (lastDepth !== undefined) edges.last = lastDepth;
  for (const seam of plainSeams) {
    const left = out.at(seam - 1);
    const right = out.at(seam);
    if (seam > 0 && left?.type === "run" && right?.type === "run" && runsMergeable(left, right))
      seams.push(seam);
  }
  const merges = new Set(seams);
  for (const seam of [...new Set([...seams, ...exactSeams.keys(), ...folds])].toSorted(
    (left, right) => right - left,
  )) {
    const left = out[seam - 1];
    const right = out[seam];
    if (left === undefined || right === undefined) continue;
    let exactDepth = exactSeams.get(seam);
    if (exactDepth !== undefined) {
      // Extend only an already recorded outer cut with an unconsumed inner
      // source cut. Equal authored containers alone never establish a seam.
      exactDepth = Math.max(
        exactDepth,
        resolution.cutEdges.get(left)?.last ?? 0,
        resolution.cutEdges.get(right)?.first ?? 0,
      );
    }
    // Later independent edits may make a recorded seam non-alike. Join
    // only the recorded depth and matching fields, preserving those edits.
    let met: InlineNode[];
    if (merges.has(seam)) {
      met = mergeAtSeam(left, right, resolution.emptied);
    } else if (exactDepth === undefined) {
      met = foldAtSeam(left, right, resolution.emptied);
    } else {
      // Earlier inline resolution can empty an original cut fragment. Its
      // surviving identity folds back before measuring the remaining depth.
      const folded = foldAtSeam(left, right, resolution.emptied);
      met =
        folded.length === 1
          ? folded
          : (mergeLists([left], [right], exactDepth, { mode: "asFarAsAlike" }) ??
            panic("An as-far-as-alike merge always returns its records."));
    }
    if (met.length === 1) {
      const merged = met.at(0) ?? panic("A single merged record exists.");
      const outerEdges: CutEdges = {};
      const first = resolution.cutEdges.get(left)?.first;
      const last = resolution.cutEdges.get(right)?.last;
      if (first !== undefined) outerEdges.first = first;
      if (last !== undefined) outerEdges.last = last;
      resolution.cutEdges.set(merged, outerEdges);
    }
    out.splice(seam - 1, 2, ...met);
  }
  const first = out.at(0);
  const last = out.at(-1);
  const nestedFirst = first === undefined ? undefined : resolution.cutEdges.get(first)?.first;
  const nestedLast = last === undefined ? undefined : resolution.cutEdges.get(last)?.last;
  if (nestedFirst !== undefined) edges.first = Math.max(edges.first ?? 0, nestedFirst);
  if (nestedLast !== undefined) edges.last = Math.max(edges.last ?? 0, nestedLast);
  return { nodes: out, changed, edges };
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
const emptiedEndsOf = (nodes: readonly InlineNode[], emptied: Emptied): EmptiedEnds => {
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
    const inner = deferSplitSeam({ left: lastChildren, right: firstChildren, depth: depth - 1 });
    if (inner.left === lastChildren && inner.right === firstChildren) return { left, right };
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
  while (rightStart < right.length && blocking(right.at(rightStart))) rightStart += 1;
  // Pending revisions can themselves be the two source fragments. Keep
  // those outside the blocker group; only insertions between them defer it.
  let sourcePair = false;
  for (let leftIndex = left.length - 1; leftIndex > leftEnd && !sourcePair; leftIndex -= 1) {
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
    const blockers = [...left.slice(leftEnd + 1), ...right.slice(0, rightStart)].flatMap((node) =>
      isTrackedWrapper(node) ? [node.info.id] : [],
    );
    const deferred = (node: InlineNode): InlineNode => {
      if (!isTrackedWrapper(node) || node.resolutionJoins === undefined)
        panic("Only a provenance-bearing pending insertion blocks this split seam.");
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
        resolutionJoins: Object.assign({}, node.resolutionJoins, { deferredRemove: groups }),
      });
    };
    return {
      left: left.map((node, index) => (index > leftEnd ? deferred(node) : node)),
      right: right.map((node, index) => (index < rightStart ? deferred(node) : node)),
    };
  }
  return { left, right };
};

/**
 * The operations that remove a paragraph's resolved mark, against the document
 * as it stands. `ends` is updated for the paragraph a join leaves.
 */
const joinOps = (
  document: Document,
  { op, story, paraId, added, depth }: JoinPlan,
  ends: EmptiedEndsById,
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
  const blocks = blockListAt(body.content, location.list);
  const next = blocks[location.index + 1];
  const empty = paragraphLength(paragraph) === 0;
  if (
    next?.type === "table" &&
    next.rows.some((row) => row.structuralChange?.type === "tableRowInsertion")
  ) {
    return Result.err(
      refusal(
        op,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        "Resolve the inserted table before removing its preceding paragraph break.",
      ),
    );
  }
  if (next?.type === "paragraph") {
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
      const seam = deferSplitSeam({ left: paragraph.content, right: next.content, depth });
      const left = seam.left.at(-1);
      const right = seam.right.at(0);
      if (depth > 0 && left !== undefined && right !== undefined) {
        const folded = foldAtSeam(left, right, emptied);
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
    // A record left as it was keeps its identity through the merge; an emptied
    // one folded into its neighbour is rebuilt, and is no longer empty.
    ends.set(nextId, emptiedEndsOf(merged, emptied));
    const survivor: Paragraph = { ...next, content: asParagraphContent(merged) };
    return Result.ok([
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story,
        expected: [paragraph, next],
        blocks: [survivor],
        ...(paragraph.sectionProperties === undefined
          ? {}
          : { sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE }),
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
  // A cell-ending mark describes the terminator, not a paragraph to retire.
  const cellFinal =
    endsItsContainer(body, location) && location.list.some((step) => step.kind === "tableCell");
  const canGo = !cellFinal && (added || !endsItsContainer(body, location));
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
      const fitting = edge === "first" ? node.resolutionJoins.after : node.resolutionJoins.before;
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
 * Resolve tracked changes by revision id, applied as the primitive
 * operations it expands to. An id no record in the story carries is
 * skipped; one carried by a record resolution does not reach yet (a table,
 * cell or section change, a record inside a field) is refused.
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
  const resolution: Resolution = {
    ids,
    decision: op.decision,
    emptied: new WeakSet(),
    acceptedClosed: { first: new WeakSet(), last: new WeakSet() },
    cutEdges: new WeakMap(),
  };
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

  const blockerIds = new Set<number>();
  const collectBlockers = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node) && isAddedRevision(node) && node.resolutionJoins !== undefined)
        blockerIds.add(node.info.id);
      collectBlockers(childNodes(node) ?? []);
    }
  };
  for (const { paragraph } of paragraphs) collectBlockers(paragraph.content);
  const validProvenance = (nodes: readonly InlineNode[]): boolean =>
    nodes.every(
      (node) =>
        validRetainedIdentities(node) &&
        (!isTrackedWrapper(node) ||
          node.resolutionJoins?.acceptance === undefined ||
          node.resolutionJoins.acceptance === INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS) &&
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
        (depth !== undefined && (!Number.isInteger(depth) || depth < 0 || depth > MAX_REVISION_ID))
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
  const emptiedEnds: EmptiedEndsById = new Map();
  const resolvedContents = new Map<string, readonly InlineNode[]>();
  const resolvedCutEdges = new Map<string, CutEdges>();
  for (const { paragraph } of paragraphs) {
    const paraId = paragraph.paraId ?? "";
    const prepared = prepareDeferredRemovals(paragraph.content, resolution);
    const resolved = resolveList(prepared, resolution);
    resolvedContents.set(idKey(paraId), resolved.nodes);
    if (resolved.edges !== undefined) resolvedCutEdges.set(idKey(paraId), resolved.edges);
    if (resolved.changed || prepared !== paragraph.content) {
      emptiedEnds.set(idKey(paraId), emptiedEndsOf(resolved.nodes, resolution.emptied));
      inline.push({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: op.story,
        blockId: paraId,
        expected: paragraph.content,
        content: asParagraphContent(resolved.nodes),
      });
    }
  }
  const retiredCutDepths = new Map<string, number>();
  const body = storyBody(document, op.story);
  for (const location of paragraphs) {
    const paragraph = location.paragraph;
    const mark = paragraph.pPrMark;
    if (mark === undefined || !markWasAdded(mark.kind) || mark.resolutionJoin === undefined)
      continue;
    const trailing = acceptedClosedBoundary({
      nodes:
        resolvedContents.get(idKey(paragraph.paraId ?? "")) ??
        panic("Every source paragraph has a resolved content result."),
      edge: "last",
      resolution,
    });
    const next = blockListAt(body.content, location.list).at(location.index + 1);
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
      next?.type === "paragraph" ? (resolvedCutEdges.get(idKey(next.paraId ?? ""))?.first ?? 0) : 0;
    const following = blockListAt(body.content, location.list);
    for (let index = location.index + 1; index < following.length; index += 1) {
      const candidate = following.at(index);
      if (candidate?.type !== "paragraph") break;
      leadingCut = Math.max(
        leadingCut,
        resolvedCutEdges.get(idKey(candidate.paraId ?? ""))?.first ?? 0,
      );
      if (
        (resolvedContents.get(idKey(candidate.paraId ?? ""))?.length ?? 0) !== 0 ||
        candidate.pPrMark === undefined ||
        !markWasAdded(candidate.pPrMark.kind)
      )
        break;
    }
    let depth = Math.max(
      mark.resolutionJoin,
      resolvedCutEdges.get(idKey(paragraph.paraId ?? ""))?.last ?? 0,
      leadingCut,
    );
    if (trailing !== undefined) depth = Math.min(depth, trailing);
    if (leading !== undefined) depth = Math.min(depth, leading);
    if (depth !== mark.resolutionJoin) retiredCutDepths.set(idKey(paragraph.paraId ?? ""), depth);
  }
  for (const location of paragraphs) {
    const { paragraph } = location;
    const paraId = paragraph.paraId ?? "";
    const mark = paragraph.pPrMark;
    const markResolved = mark !== undefined && ids.has(mark.info.id);
    const added = mark !== undefined && markWasAdded(mark.kind);
    const keepsBreak = markResolved && added === (op.decision === REVISION_DECISIONS.ACCEPT);
    if (markResolved && !keepsBreak) {
      let depth = retiredCutDepths.get(idKey(paraId)) ?? mark.resolutionJoin;
      if (added && depth !== undefined && op.decision === REVISION_DECISIONS.REJECT) {
        const next = blockListAt(body.content, location.list).at(location.index + 1);
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
      const remainingMark = review === undefined ? paragraph.pPrMark : review.pPrMark;
      if (remainingMark !== undefined) {
        review ??= reviewFieldsOf(paragraph);
        review.pPrMark = Object.assign({}, remainingMark, { resolutionJoin: retiredDepth });
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
  const rows = resolveTableRows({ document: staged.value.document, op, applyOps });
  if (rows.isErr()) return Result.err(rows.error);
  const removedParagraphs = new Set(rows.value.touched.removed.map(idKey));
  const edits: DocumentEdit[] = [staged.value, rows.value];
  let current = rows.value.document;
  for (const join of joins.toReversed()) {
    if (removedParagraphs.has(idKey(join.paraId))) continue;
    const planned = joinOps(current, join, emptiedEnds);
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
