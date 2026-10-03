/**
 * How an edit keeps the ids of identified records unique (see `slots.ts`).
 *
 * Cutting an identified record in two would leave both halves with its one
 * id, and operations mint no ids, so the operation that makes the cut names
 * the new ones (`newIds`): the half holding the start keeps the id, and every
 * later half takes the next new id, in document order. Merging two identified
 * records keeps the first one's ids and retires the second's; the inverse
 * that cuts them apart again names the retired ids, so it recreates them
 * exactly.
 */

import { panic } from "better-result";

import { MAX_REVISION_ID, type Paragraph, type Deletion } from "../model/document";
import { IDENTITY_SPACES, type IdentitySlot, slotKey } from "./ids";
import { asParagraphContent, childNodes, type InlineNode, rebuildNode } from "./leaves";
import { identitySlots, withInlineIdentity, withParagraphIdentity } from "./slots";
import type { NewIds } from "./types";
import { isAddedRevision, isTrackedWrapper } from "./review";

const visitNodes = (nodes: readonly InlineNode[], visit: (node: InlineNode) => void): void => {
  for (const node of nodes) {
    visit(node);
    visitNodes(childNodes(node) ?? [], visit);
  }
};

/** The slot keys of paragraphs and everything in them, in document order. */
export const identityKeysOf = (paragraphs: readonly Paragraph[]): string[] => {
  const out: string[] = [];
  for (const paragraph of paragraphs) {
    for (const slot of identitySlots(paragraph)) out.push(slotKey(slot));
    visitNodes(paragraph.content, (node) => {
      for (const slot of identitySlots(node)) out.push(slotKey(slot));
    });
  }
  return out;
};

type DeferredRemovalGroup = NonNullable<
  NonNullable<Deletion["resolutionJoins"]>["deferredRemove"]
>[number];
/** Identity substitutions can collapse distinct source groups into one exact fact. */
export const canonicalDeferredGroups = (
  groups: readonly DeferredRemovalGroup[],
): DeferredRemovalGroup[] => {
  const keys = new Set<string>();
  const out: DeferredRemovalGroup[] = [];
  for (const group of groups) {
    const key = `${group.depth}:${group.blockers.join(",")}`;
    if (keys.has(key)) continue;
    keys.add(key);
    out.push(group);
  }
  return out;
};

type BindDeferredIdentitiesOptions = { before: readonly Paragraph[]; after: readonly Paragraph[] };
/** Deferred seam references follow the actual source wrapper's freshened fragments. */
const bindDeferredIdentities = ({ before, after }: BindDeferredIdentitiesOptions): Paragraph[] => {
  const references = new Map<number, Set<number>>();
  type CollectOptions = { oldNodes: readonly InlineNode[]; freshNodes: readonly InlineNode[] };
  const collect = ({ oldNodes, freshNodes }: CollectOptions): void => {
    for (const [index, old] of oldNodes.entries()) {
      const fresh = freshNodes.at(index) ?? panic("Freshening preserves each source position.");
      if (isTrackedWrapper(old) && isAddedRevision(old) && isTrackedWrapper(fresh)) {
        const ids = references.get(old.info.id) ?? new Set<number>();
        ids.add(fresh.info.id);
        references.set(old.info.id, ids);
      }
      collect({ oldNodes: childNodes(old) ?? [], freshNodes: childNodes(fresh) ?? [] });
    }
  };
  for (const [index, old] of before.entries()) {
    const fresh = after.at(index) ?? panic("Freshening preserves each source paragraph position.");
    collect({ oldNodes: old.content, freshNodes: fresh.content });
  }
  const rewrite = (nodes: readonly InlineNode[]): readonly InlineNode[] => {
    let changed = false;
    const out: InlineNode[] = [];
    for (const node of nodes) {
      let own = node;
      if (isTrackedWrapper(node) && node.resolutionJoins?.deferredRemove !== undefined) {
        const groups = node.resolutionJoins.deferredRemove;
        if (
          Array.isArray(groups) &&
          groups.every(
            (group) => group !== null && typeof group === "object" && Array.isArray(group.blockers),
          )
        ) {
          const remapGroup = (group: DeferredRemovalGroup) => ({
            depth: group.depth,
            blockers: [
              ...new Set(group.blockers.flatMap((id) => [...(references.get(id) ?? [id])])),
            ],
          });
          const deferredRemove = canonicalDeferredGroups(groups.map(remapGroup));
          own = Object.assign({}, node, {
            resolutionJoins: Object.assign({}, node.resolutionJoins, { deferredRemove }),
          });
        }
      }
      const children = childNodes(own);
      const rewritten = children === undefined ? children : rewrite(children);
      const result = rewritten === children ? own : rebuildNode(own, rewritten ?? []);
      changed ||= result !== node;
      out.push(result);
    }
    return changed ? out : nodes;
  };
  return after.map((paragraph) => {
    const content = rewrite(paragraph.content);
    return content === paragraph.content
      ? paragraph
      : Object.assign({}, paragraph, { content: asParagraphContent(content) });
  });
};

export type FreshenOptions = {
  /** The paragraphs the operation replaces. */
  before: readonly Paragraph[];
  /** The paragraphs it puts in their place. */
  after: readonly Paragraph[];
  newIds: NewIds;
  /** Slot keys used anywhere in the package outside `before`; asked for only when needed. */
  usedElsewhere: () => ReadonlySet<string>;
  /** Records holding only content the edit inserted: they yield an id to records that held it before. */
  inserted?: ReadonlySet<object>;
  /** Pending cut references must not be reused by a fresh identity. */
  reserved?: ReadonlySet<string>;
};

export type FreshenOutcome =
  | { kind: "fresh"; paragraphs: Paragraph[] }
  | { kind: "needsIds"; missing: number }
  | { kind: "invalidId"; id: number };

const recordsOf = (paragraphs: readonly Paragraph[]): Set<object> => {
  const out = new Set<object>(paragraphs);
  for (const paragraph of paragraphs) visitNodes(paragraph.content, (node) => out.add(node));
  return out;
};

type RetainedIdentity = NonNullable<
  NonNullable<Deletion["resolutionJoins"]>["retainedAfter"]
>[number];

const matchesSlots = (node: InlineNode, expected: RetainedIdentity["target"]): boolean => {
  const slots = identitySlots(node);
  return expected.every((target) =>
    slots.some((slot) => slot.space === target.space && slot.id === target.id),
  );
};

/** Bind the exact right fragment after its source identities were freshened. */
const bindRetainedIdentities = (
  before: readonly InlineNode[],
  after: readonly InlineNode[],
): InlineNode[] => {
  type AtDepthOptions = {
    old: InlineNode;
    fresh: InlineNode;
    entry: RetainedIdentity;
    depth: number;
  };
  const atDepth = ({
    old,
    fresh,
    entry,
    depth,
  }: AtDepthOptions): readonly IdentitySlot[] | undefined => {
    if (depth === entry.depth) {
      if (!matchesSlots(old, entry.target)) return undefined;
      const oldSlots = identitySlots(old);
      const freshSlots = identitySlots(fresh);
      return entry.target.map((target) => {
        const index = oldSlots.findIndex(
          (slot) => slot.space === target.space && slot.id === target.id,
        );
        return freshSlots.at(index) ?? target;
      });
    }
    const oldChildren = childNodes(old) ?? [];
    const freshChildren = childNodes(fresh) ?? [];
    for (const [index, child] of oldChildren.entries()) {
      const freshChild = freshChildren.at(index);
      if (freshChild === undefined) continue;
      const slots = atDepth({ old: child, fresh: freshChild, entry, depth: depth + 1 });
      if (slots !== undefined) return slots;
    }
    return undefined;
  };
  return after.map((node, index) => {
    if (node.type !== "deletion" && node.type !== "moveFrom") return node;
    const joins = node.resolutionJoins;
    if (joins === undefined || joins.retainedAfter === undefined) return node;
    const retained = joins.retainedAfter;
    const oldRight = before.at(index + 1);
    const freshRight = after.at(index + 1);
    if (oldRight === undefined || freshRight === undefined) return node;
    const oldOwner = before.at(index) ?? panic("Freshening preserves each source record position.");
    const bound = retained.map((entry) => {
      const target = atDepth({ old: oldRight, fresh: freshRight, entry, depth: 0 });
      let source = entry.source;
      {
        const oldChildren = childNodes(oldOwner) ?? [];
        const freshChildren = childNodes(node) ?? [];
        for (const [childIndex, old] of oldChildren.entries()) {
          const fresh = freshChildren.at(childIndex);
          if (fresh === undefined) panic("Freshening preserves every source record position.");
          const rebound = atDepth({
            old,
            fresh,
            entry: { depth: entry.depth, source: entry.source, target: entry.source },
            depth: 0,
          });
          if (rebound !== undefined) {
            source = rebound;
            break;
          }
        }
      }
      return { depth: entry.depth, source, target: target ?? entry.target };
    });
    return Object.assign({}, node, {
      resolutionJoins: Object.assign({}, joins, { retainedAfter: bound }),
    });
  });
};

type CutPair = { old: InlineNode; fresh: InlineNode };

/** A cut shares the original hints object; independent wrappers never do. */
type BindCutRetainedIdentitiesOptions = {
  before: readonly Paragraph[];
  after: readonly Paragraph[];
};
const bindCutRetainedIdentities = ({
  before,
  after,
}: BindCutRetainedIdentitiesOptions): Paragraph[] => {
  const groups = new Map<NonNullable<Deletion["resolutionJoins"]>, CutPair[]>();
  type PairedNodesOptions = { oldNodes: readonly InlineNode[]; freshNodes: readonly InlineNode[] };
  const collect = ({ oldNodes, freshNodes }: PairedNodesOptions): void => {
    for (const [index, old] of oldNodes.entries()) {
      const fresh = freshNodes.at(index);
      if (fresh === undefined) panic("Freshening preserves every source record position.");
      if (
        (old.type === "deletion" || old.type === "moveFrom") &&
        old.resolutionJoins?.retainedAfter !== undefined
      ) {
        const pieces = groups.get(old.resolutionJoins) ?? [];
        pieces.push({ old, fresh });
        groups.set(old.resolutionJoins, pieces);
      }
      collect({ oldNodes: childNodes(old) ?? [], freshNodes: childNodes(fresh) ?? [] });
    }
  };
  for (const [index, old] of before.entries()) {
    const fresh = after.at(index);
    if (fresh === undefined) panic("Freshening preserves every source paragraph position.");
    collect({ oldNodes: old.content, freshNodes: fresh.content });
  }
  type PairedSlots = { depth: number; old: IdentitySlot[]; fresh: IdentitySlot[] };
  const slotsIn = (pair: CutPair): PairedSlots[] => {
    const slots: PairedSlots[] = [];
    type VisitOptions = PairedNodesOptions & { depth: number };
    const visit = ({ oldNodes, freshNodes, depth }: VisitOptions): void => {
      for (const [index, old] of oldNodes.entries()) {
        const fresh = freshNodes.at(index);
        if (fresh === undefined) panic("Freshening preserves every source record position.");
        const oldSlots = identitySlots(old);
        if (oldSlots.length > 0) slots.push({ depth, old: oldSlots, fresh: identitySlots(fresh) });
        visit({
          oldNodes: childNodes(old) ?? [],
          freshNodes: childNodes(fresh) ?? [],
          depth: depth + 1,
        });
      }
    };
    visit({
      oldNodes: childNodes(pair.old) ?? [],
      freshNodes: childNodes(pair.fresh) ?? [],
      depth: 0,
    });
    return slots;
  };
  const replacements = new Map<InlineNode, InlineNode>();
  for (const [joins, pieces] of groups) {
    for (const [index, piece] of pieces.entries()) {
      const fresh = piece.fresh;
      if (fresh.type !== "deletion" && fresh.type !== "moveFrom") continue;
      const sources = slotsIn(piece);
      const next = pieces.at(index + 1);
      const retained: RetainedIdentity[] = [];
      if (next !== undefined) {
        const targets = slotsIn(next);
        // Every identified record duplicated by this exact cut transfers its
        // identity to the next piece, including newly introduced inner seams.
        for (const source of sources) {
          const target = targets.find(
            (candidate) =>
              candidate.depth === source.depth &&
              candidate.old.some((slot) =>
                source.old.some((own) => slotKey(own) === slotKey(slot)),
              ),
          );
          if (target === undefined) continue;
          const sourceIds: IdentitySlot[] = [];
          const targetIds: IdentitySlot[] = [];
          for (const [slotIndex, slot] of source.old.entries()) {
            const targetIndex = target.old.findIndex((other) => slotKey(slot) === slotKey(other));
            if (targetIndex < 0) continue;
            const sourceId = source.fresh.at(slotIndex);
            const targetId = target.fresh.at(targetIndex);
            if (sourceId === undefined || targetId === undefined)
              panic("Freshening preserves the identity slot layout of each source record.");
            sourceIds.push(sourceId);
            targetIds.push(targetId);
          }
          if (sourceIds.length > 0)
            retained.push({ depth: source.depth, source: sourceIds, target: targetIds });
        }
      } else {
        for (const entry of fresh.resolutionJoins?.retainedAfter ?? joins.retainedAfter ?? []) {
          const source = sources.find(
            (candidate) =>
              candidate.depth === entry.depth &&
              entry.source.every(
                (slot) =>
                  candidate.old.some((own) => slotKey(own) === slotKey(slot)) ||
                  candidate.fresh.some((own) => slotKey(own) === slotKey(slot)),
              ),
          );
          if (source === undefined) continue;
          const rebound = entry.source.map(
            (slot) =>
              source.fresh.at(
                source.old.findIndex(
                  (own, slotIndex) =>
                    slotKey(own) === slotKey(slot) ||
                    slotKey(source.fresh.at(slotIndex) ?? own) === slotKey(slot),
                ),
              ) ?? panic("Freshening preserves every matched source identity slot."),
          );
          retained.push({ depth: entry.depth, source: rebound, target: entry.target });
        }
      }
      replacements.set(
        fresh,
        Object.assign({}, fresh, {
          resolutionJoins: Object.assign({}, fresh.resolutionJoins ?? joins, {
            retainedAfter: retained,
          }),
        }),
      );
    }
  }
  const rewrite = (nodes: readonly InlineNode[]): readonly InlineNode[] => {
    let changed = false;
    const out: InlineNode[] = [];
    for (const node of nodes) {
      const own = replacements.get(node) ?? node;
      const children = childNodes(own);
      const rewritten = children === undefined ? children : rewrite(children);
      const result = rewritten === children ? own : rebuildNode(own, rewritten ?? []);
      changed ||= result !== node;
      out.push(result);
    }
    return changed ? out : nodes;
  };
  return after.map((paragraph) => {
    const content = rewrite(paragraph.content);
    return content === paragraph.content
      ? paragraph
      : Object.assign({}, paragraph, { content: asParagraphContent(content) });
  });
};

/**
 * Give fresh ids to the records an edit created that would otherwise share an
 * id. A record carried over from `before` unchanged keeps its ids. A record
 * the edit created (a half of a cut record, inserted content, a rebuilt
 * container) keeps an id no other record holds, and takes the next of
 * `newIds`, in document order, for one another record holds: the first half
 * of a cut record keeps the id, and the halves after it take new ones. Where
 * a record holding content that was there before and a record holding only
 * inserted content share an id, the first keeps it: the id stays with the
 * content it named.
 *
 * The seed contract makes every id unique, so the only records that share one
 * are those the edit created.
 */
export const freshenIdentities = ({
  before,
  after,
  newIds,
  usedElsewhere,
  inserted = new Set(),
  reserved = new Set(),
}: FreshenOptions): FreshenOutcome => {
  const carried = recordsOf(before);
  const fixed = new Set<string>();
  const createdKeys: string[] = [];
  const note = (record: object, slots: readonly IdentitySlot[]): void => {
    for (const slot of slots) {
      if (carried.has(record)) {
        fixed.add(slotKey(slot));
      } else {
        createdKeys.push(slotKey(slot));
      }
    }
  };
  for (const paragraph of after) {
    note(paragraph, identitySlots(paragraph));
    visitNodes(paragraph.content, (node) => note(node, identitySlots(node)));
  }
  const repeats = new Set<string>();
  const needsFresh = createdKeys.some((key) => {
    const repeated = fixed.has(key) || repeats.has(key);
    repeats.add(key);
    return repeated;
  });
  if (!needsFresh && createdKeys.length === 0) {
    return { kind: "fresh", paragraphs: [...after] };
  }

  const elsewhere = usedElsewhere();
  if (!needsFresh && !createdKeys.some((key) => elsewhere.has(key))) {
    return { kind: "fresh", paragraphs: [...after] };
  }
  // Which created record keeps each id: those holding earlier content first,
  // then those holding only inserted content, each in document order.
  const keepers = new Map<object, Set<number>>();
  const claimed = new Set<string>();
  const created: [object, IdentitySlot[]][] = [];
  const collect = (record: object, slots: IdentitySlot[]): void => {
    if (!carried.has(record) && slots.length > 0) created.push([record, slots]);
  };
  for (const paragraph of after) {
    collect(paragraph, identitySlots(paragraph));
    visitNodes(paragraph.content, (node) => collect(node, identitySlots(node)));
  }
  for (const [record, slots] of [
    ...created.filter(([candidate]) => !inserted.has(candidate)),
    ...created.filter(([candidate]) => inserted.has(candidate)),
  ]) {
    for (const [index, slot] of slots.entries()) {
      const key = slotKey(slot);
      if (elsewhere.has(key) || fixed.has(key) || claimed.has(key)) continue;
      claimed.add(key);
      const kept = keepers.get(record) ?? new Set<number>();
      kept.add(index);
      keepers.set(record, kept);
    }
  }
  const taken = new Set([...elsewhere, ...identityKeysOf(before), ...identityKeysOf(after)]);
  const pools = {
    [IDENTITY_SPACES.REVISION]: newIds.revision ?? [],
    [IDENTITY_SPACES.CONTROL]: newIds.control ?? [],
  };
  const next = { [IDENTITY_SPACES.REVISION]: 0, [IDENTITY_SPACES.CONTROL]: 0 };
  let missing = 0;
  let invalid: number | undefined;

  const idsFor = (record: object, slots: readonly IdentitySlot[]): number[] | undefined => {
    if (carried.has(record)) return undefined;
    let changed = false;
    const ids = slots.map((slot, index) => {
      if (keepers.get(record)?.has(index) === true) {
        return slot.id;
      }
      // A new id a record already carries (content the operation brings
      // restores it, say) is passed over: the next one is for this record.
      const pool = pools[slot.space];
      const isTaken = (id: number) => {
        const key = slotKey({ space: slot.space, id });
        return taken.has(key) || reserved.has(key);
      };
      let fresh = pool[next[slot.space]];
      while (fresh !== undefined && Number.isInteger(fresh) && isTaken(fresh)) {
        next[slot.space] += 1;
        fresh = pool[next[slot.space]];
      }
      next[slot.space] += 1;
      if (fresh === undefined) {
        missing += 1;
        return slot.id;
      }
      if (!Number.isInteger(fresh) || fresh < 0 || fresh > MAX_REVISION_ID) {
        invalid ??= fresh;
        return slot.id;
      }
      taken.add(slotKey({ space: slot.space, id: fresh }));
      changed = true;
      return fresh;
    });
    return changed ? ids : undefined;
  };

  const freshenNodes = (nodes: readonly InlineNode[]): readonly InlineNode[] => {
    let changed = false;
    const out: InlineNode[] = [];
    for (const node of nodes) {
      const ids = idsFor(node, identitySlots(node));
      const own = ids === undefined ? node : withInlineIdentity(node, ids);
      const children = childNodes(node);
      const freshened = children === undefined ? children : freshenNodes(children);
      const result = freshened === children ? own : rebuildNode(own, freshened ?? []);
      changed ||= result !== node;
      out.push(result);
    }
    return changed ? bindRetainedIdentities(nodes, out) : nodes;
  };

  const paragraphs: Paragraph[] = [];
  for (const paragraph of after) {
    const ids = idsFor(paragraph, identitySlots(paragraph));
    const own = ids === undefined ? paragraph : withParagraphIdentity(paragraph, ids);
    const content = freshenNodes(paragraph.content);
    paragraphs.push(
      content === paragraph.content ? own : { ...own, content: asParagraphContent(content) },
    );
  }
  if (missing > 0) {
    return { kind: "needsIds", missing };
  }
  return invalid === undefined
    ? {
        kind: "fresh",
        paragraphs: bindDeferredIdentities({
          before: after,
          after: bindCutRetainedIdentities({ before: after, after: paragraphs }),
        }),
      }
    : { kind: "invalidId", id: invalid };
};
