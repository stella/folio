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

import { MAX_REVISION_ID, type Paragraph } from "../model/document";
import { IDENTITY_SPACES, type IdentitySlot, slotKey } from "./ids";
import { asParagraphContent, childNodes, type InlineNode, rebuildNode } from "./leaves";
import { identitySlots, withInlineIdentity, withParagraphIdentity } from "./slots";
import type { NewIds } from "./types";

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
      const isTaken = (id: number) => taken.has(slotKey({ space: slot.space, id }));
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
    return changed ? out : nodes;
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
  return invalid === undefined ? { kind: "fresh", paragraphs } : { kind: "invalidId", id: invalid };
};
