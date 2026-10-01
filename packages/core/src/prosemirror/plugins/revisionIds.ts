/**
 * Tracked-revision id minting (`w:id` on `<w:ins>` / `<w:del>`).
 *
 * Uniqueness comes from `seedRevisionIdsAbove` (called on load with the
 * document's own max id), NOT from the counter's starting value. A clock seed
 * bought uniqueness at the cost of ids no consumer could read — see
 * `MAX_REVISION_ID`. Port of eigenpal/docx-editor#1093.
 *
 * Kept free of `suggestionMode` imports so the suggestion plugin can mint
 * without a load-order cycle.
 */

import type { Node as PmNode } from "prosemirror-model";
import { TaggedError } from "better-result";

import { MAX_REVISION_ID } from "@stll/docx-core/model";

/**
 * Next id to hand out. Starts at 1 (0 is the serializer's "unusable metadata"
 * fallback) and is raised past a loaded document's existing ids by
 * `seedRevisionIdsAbove`.
 *
 * NOT seeded from a clock. `Date.now()` (~1.8e12) overflows the signed 32-bit
 * int consumers read `w:id` into.
 */
let counter = 1;
const occupiedIds = new Set<number>();

export class RevisionIdAllocationError extends TaggedError("RevisionIdAllocationError")<{
  message: string;
  first: number;
  next: number;
}> {}

/** Batches use contiguous ids; on rollover prefer the largest unused interval. */
const largestFreeIntervalStart = (reservedIds: ReadonlySet<number>): number => {
  let first = 1;
  let bestFirst = 1;
  let bestSize = 0;
  const boundaries = [...reservedIds].filter((id) => id > 0).toSorted((a, b) => a - b);
  boundaries.push(MAX_REVISION_ID + 1);
  for (const boundary of boundaries) {
    const size = boundary - first;
    if (size > bestSize) {
      bestFirst = first;
      bestSize = size;
    }
    first = boundary + 1;
  }
  if (bestSize === 0) {
    throw new RevisionIdAllocationError({
      message: "Revision id space exhausted",
      first: 1,
      next: 1,
    });
  }
  return bestFirst;
};

/** Deterministic package-local batches prefer ids above loaded revisions, then a free interval. */
export const revisionIdSeedAbove = (existingIds: Iterable<number>): number => {
  const reservedIds = new Set<number>();
  let maxId = 0;
  for (const id of existingIds) {
    if (!Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID) continue;
    reservedIds.add(id);
    maxId = Math.max(maxId, id);
  }
  return maxId < MAX_REVISION_ID ? maxId + 1 : largestFreeIntervalStart(reservedIds);
};

/** The next free id, including after the signed 32-bit boundary wraps. */
export const nextRevisionId = (): number => {
  if (counter > MAX_REVISION_ID) counter = largestFreeIntervalStart(occupiedIds);
  while (occupiedIds.has(counter)) {
    counter += 1;
    if (counter > MAX_REVISION_ID) counter = largestFreeIntervalStart(occupiedIds);
    if (occupiedIds.size >= MAX_REVISION_ID) {
      throw new RevisionIdAllocationError({
        message: "Revision id space exhausted",
        first: counter,
        next: counter,
      });
    }
  }
  return counter;
};

/**
 * A shared batch increments its seed locally, so a single free id is not
 * enough: start in the largest free interval. Every contiguous batch that
 * can fit anywhere fits here, including after loaded ids shorten the tail.
 * Continue within this interval until another producer changes the cursor.
 */
export const nextRevisionIdRange = (): number => {
  counter = largestFreeIntervalStart(occupiedIds);
  return counter;
};

const validateRevisionIdRange = (first: number, next: number): void => {
  if (
    !Number.isInteger(first) ||
    !Number.isInteger(next) ||
    first < 0 ||
    next < first ||
    next > MAX_REVISION_ID + 1
  ) {
    throw new RevisionIdAllocationError({
      message: "Revision id range exceeds the OOXML limit",
      first,
      next,
    });
  }
};

/** Reserve a deterministic batch without changing its caller-supplied ids. */
export const reserveRevisionIds = (first: number, next: number): void => {
  validateRevisionIdRange(first, next);
  for (let id = first; id < next; id += 1) occupiedIds.add(id);
  if (next > counter) counter = next;
};

/** Shared producers must never cross a loaded or already issued id on wrap. */
export const claimRevisionIds = (first: number, next: number): void => {
  validateRevisionIdRange(first, next);
  for (let id = first; id < next && id <= MAX_REVISION_ID; id += 1) {
    if (occupiedIds.has(id)) {
      throw new RevisionIdAllocationError({
        message: "Revision id range overlaps an occupied id",
        first,
        next,
      });
    }
  }
  reserveRevisionIds(first, next);
};

/** Mint a bounded id without reusing loaded or previously issued ids. */
export function mintRevisionId(): number {
  const id = nextRevisionId();
  claimRevisionIds(id, id + 1);
  return id;
}

/** Reserve a valid loaded id and prefer the range above it. */
export function seedRevisionIdsAbove(maxId: number): void {
  if (!Number.isInteger(maxId) || maxId < 0 || maxId > MAX_REVISION_ID) return;
  occupiedIds.add(maxId);
  if (maxId >= counter) counter = maxId + 1;
}

/** Node attrs that carry a revision id directly (`revisionId` field). */
const DIRECT_REVISION_ATTR_KEYS = ["trIns", "trDel"] as const;

/**
 * The highest id in `doc` from any part of the annotation space.
 *
 * Preserve the editor's conservative annotation seeding: revisions reserve
 * comment and bookmark ids too, although OOXML defines separate id spaces.
 */
const visitAnnotationIds = (doc: PmNode, visit: (id: number) => void): number => {
  let max = 0;

  const consider = (id: unknown): void => {
    if (typeof id === "number" && Number.isInteger(id) && id >= 0 && id <= MAX_REVISION_ID) {
      visit(id);
      max = Math.max(max, id);
    }
  };

  const considerMetadata = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    for (const [key, child] of Object.entries(value)) {
      if (key === "id" || key === "revisionId" || key === "commentId") consider(child);
      else considerMetadata(child);
    }
  };

  doc.descendants((node) => {
    for (const mark of node.marks) {
      consider(mark.attrs["revisionId"]);
      consider(mark.attrs["commentId"]);
      considerMetadata(mark.attrs["_docxRevisionAncestors"]);
      considerMetadata(mark.attrs["changes"]);
    }

    if (node.type.name === "bookmarkBoundary") {
      consider(node.attrs["id"]);
    }

    const attrs = node.attrs;
    considerMetadata(attrs["_propertyChanges"]);

    considerMetadata(attrs["pPrMark"]);
    for (const key of DIRECT_REVISION_ATTR_KEYS) considerMetadata(attrs[key]);
    considerMetadata(attrs["cellMarker"]);
  });

  return max;
};

export const maxAnnotationIdInDoc = (doc: PmNode): number => visitAnnotationIds(doc, () => {});

/**
 * Raise the counter above every annotation id already present in `doc`.
 * Called from the suggestion-mode plugin's `state.init`.
 */
export function seedRevisionIdsFromDoc(doc: PmNode): void {
  visitAnnotationIds(doc, seedRevisionIdsAbove);
}
