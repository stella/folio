/**
 * Id leases: disjoint blocks of fresh ids an author may mint.
 *
 * Several authors editing one document each mint new paragraph, revision,
 * content-control and comment ids. Minting the lowest free id, as a single
 * editor does, makes two concurrent authors pick the same one. A lease names
 * the blocks one author draws from; a sequencer grants blocks that no other
 * author holds (see `sequencing/idLeases.ts`), so ids minted from different
 * leases never meet. Without a lease every allocator draws from the whole
 * id space, lowest first.
 */
import { type Document, MAX_REVISION_ID } from "../model/document";
import {
  IDENTITY_SPACES,
  packageIdentityKeyIndex,
  packageParagraphIdIndex,
} from "./ids";

export const ID_LEASE_KINDS: Readonly<{
  readonly PARAGRAPH: "paragraph";
  readonly REVISION: "revision";
  readonly CONTROL: "control";
  readonly COMMENT: "comment";
}> = Object.freeze({
  /** `w14:paraId`. */
  PARAGRAPH: "paragraph",
  /** `w:id` of tracked changes and tracked property changes. */
  REVISION: "revision",
  /** `w:sdtPr/w:id`. */
  CONTROL: "control",
  /** `w:id` of comments and their anchors. */
  COMMENT: "comment",
} as const);

export type IdLeaseKind = (typeof ID_LEASE_KINDS)[keyof typeof ID_LEASE_KINDS];

export const ID_LEASE_KIND_LIST: readonly IdLeaseKind[] = Object.freeze(
  Object.values(ID_LEASE_KINDS),
);

/**
 * One past the largest id of each kind: a paragraph id stays below
 * 0x80000000, and `w:id` is a signed 32-bit integer.
 */
export const ID_KIND_LIMITS: Readonly<Record<IdLeaseKind, number>> =
  Object.freeze({
    paragraph: 0x8000_0000,
    revision: MAX_REVISION_ID + 1,
    control: MAX_REVISION_ID + 1,
    comment: MAX_REVISION_ID + 1,
  });

/** Ids from `start` up to, not including, `end`. */
export type IdRange = { readonly start: number; readonly end: number };

/** One leased range of one id kind. */
export type IdBlock = IdRange & { readonly kind: IdLeaseKind };

/** The ids one author may still mint: disjoint blocks, ascending per kind. */
export type IdLease = { readonly blocks: readonly IdBlock[] };

/** A paragraph id's spelling for its number. */
export const formatParaId = (id: number): string =>
  id.toString(16).padStart(8, "0").toUpperCase();

const parseParaId = (id: string): number | undefined => {
  const value = Number.parseInt(id, 16);
  return Number.isSafeInteger(value) ? value : undefined;
};

const blocksOf = (lease: IdLease, kind: IdLeaseKind): IdBlock[] =>
  lease.blocks
    .filter((block) => block.kind === kind)
    .toSorted((left, right) => left.start - right.start);

/**
 * The ranges an allocator walks for `kind`, ascending, none below `from`:
 * the lease's blocks, or, without a lease, every id of the kind.
 */
export const idRanges = (
  lease: IdLease | undefined,
  kind: IdLeaseKind,
  from = 1,
): readonly IdRange[] => {
  const floor = Math.max(1, from);
  if (lease === undefined)
    return floor < ID_KIND_LIMITS[kind]
      ? [{ start: floor, end: ID_KIND_LIMITS[kind] }]
      : [];
  const ranges: IdRange[] = [];
  for (const block of blocksOf(lease, kind)) {
    const start = Math.max(block.start, floor);
    if (start < block.end) ranges.push({ start, end: block.end });
  }
  return ranges;
};

/** Whether `lease` holds `id` of `kind`. */
export const idLeaseHolds = (
  lease: IdLease,
  kind: IdLeaseKind,
  id: number,
): boolean =>
  lease.blocks.some(
    (block) => block.kind === kind && block.start <= id && id < block.end,
  );

/** How many ids of `kind` the lease still holds. */
export const idLeaseRemaining = (lease: IdLease, kind: IdLeaseKind): number =>
  blocksOf(lease, kind).reduce(
    (sum, block) => sum + Math.max(0, block.end - block.start),
    0,
  );

/** The ids of each kind a document holds anywhere in its package. */
export type LeasedIdCensus = Readonly<Record<IdLeaseKind, ReadonlySet<number>>>;

/** Every paragraph, revision, content-control and comment id of a document. */
export const leasedIdCensus = (document: Document): LeasedIdCensus => {
  const paragraph = new Set<number>();
  for (const key of packageParagraphIdIndex(document.package).counts().keys()) {
    const id = parseParaId(key);
    if (id !== undefined) paragraph.add(id);
  }
  const revision = new Set<number>();
  const control = new Set<number>();
  for (const key of packageIdentityKeyIndex(document.package).counts().keys()) {
    const separator = key.lastIndexOf(":");
    const space = key.slice(0, separator);
    const id = Number(key.slice(separator + 1));
    if (!Number.isSafeInteger(id)) continue;
    if (space === IDENTITY_SPACES.REVISION) revision.add(id);
    else if (space === IDENTITY_SPACES.CONTROL) control.add(id);
  }
  const comment = new Set<number>();
  for (const record of document.package.document.comments ?? [])
    comment.add(record.id);
  return { paragraph, revision, control, comment };
};

/** One id of one kind. */
export type LeasedId = { readonly kind: IdLeaseKind; readonly id: number };

/** The ids `after` holds that `before` does not, by kind and then by id. */
export const introducedIds = (
  before: LeasedIdCensus,
  after: LeasedIdCensus,
): LeasedId[] => {
  const introduced: LeasedId[] = [];
  for (const kind of ID_LEASE_KIND_LIST) {
    const ids: number[] = [];
    for (const id of after[kind]) if (!before[kind].has(id)) ids.push(id);
    ids.sort((left, right) => left - right);
    for (const id of ids) introduced.push({ kind, id });
  }
  return introduced;
};

/**
 * `lease` without any id at or below the highest of `ids` it holds, per kind:
 * an author mints upward through its lease and never mints an id twice, so an
 * id it minted and later removed cannot meet another author's restoration of it.
 */
export const trimIdLease = (
  lease: IdLease,
  ids: readonly LeasedId[],
): IdLease => {
  const floors = new Map<IdLeaseKind, number>();
  for (const { kind, id } of ids) {
    if (!idLeaseHolds(lease, kind, id)) continue;
    floors.set(kind, Math.max(floors.get(kind) ?? 0, id));
  }
  if (floors.size === 0) return lease;
  const blocks: IdBlock[] = [];
  for (const block of lease.blocks) {
    const floor = floors.get(block.kind);
    const start =
      floor === undefined ? block.start : Math.max(block.start, floor + 1);
    if (start < block.end) blocks.push({ ...block, start });
  }
  return { blocks };
};

type FreshLeasedIdsOptions = {
  lease: IdLease | undefined;
  kind: IdLeaseKind;
  count: number;
  /** Whether an id is already in use and cannot be minted. */
  taken: (id: number) => boolean;
  /** The lowest candidate, 1 by default. */
  from?: number;
};

/**
 * Up to `count` ids of `kind` that `taken` refuses, lowest first, from the
 * lease, or from the whole space without one. Fewer come back only when the
 * lease or the space has no more: the caller refuses the edit, never wraps.
 */
export const freshLeasedIds = ({
  lease,
  kind,
  count,
  taken,
  from = 1,
}: FreshLeasedIdsOptions): number[] => {
  const ids: number[] = [];
  for (const { start, end } of idRanges(lease, kind, from)) {
    for (let id = start; id < end && ids.length < count; id += 1) {
      if (!taken(id)) ids.push(id);
    }
    if (ids.length >= count) break;
  }
  return ids;
};
