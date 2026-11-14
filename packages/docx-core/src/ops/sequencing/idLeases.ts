import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import type { Document } from "../../model/document";
import {
  ID_KIND_LIMITS,
  ID_LEASE_KIND_LIST,
  type IdBlock,
  type IdLease,
  type IdLeaseKind,
  introducedIds,
  type LeasedId,
  type LeasedIdCensus,
  leasedIdCensus,
} from "../idLease";

/** How many ids one grant or renewal leases, per kind. */
export const DEFAULT_ID_BLOCK_SIZES: Readonly<Record<IdLeaseKind, number>> =
  Object.freeze({
    paragraph: 4096,
    revision: 4096,
    control: 1024,
    comment: 1024,
  });

/** One block leased to one actor. */
export type IdLeaseGrant = { readonly actor: string; readonly block: IdBlock };

/** The highest id of a kind one actor's sequenced batches introduced. */
export type IdLeaseMinted = {
  readonly actor: string;
  readonly kind: IdLeaseKind;
  readonly highest: number;
};

/**
 * A room's lease table, part of its sequenced state: kept with the journal
 * and given back to a restarted sequencer, so each actor keeps its blocks.
 */
export type IdLeaseState = {
  readonly blockSizes: Readonly<Record<IdLeaseKind, number>>;
  /** Every block granted, in grant order. */
  readonly grants: readonly IdLeaseGrant[];
  readonly minted: readonly IdLeaseMinted[];
};

export const ID_LEASE_REFUSAL_REASONS: Readonly<{
  readonly EXHAUSTED: "exhausted";
  readonly NOT_DUE: "notDue";
  readonly INVALID_STATE: "invalidState";
  readonly UNAVAILABLE: "unavailable";
}> = Object.freeze({
  /** No free block of the kind is left below the kind's id limit. */
  EXHAUSTED: "exhausted",
  /** The actor still holds more than a block of unused ids of the kind. */
  NOT_DUE: "notDue",
  /** A restored lease table is malformed or its blocks overlap. */
  INVALID_STATE: "invalidState",
  /**
   * No authority answered: the transport closed or stopped following the
   * room before an answer arrived. Asking again on a working transport may
   * succeed.
   */
  UNAVAILABLE: "unavailable",
} as const);

export type IdLeaseRefusalReason =
  (typeof ID_LEASE_REFUSAL_REASONS)[keyof typeof ID_LEASE_REFUSAL_REASONS];

const IdLeaseRefusalBase: TaggedErrorClass<"IdLeaseRefusal"> =
  TaggedError("IdLeaseRefusal");

export class IdLeaseRefusal extends IdLeaseRefusalBase<{
  message: string;
  reason: IdLeaseRefusalReason;
  kind?: IdLeaseKind;
}> {}

/** Kinds that draw from one numeric space take blocks that never overlap. */
const SPACES: Readonly<Record<IdLeaseKind, "paragraph" | "annotation">> = {
  paragraph: "paragraph",
  revision: "annotation",
  control: "annotation",
  comment: "annotation",
};
type Space = (typeof SPACES)[IdLeaseKind];

/** An id outside the submitter's lease, as the sequencer refuses it. */
export type IdLeaseViolation = LeasedId & { readonly actor: string };

/**
 * The authority's lease table: grants disjoint blocks, renews them, and
 * judges which ids a batch may introduce.
 */
export type IdLeaseTable = {
  /** The actor's remaining lease, granting one block per kind the first time. */
  grant: (actor: string) => Result<IdLease, IdLeaseRefusal>;
  /** One more block of `kind`, while the actor holds at most a block of unused ids. */
  renew: (actor: string, kind: IdLeaseKind) => Result<IdBlock, IdLeaseRefusal>;
  /**
   * The first id `after` introduces that `actor` may not: an id is accepted
   * when the room held it before (a restoration), when it lies in the actor's
   * lease above what the actor already minted, or when it lies at or below
   * what its lessee already minted (a restoration nobody can mint again).
   */
  check: (
    actor: string,
    after: Document,
  ) => Result<readonly LeasedId[], IdLeaseViolation>;
  /** Record a sequenced batch's introduced ids and its document. */
  record: (
    actor: string,
    after: Document,
    introduced: readonly LeasedId[],
  ) => void;
  readonly state: IdLeaseState;
};

const invalidState = (message: string) =>
  new IdLeaseRefusal({
    reason: ID_LEASE_REFUSAL_REASONS.INVALID_STATE,
    message,
  });

const validBlockSizes = (sizes: IdLeaseState["blockSizes"]): boolean =>
  ID_LEASE_KIND_LIST.every((kind) => {
    const size = sizes[kind];
    return (
      Number.isSafeInteger(size) && size > 0 && size < ID_KIND_LIMITS[kind]
    );
  });

/** A lease table over `document`, restoring `state` when given. */
export const createIdLeaseTable = (
  document: Document,
  state?: IdLeaseState,
): Result<IdLeaseTable, IdLeaseRefusal> => {
  const blockSizes = state?.blockSizes ?? DEFAULT_ID_BLOCK_SIZES;
  if (!validBlockSizes(blockSizes))
    return Result.err(invalidState("Lease block sizes must be positive."));
  const grants: IdLeaseGrant[] = [];
  const minted = new Map<string, Map<IdLeaseKind, number>>();
  const setMinted = (actor: string, kind: IdLeaseKind, highest: number) => {
    const kinds = minted.get(actor) ?? new Map<IdLeaseKind, number>();
    kinds.set(kind, highest);
    minted.set(actor, kinds);
  };
  const cursors: Record<Space, number> = { paragraph: 1, annotation: 1 };
  for (const grant of state?.grants ?? []) {
    const { kind, start, end } = grant.block;
    if (
      !ID_LEASE_KIND_LIST.includes(kind) ||
      !Number.isSafeInteger(start) ||
      !Number.isSafeInteger(end) ||
      start < 1 ||
      end <= start ||
      end > ID_KIND_LIMITS[kind] ||
      grant.actor.length === 0
    )
      return Result.err(invalidState("A restored lease block is malformed."));
    const space = SPACES[kind];
    if (
      grants.some(
        ({ block }) =>
          SPACES[block.kind] === space &&
          block.start < end &&
          start < block.end,
      )
    )
      return Result.err(invalidState("Restored lease blocks overlap."));
    grants.push(grant);
    cursors[space] = Math.max(cursors[space], end);
  }
  for (const { actor, kind, highest } of state?.minted ?? []) {
    if (!ID_LEASE_KIND_LIST.includes(kind) || !Number.isSafeInteger(highest))
      return Result.err(invalidState("A restored minted id is malformed."));
    setMinted(actor, kind, highest);
  }

  let census: LeasedIdCensus = leasedIdCensus(document);
  // Ids the room has held, by kind: a batch may restore them.
  const seen: Record<IdLeaseKind, Set<number>> = {
    paragraph: new Set(census.paragraph),
    revision: new Set(census.revision),
    control: new Set(census.control),
    comment: new Set(census.comment),
  };

  const lesseeOf = (kind: IdLeaseKind, id: number): string | undefined =>
    grants.find(
      ({ block }) => block.kind === kind && block.start <= id && id < block.end,
    )?.actor;
  const highestOf = (actor: string, kind: IdLeaseKind): number =>
    minted.get(actor)?.get(kind) ?? 0;

  /** The actor's blocks above what it already minted. */
  const remainingLease = (actor: string): IdLease => {
    const blocks: IdBlock[] = [];
    for (const { actor: holder, block } of grants) {
      if (holder !== actor) continue;
      const start = Math.max(block.start, highestOf(actor, block.kind) + 1);
      if (start < block.end) blocks.push({ ...block, start });
    }
    return { blocks };
  };

  /** The lowest block of `kind` above the space's cursor holding no id the room held. */
  const nextBlock = (kind: IdLeaseKind): IdBlock | undefined => {
    const space = SPACES[kind];
    const size = blockSizes[kind];
    const held: number[] = [];
    for (const other of ID_LEASE_KIND_LIST) {
      if (SPACES[other] !== space) continue;
      for (const id of seen[other]) if (id >= cursors[space]) held.push(id);
    }
    held.sort((left, right) => left - right);
    let start = cursors[space];
    for (const id of held) {
      if (id >= start + size) break;
      if (id >= start) start = id + 1;
    }
    return start + size > ID_KIND_LIMITS[kind]
      ? undefined
      : { kind, start, end: start + size };
  };

  const take = (actor: string, block: IdBlock): void => {
    grants.push({ actor, block });
    cursors[SPACES[block.kind]] = block.end;
  };

  const exhausted = (kind: IdLeaseKind) =>
    new IdLeaseRefusal({
      reason: ID_LEASE_REFUSAL_REASONS.EXHAUSTED,
      kind,
      message: `No ${kind} id block is left below the ${kind} id limit.`,
    });

  const grant = (actor: string): Result<IdLease, IdLeaseRefusal> => {
    if (grants.some((entry) => entry.actor === actor))
      return Result.ok(remainingLease(actor));
    const granted = grants.length;
    const saved = { ...cursors };
    const blocks: IdBlock[] = [];
    for (const kind of ID_LEASE_KIND_LIST) {
      const block = nextBlock(kind);
      if (block === undefined) {
        // All or nothing: the blocks taken so far go back.
        grants.length = granted;
        Object.assign(cursors, saved);
        return Result.err(exhausted(kind));
      }
      take(actor, block);
      blocks.push(block);
    }
    return Result.ok({ blocks });
  };

  const renew = (
    actor: string,
    kind: IdLeaseKind,
  ): Result<IdBlock, IdLeaseRefusal> => {
    const remaining = remainingLease(actor)
      .blocks.filter((block) => block.kind === kind)
      .reduce((sum, block) => sum + (block.end - block.start), 0);
    if (remaining > blockSizes[kind])
      return Result.err(
        new IdLeaseRefusal({
          reason: ID_LEASE_REFUSAL_REASONS.NOT_DUE,
          kind,
          message: `The actor still holds ${remaining} unused ${kind} ids.`,
        }),
      );
    const block = nextBlock(kind);
    if (block === undefined) return Result.err(exhausted(kind));
    take(actor, block);
    return Result.ok(block);
  };

  const check = (
    actor: string,
    after: Document,
  ): Result<readonly LeasedId[], IdLeaseViolation> => {
    const introduced = introducedIds(census, leasedIdCensus(after));
    for (const entry of introduced) {
      if (seen[entry.kind].has(entry.id)) continue;
      const lessee = lesseeOf(entry.kind, entry.id);
      if (lessee === actor && entry.id > highestOf(actor, entry.kind)) continue;
      if (lessee !== undefined && entry.id <= highestOf(lessee, entry.kind))
        continue;
      return Result.err({ ...entry, actor });
    }
    return Result.ok(introduced);
  };

  const record = (
    actor: string,
    after: Document,
    introduced: readonly LeasedId[],
  ): void => {
    census = leasedIdCensus(after);
    for (const { kind, id } of introduced) {
      seen[kind].add(id);
      if (lesseeOf(kind, id) === actor && id > highestOf(actor, kind))
        setMinted(actor, kind, id);
    }
  };

  return Result.ok({
    grant,
    renew,
    check,
    record,
    get state(): IdLeaseState {
      return {
        blockSizes: { ...blockSizes },
        grants: grants.slice(),
        minted: [...minted].flatMap(([actor, kinds]) =>
          [...kinds].map(([kind, highest]) => ({ actor, kind, highest })),
        ),
      };
    },
  });
};
