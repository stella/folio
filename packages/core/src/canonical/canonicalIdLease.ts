import type { Document } from "@stll/docx-core/model";
import {
  type IdBlock,
  type IdLease,
  idLeaseHolds,
  type IdLeaseKind,
  introducedIds,
  type LeasedId,
  type LeasedIdCensus,
  leasedIdCensus,
  trimIdLease,
} from "@stll/docx-core/ops";

/** Union of the blocks, per kind, overlapping and adjacent ones merged. */
const mergeBlocks = (blocks: readonly IdBlock[]): IdBlock[] => {
  const merged: IdBlock[] = [];
  const sorted = blocks.toSorted(
    (left, right) =>
      left.kind.localeCompare(right.kind) || left.start - right.start,
  );
  for (const block of sorted) {
    const last = merged.at(-1);
    if (
      last !== undefined &&
      last.kind === block.kind &&
      block.start <= last.end
    )
      merged[merged.length - 1] = {
        ...last,
        end: Math.max(last.end, block.end),
      };
    else merged.push(block);
  }
  return merged;
};

/**
 * What a session holding an id lease may mint: the lease above every id it
 * minted, and the ids its documents have held, which an edit may restore.
 */
export class CanonicalIdLease {
  private current: IdLease;
  private census: LeasedIdCensus;
  private readonly seen: Record<IdLeaseKind, Set<number>>;
  private readonly floors = new Map<IdLeaseKind, number>();

  constructor(lease: IdLease, document: Document) {
    this.current = { blocks: mergeBlocks(lease.blocks) };
    this.census = leasedIdCensus(document);
    this.seen = {
      paragraph: new Set(this.census.paragraph),
      revision: new Set(this.census.revision),
      control: new Set(this.census.control),
      comment: new Set(this.census.comment),
    };
  }

  get lease(): IdLease {
    return this.current;
  }

  /**
   * The ids `after` introduces, or the first one this session may not mint:
   * every new id lies in the lease or restores one its documents held.
   */
  judge(
    after: Document,
  ):
    | { type: "allowed"; census: LeasedIdCensus; ids: readonly LeasedId[] }
    | { type: "outside"; id: LeasedId } {
    const census = leasedIdCensus(after);
    const ids = introducedIds(this.census, census);
    for (const id of ids) {
      if (this.seen[id.kind].has(id.id)) continue;
      if (idLeaseHolds(this.current, id.kind, id.id)) continue;
      return { type: "outside", id };
    }
    return { type: "allowed", census, ids };
  }

  /** Record a published commit: its ids are held, the lease moves past them. */
  commit(census: LeasedIdCensus, ids: readonly LeasedId[]): void {
    this.census = census;
    for (const { kind, id } of ids) {
      this.seen[kind].add(id);
      if (idLeaseHolds(this.current, kind, id))
        this.floors.set(kind, Math.max(this.floors.get(kind) ?? 0, id));
    }
    this.current = trimIdLease(this.current, ids);
  }

  /** Add granted blocks, never below an id this session already minted. */
  extend(lease: IdLease): void {
    const blocks: IdBlock[] = [];
    for (const block of mergeBlocks([
      ...this.current.blocks,
      ...lease.blocks,
    ])) {
      const floor = this.floors.get(block.kind);
      const start =
        floor === undefined ? block.start : Math.max(block.start, floor + 1);
      if (start < block.end) blocks.push({ ...block, start });
    }
    this.current = { blocks };
  }
}
