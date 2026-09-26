/**
 * Carry many positions through the step maps of one or more transactions in
 * one pass.
 *
 * `Mapping.map` walks every step map for every position, so a plugin that maps
 * each paragraph, or each step's range, through a transaction costs
 * O(positions x steps). That is quadratic on the one-transaction batches that
 * paste, replace-all, accept-all and document compare dispatch.
 *
 * A step map only decides the positions its changed ranges cover; everything
 * after a range moves by that range's size difference, and everything before
 * it stays put. So the positions are kept sorted, and each map is applied
 * once: the covered positions are mapped exactly, by the map itself, and the
 * rest are shifted together. Mapping is monotone for a fixed `assoc`, so the
 * order survives every map and one sorted lane per `assoc` suffices.
 *
 * The answer for each query is what the per-step walk returns:
 * `mapping.slice(from).mapResult(pos, assoc)` through the transaction the
 * query enters in, then `mapResult` through each later one, with
 * `deletedAcross` true when any of those reported it.
 */

import type { Mapping, StepMap } from "prosemirror-transform";

export type PositionQuery = {
  pos: number;
  assoc: -1 | 1;
  /**
   * Index of the first step map to apply, counting across every mapping
   * passed in order: the position is in the document that step read.
   */
  from: number;
};

export type SweptPosition = {
  pos: number;
  deletedAcross: boolean;
};

/** Positions kept per chunk before it splits; small enough to splice cheaply. */
const CHUNK_SIZE = 128;

type Chunk = {
  /** Added to every entry of `values`; how a whole chunk moves in O(1). */
  shift: number;
  values: number[];
  ids: number[];
};

type Cursor = { chunk: number; index: number };

const cursorBefore = (left: Cursor, right: Cursor): boolean =>
  left.chunk < right.chunk || (left.chunk === right.chunk && left.index < right.index);

/** Positions sharing one `assoc`, sorted by their current value. */
class Lane {
  private readonly chunks: Chunk[] = [];

  private readonly assoc: -1 | 1;

  constructor(assoc: -1 | 1) {
    this.assoc = assoc;
  }

  get isEmpty(): boolean {
    return this.chunks.length === 0;
  }

  /** Replace an empty lane's contents with `entries`, sorted by value. */
  load(entries: readonly (readonly [value: number, id: number])[]): void {
    for (let start = 0; start < entries.length; start += CHUNK_SIZE) {
      const slice = entries.slice(start, start + CHUNK_SIZE);
      this.chunks.push({
        shift: 0,
        values: slice.map(([value]) => value),
        ids: slice.map(([, id]) => id),
      });
    }
  }

  insert(value: number, id: number): void {
    const last = this.chunks.at(-1);
    if (!last) {
      this.chunks.push({ shift: 0, values: [value], ids: [id] });
      return;
    }
    const cursor = this.bound(value, true);
    const chunkIndex = Math.min(cursor.chunk, this.chunks.length - 1);
    // SAFETY: chunkIndex is clamped into the non-empty chunk list.
    const chunk = this.chunks[chunkIndex]!;
    const index = cursor.chunk === chunkIndex ? cursor.index : chunk.values.length;
    chunk.values.splice(index, 0, value - chunk.shift);
    chunk.ids.splice(index, 0, id);
    if (chunk.values.length > 2 * CHUNK_SIZE) {
      this.chunks.splice(chunkIndex + 1, 0, {
        shift: chunk.shift,
        values: chunk.values.splice(CHUNK_SIZE),
        ids: chunk.ids.splice(CHUNK_SIZE),
      });
    }
  }

  /**
   * The first entry whose value is at least `value` (`strict`: greater than
   * it), or the end of the lane.
   */
  private bound(value: number, strict: boolean): Cursor {
    const passes = (candidate: number): boolean =>
      strict ? candidate > value : candidate >= value;
    let low = 0;
    let high = this.chunks.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      // SAFETY: middle < chunks.length, and a chunk is never empty.
      const chunk = this.chunks[middle]!;
      if (passes(chunk.values.at(-1)! + chunk.shift)) {
        high = middle;
      } else {
        low = middle + 1;
      }
    }
    const chunk = this.chunks[low];
    if (!chunk) {
      return { chunk: this.chunks.length, index: 0 };
    }
    let first = 0;
    let last = chunk.values.length;
    while (first < last) {
      const middle = (first + last) >>> 1;
      // SAFETY: middle < values.length.
      if (passes(chunk.values[middle]! + chunk.shift)) {
        last = middle;
      } else {
        first = middle + 1;
      }
    }
    return { chunk: low, index: first };
  }

  /** Visit every entry from `start` up to, not including, `end`. */
  private forEachBetween(
    start: Cursor,
    end: Cursor,
    visit: (chunk: Chunk, index: number) => void,
  ): void {
    for (let chunkIndex = start.chunk; chunkIndex <= end.chunk; chunkIndex++) {
      const chunk = this.chunks[chunkIndex];
      if (!chunk) {
        return;
      }
      const first = chunkIndex === start.chunk ? start.index : 0;
      const last = chunkIndex === end.chunk ? end.index : chunk.values.length;
      for (let index = first; index < last; index++) {
        visit(chunk, index);
      }
    }
  }

  /** Add `delta` to every entry from `start` to the end of the lane. */
  private shiftFrom(start: Cursor, end: Cursor | null, delta: number): void {
    if (delta === 0) {
      return;
    }
    const stop = end ?? { chunk: this.chunks.length, index: 0 };
    for (let chunkIndex = start.chunk; chunkIndex <= stop.chunk; chunkIndex++) {
      const chunk = this.chunks[chunkIndex];
      if (!chunk) {
        return;
      }
      const first = chunkIndex === start.chunk ? start.index : 0;
      const last = chunkIndex === stop.chunk ? stop.index : chunk.values.length;
      if (first === 0 && last === chunk.values.length) {
        chunk.shift += delta;
        continue;
      }
      for (let index = first; index < last; index++) {
        // SAFETY: index < values.length.
        chunk.values[index]! += delta;
      }
    }
  }

  apply(map: StepMap, deletedAcross: boolean[]): void {
    if (this.chunks.length === 0) {
      return;
    }
    // `_map` in prosemirror-transform decides a position with the first range
    // whose [oldStart, oldEnd] holds it and shifts it by the size difference
    // of every range that ended before it. The bounds are all taken before
    // anything moves: a mapped value may land inside an earlier range.
    const ranges: { start: Cursor; end: Cursor; delta: number }[] = [];
    let previousEnd: Cursor = { chunk: 0, index: 0 };
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
    map.forEach((oldStart, oldEnd, newStart, newEnd) => {
      const lower = this.bound(oldStart, false);
      const start = cursorBefore(lower, previousEnd) ? previousEnd : lower;
      const end = this.bound(oldEnd, true);
      ranges.push({ start, end, delta: newEnd - newStart - (oldEnd - oldStart) });
      previousEnd = end;
    });
    let cursor: Cursor = { chunk: 0, index: 0 };
    let delta = 0;
    for (const range of ranges) {
      this.shiftFrom(cursor, range.start, delta);
      this.forEachBetween(range.start, range.end, (chunk, index) => {
        // SAFETY: forEachBetween only visits indices inside the chunk.
        const result = map.mapResult(chunk.values[index]! + chunk.shift, this.assoc);
        chunk.values[index] = result.pos - chunk.shift;
        if (result.deletedAcross) {
          // SAFETY: as above.
          deletedAcross[chunk.ids[index]!] = true;
        }
      });
      delta += range.delta;
      cursor = range.end;
    }
    this.shiftFrom(cursor, null, delta);
  }

  collect(positions: number[]): void {
    for (const chunk of this.chunks) {
      for (let index = 0; index < chunk.values.length; index++) {
        // SAFETY: ids and values have equal lengths.
        positions[chunk.ids[index]!] = chunk.values[index]! + chunk.shift;
      }
    }
  }
}

const hasMirrors = (mapping: Mapping): boolean => {
  for (let index = mapping.from; index < mapping.to; index++) {
    if (mapping.getMirror(index) !== undefined) {
      return true;
    }
  }
  return false;
};

/**
 * The per-step walk, kept for mappings that carry mirror pairs (collaborative
 * rebases): a mirrored map recovers a deleted position from its twin, which
 * only a walk through every map can see.
 */
const mapPerStep = (
  mappings: readonly Mapping[],
  queries: readonly PositionQuery[],
): SweptPosition[] =>
  queries.map(({ pos, assoc, from }) => {
    let mapped = pos;
    let deletedAcross = false;
    let offset = 0;
    for (const mapping of mappings) {
      const length = mapping.to - mapping.from;
      if (from < offset + length) {
        const result = mapping
          .slice(mapping.from + Math.max(0, from - offset), mapping.to)
          .mapResult(mapped, assoc);
        mapped = result.pos;
        deletedAcross ||= result.deletedAcross;
      }
      offset += length;
    }
    return { pos: mapped, deletedAcross };
  });

/**
 * Map every query through `mappings` (one per transaction, in order), each
 * entering at its own step. O((positions + steps) x log positions) plus the
 * positions a step's changed ranges actually cover, rather than
 * O(positions x steps).
 */
export const sweepPositions = (
  mappings: readonly Mapping[],
  queries: readonly PositionQuery[],
): SweptPosition[] => {
  if (queries.length === 0) {
    return [];
  }
  if (mappings.some(hasMirrors)) {
    return mapPerStep(mappings, queries);
  }
  const maps: StepMap[] = mappings.flatMap((mapping) =>
    mapping.maps.slice(mapping.from, mapping.to),
  );
  const order = queries
    .map((query, id) => ({ id, from: Math.min(Math.max(query.from, 0), maps.length) }))
    .sort((left, right) => left.from - right.from || left.id - right.id);
  const backward = new Lane(-1);
  const forward = new Lane(1);
  const deletedAcross = queries.map(() => false);

  let next = 0;
  const enter = (step: number): void => {
    const entering = new Map<Lane, [number, number][]>([
      [backward, []],
      [forward, []],
    ]);
    for (; next < order.length && order[next]!.from === step; next++) {
      // SAFETY: next < order.length, and every order id indexes queries.
      const { id } = order[next]!;
      const query = queries[id]!;
      entering.get(query.assoc < 0 ? backward : forward)!.push([query.pos, id]);
    }
    for (const [lane, entries] of entering) {
      if (lane.isEmpty) {
        lane.load(entries.sort((left, right) => left[0] - right[0]));
        continue;
      }
      for (const [value, id] of entries) {
        lane.insert(value, id);
      }
    }
  };

  for (let step = 0; step < maps.length; step++) {
    if (next < order.length && order[next]!.from === step) {
      enter(step);
    }
    // SAFETY: step < maps.length.
    const map = maps[step]!;
    backward.apply(map, deletedAcross);
    forward.apply(map, deletedAcross);
  }
  if (next < order.length) {
    enter(maps.length);
  }

  const positions: number[] = queries.map(({ pos }) => pos);
  backward.collect(positions);
  forward.collect(positions);
  return positions.map((pos, id) => ({ pos, deletedAcross: deletedAcross[id] ?? false }));
};
