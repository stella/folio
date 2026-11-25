/**
 * The source gaps of one paragraph's logical text, by offset.
 *
 * Each offset of the text (`0` to `text.length`) has one or more gaps: an
 * editor position, relative to the paragraph, that stands for it, the
 * zero-width source ordinal that position names, and the ordinals its seam
 * covers (`first` to `last`). Gaps at one offset keep the order they were
 * added in.
 *
 * Nearly every offset has exactly one gap, so the first gap of each offset is
 * held flat, four integers per offset in one `Int32Array`, and only the rare
 * further gaps (zero-width markers at one offset that several editor positions
 * reach) go to a side map. A paragraph's gaps then cost about sixteen bytes
 * per character, not two arrays and two objects per character; a canonical
 * session keeps them for every paragraph of the story.
 */

/** An editor position (relative to its paragraph) and the source ordinal it names. */
export type CanonicalInlineGap = { position: number; zeroWidthBefore: number };
/** The source ordinals one editor position stands for: from `first` to `last`. */
export type CanonicalSeamExtent = { first: number; last: number };
/** One gap with its seam. */
export type CanonicalGapEntry = CanonicalInlineGap & CanonicalSeamExtent;

const SLOTS = 4;
const POSITION = 0;
const ZERO_WIDTH_BEFORE = 1;
const FIRST = 2;
const LAST = 3;
/** No gap at this offset yet. Positions are never negative. */
const EMPTY = -1;

export class CanonicalInlineGaps {
  /** The first gap of each offset: position, zeroWidthBefore, first, last. */
  private readonly flat: Int32Array;
  /** The second and later gaps of an offset, in the order they were added. */
  private readonly more = new Map<number, CanonicalGapEntry[]>();

  /** Room for `offsets` offsets (`text.length + 1`), none holding a gap yet. */
  constructor(offsets: number) {
    this.flat = new Int32Array(offsets * SLOTS).fill(EMPTY);
  }

  /** How many offsets the table holds: the text's length plus one. */
  get offsets(): number {
    return this.flat.length / SLOTS;
  }

  /**
   * Record a gap at `offset`. A gap at an editor position the offset already
   * has widens that gap's seam to the ordinal, and the gap then names the
   * ordinal of the two that is nearer to `preferred()` (asked only then); a
   * tie keeps the earlier one. Returns false when the offset is outside the
   * text.
   */
  add(
    offset: number,
    position: number,
    zeroWidthBefore: number,
    preferred: () => number,
  ): boolean {
    if (!Number.isInteger(offset) || offset < 0 || offset >= this.offsets)
      return false;
    const at = offset * SLOTS;
    const flat = this.flat;
    if (flat[at + POSITION] === EMPTY) {
      flat[at + POSITION] = position;
      flat[at + ZERO_WIDTH_BEFORE] = zeroWidthBefore;
      flat[at + FIRST] = zeroWidthBefore;
      flat[at + LAST] = zeroWidthBefore;
      return true;
    }
    if (flat[at + POSITION] === position) {
      flat[at + FIRST] = Math.min(flat[at + FIRST] ?? 0, zeroWidthBefore);
      flat[at + LAST] = Math.max(flat[at + LAST] ?? 0, zeroWidthBefore);
      const kept = flat[at + ZERO_WIDTH_BEFORE] ?? 0;
      const target = preferred();
      if (Math.abs(zeroWidthBefore - target) < Math.abs(kept - target))
        flat[at + ZERO_WIDTH_BEFORE] = zeroWidthBefore;
      return true;
    }
    const more = this.more.get(offset);
    const known = more?.find((entry) => entry.position === position);
    if (known === undefined) {
      const entry = {
        position,
        zeroWidthBefore,
        first: zeroWidthBefore,
        last: zeroWidthBefore,
      };
      if (more === undefined) this.more.set(offset, [entry]);
      else more.push(entry);
      return true;
    }
    known.first = Math.min(known.first, zeroWidthBefore);
    known.last = Math.max(known.last, zeroWidthBefore);
    const target = preferred();
    if (
      Math.abs(zeroWidthBefore - target) <
      Math.abs(known.zeroWidthBefore - target)
    )
      known.zeroWidthBefore = zeroWidthBefore;
    return true;
  }

  /** Whether some offset has no gap. */
  hasEmptyOffset(): boolean {
    for (let offset = 0; offset < this.offsets; offset += 1)
      if (this.flat[offset * SLOTS + POSITION] === EMPTY) return true;
    return false;
  }

  /** The gaps at `offset` with their seams, in order; undefined outside the text. */
  at(offset: number): readonly CanonicalGapEntry[] | undefined {
    if (!Number.isInteger(offset) || offset < 0 || offset >= this.offsets)
      return undefined;
    const at = offset * SLOTS;
    const flat = this.flat;
    const position = flat[at + POSITION] ?? EMPTY;
    if (position === EMPTY) return [];
    const head: CanonicalGapEntry = {
      position,
      zeroWidthBefore: flat[at + ZERO_WIDTH_BEFORE] ?? 0,
      first: flat[at + FIRST] ?? 0,
      last: flat[at + LAST] ?? 0,
    };
    const more = this.more.get(offset);
    return more === undefined
      ? [head]
      : [
          head,
          ...more.map((entry) => ({
            position: entry.position,
            zeroWidthBefore: entry.zeroWidthBefore,
            first: entry.first,
            last: entry.last,
          })),
        ];
  }

  /** The first offset with a gap at editor `position`, or -1. */
  offsetOf(position: number): number {
    const flat = this.flat;
    const hasMore = this.more.size > 0;
    for (let offset = 0; offset < this.offsets; offset += 1) {
      if (flat[offset * SLOTS + POSITION] === position) return offset;
      if (
        hasMore &&
        this.more.get(offset)?.some((entry) => entry.position === position)
      )
        return offset;
    }
    return -1;
  }

  /** Every offset's gaps, without their seams. */
  toBoundaries(): CanonicalInlineGap[][] {
    return Array.from({ length: this.offsets }, (_, offset) =>
      (this.at(offset) ?? []).map(({ position, zeroWidthBefore }) => ({
        position,
        zeroWidthBefore,
      })),
    );
  }

  /** Every offset's seams, parallel to {@link toBoundaries}. */
  toSeams(): CanonicalSeamExtent[][] {
    return Array.from({ length: this.offsets }, (_, offset) =>
      (this.at(offset) ?? []).map(({ first, last }) => ({ first, last })),
    );
  }
}
