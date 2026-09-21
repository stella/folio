/**
 * A comment or tracked-move range that spans no content, carried across the
 * editor projection as one zero-width atom.
 *
 * A reviewer who comments at an insertion point rather than over a selection
 * gets `w:commentRangeStart` immediately followed by `w:commentRangeEnd`, and
 * then the `w:commentReference` run. The editor spells a comment as a `comment`
 * mark over the inline content its range covers, and a range over nothing has
 * no leaf to mark, so the projection dropped both markers. The comment still
 * listed — the reference survives as its own atom — but its anchor position was
 * gone, and the save wrote a reference with no range at all. A tracked move
 * whose range spans nothing loses its `w:name` the same way, and with it the
 * only thing binding the move's source to its destination.
 *
 * The pair is one thing rather than two, so it is one node. Two adjacent
 * boundary nodes, which is what the bookmark precedent would suggest, would
 * leave a position between them for a caret to sit in; typing there would turn
 * a point comment into a range the reviewer never drew.
 *
 * What the node holds is the model's own markers, not a copy of their fields:
 * an attribute added to `CT_MoveBookmark` is carried without being named here.
 *
 * A range that spans a bookmark boundary, a comment reference or captured
 * markup is not in this class and needs no anchor: each of those is an inline
 * node that carries a mark, so the mark path already keeps the range. Only a
 * range with nothing whatever between its two markers has no carrier.
 */

import { panic } from "better-result";

import type { ParagraphContent } from "../types/content";

/** A marker that opens a range the editor otherwise projects as a mark. */
type RangeStartType = Extract<ParagraphContent["type"], `${string}RangeStart`>;
/** A marker that closes one. */
type RangeEndType = Extract<ParagraphContent["type"], `${string}RangeEnd`>;

/**
 * Which end closes which start.
 *
 * Both sides are derived from the model's own content union rather than
 * listed, so a `…RangeStart` added to `ParagraphContent` fails this
 * `satisfies` until somebody has said what closes it.
 */
export const RANGE_END_FOR_START = {
  commentRangeStart: "commentRangeEnd",
  moveFromRangeStart: "moveFromRangeEnd",
  moveToRangeStart: "moveToRangeEnd",
} as const satisfies Record<RangeStartType, RangeEndType>;

/**
 * The `rangeAnchor` node's attributes: a start and the end that closes it,
 * with nothing between them.
 *
 * Distributing over the start types pairs each start with exactly its own end,
 * so a comment start holding a move end is not a value this type has.
 */
export type RangeAnchorAttrs = {
  [Start in RangeStartType]: {
    start: Extract<ParagraphContent, { type: Start }>;
    end: Extract<ParagraphContent, { type: (typeof RANGE_END_FOR_START)[Start] }>;
  };
}[RangeStartType];

const RANGE_MARKER_TYPES: ReadonlySet<string> = new Set<string>([
  ...Object.keys(RANGE_END_FOR_START),
  ...Object.values(RANGE_END_FOR_START),
]);

/** Whether this content is one of the six range markers, either half. */
export const isRangeMarker = (content: ParagraphContent): boolean =>
  RANGE_MARKER_TYPES.has(content.type);

const isRangeStart = (
  content: ParagraphContent,
): content is Extract<ParagraphContent, { type: RangeStartType }> =>
  content.type in RANGE_END_FOR_START;

const isRangeEnd = (
  content: ParagraphContent,
): content is Extract<ParagraphContent, { type: RangeEndType }> =>
  content.type === "commentRangeEnd" ||
  content.type === "moveFromRangeEnd" ||
  content.type === "moveToRangeEnd";

export type EmptyRangePlan = {
  /** The anchor to emit in place of the marker at this index. */
  anchorAt: ReadonlyMap<number, RangeAnchorAttrs>;
  /** An index whose marker an anchor at an earlier index already holds. */
  closedAt: ReadonlySet<number>;
};

const NO_EMPTY_RANGES: EmptyRangePlan = {
  anchorAt: new Map(),
  closedAt: new Set(),
};

/**
 * The empty ranges in one inline sequence, by the index each one's start sits
 * at.
 *
 * The window a pair may close within is a maximal run of range markers and
 * nothing else, which is what "spans no content" means once the transparent
 * wrappers are lifted. Nesting inside such a run is not meaningful — neither
 * order covers anything — so the pairs are taken nearest first and each anchor
 * stays at its own start's index. A start whose end is outside the run is a
 * range over content and keeps the mark path; so does an end with no start.
 */
export const planEmptyRanges = (content: readonly ParagraphContent[]): EmptyRangePlan => {
  const anchorAt = new Map<number, RangeAnchorAttrs>();
  const closedAt = new Set<number>();

  let index = 0;
  while (index < content.length) {
    const item = content[index];
    if (item === undefined || !isRangeMarker(item)) {
      index += 1;
      continue;
    }
    let runEnd = index;
    while (runEnd < content.length) {
      const candidate = content[runEnd];
      if (candidate === undefined || !isRangeMarker(candidate)) {
        break;
      }
      runEnd += 1;
    }
    pairWithinRun({ content, from: index, to: runEnd, anchorAt, closedAt });
    index = runEnd;
  }

  if (anchorAt.size === 0) {
    return NO_EMPTY_RANGES;
  }
  return { anchorAt, closedAt };
};

/**
 * The two markers as one anchor, or `null` when the closer does not close this
 * start.
 *
 * The branches restate no pairing: `RangeAnchorAttrs` is derived from
 * {@link RANGE_END_FOR_START}, so a branch that named the wrong end would not
 * be assignable, and a start type the model gains fails the exhaustiveness
 * check below.
 */
const pairedAnchor = (
  opener: Extract<ParagraphContent, { type: RangeStartType }>,
  closer: ParagraphContent,
): RangeAnchorAttrs | null => {
  switch (opener.type) {
    case "commentRangeStart":
      return closer.type === "commentRangeEnd" && closer.id === opener.id
        ? { start: opener, end: closer }
        : null;
    case "moveFromRangeStart":
      return closer.type === "moveFromRangeEnd" && closer.id === opener.id
        ? { start: opener, end: closer }
        : null;
    case "moveToRangeStart":
      return closer.type === "moveToRangeEnd" && closer.id === opener.id
        ? { start: opener, end: closer }
        : null;
    default: {
      const unsupported: never = opener;
      panic(`Unsupported range start: ${JSON.stringify(unsupported)}`);
    }
  }
};

type PairWithinRunOptions = {
  content: readonly ParagraphContent[];
  from: number;
  to: number;
  anchorAt: Map<number, RangeAnchorAttrs>;
  closedAt: Set<number>;
};

const pairWithinRun = ({ content, from, to, anchorAt, closedAt }: PairWithinRunOptions): void => {
  const endIndexesByTypeAndId = new Map<string, number[]>();
  for (let end = from; end < to; end += 1) {
    const closer = content[end];
    if (closer === undefined || !isRangeEnd(closer)) {
      continue;
    }
    const key = rangeKey(closer.type, closer.id);
    const indexes = endIndexesByTypeAndId.get(key);
    if (indexes === undefined) {
      endIndexesByTypeAndId.set(key, [end]);
      continue;
    }
    indexes.push(end);
  }

  const nextEndByTypeAndId = new Map<string, number>();
  for (let start = from; start < to; start += 1) {
    const opener = content[start];
    if (opener === undefined || closedAt.has(start) || !isRangeStart(opener)) {
      continue;
    }
    const key = rangeKey(RANGE_END_FOR_START[opener.type], opener.id);
    const indexes = endIndexesByTypeAndId.get(key);
    if (indexes === undefined) {
      continue;
    }
    let nextEnd = nextEndByTypeAndId.get(key) ?? 0;
    while (nextEnd < indexes.length && (indexes[nextEnd] ?? -1) <= start) {
      nextEnd += 1;
    }
    const end = indexes[nextEnd];
    if (end === undefined) {
      nextEndByTypeAndId.set(key, nextEnd);
      continue;
    }
    nextEndByTypeAndId.set(key, nextEnd + 1);
    const closer = content[end];
    if (closer === undefined) {
      continue;
    }
    const anchor = pairedAnchor(opener, closer);
    if (anchor === null) {
      continue;
    }
    anchorAt.set(start, anchor);
    closedAt.add(start);
    closedAt.add(end);
  }
};

const rangeKey = (type: ParagraphContent["type"], id: number): string => `${type}:${id}`;
