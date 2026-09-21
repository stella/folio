/**
 * The move ranges a paragraph holds, carried across the editor projection.
 *
 * A tracked move is two halves: `w:moveFrom` at the source, `w:moveTo` at the
 * destination, and a `w:moveFromRangeStart` / `w:moveToRangeStart` pair that
 * carries the `w:name` binding one to the other. The editor projects the
 * wrappers as a tracked-change mark with a `moveKind` and had no node for the
 * markers, so `toProseDoc` dropped them: the first save after an edit wrote two
 * unrelated revisions where the document had one move, and the range's own
 * `w:id`, `w:author` and `w:date` went with the name.
 *
 * The markers are not content a caret can sit in, so they ride on the
 * paragraph beside `bookmarks` rather than becoming inline nodes. What is
 * carried is the model's own marker, not a copy of its fields: a field added to
 * `CT_MoveBookmark` is carried here without being named.
 */

import { panic } from "better-result";

import type { ParagraphContent } from "../types/content";

/**
 * The markers that delimit a tracked move, and the source of the union below.
 *
 * Listed once: `MoveRangeMarker` is `Extract`ed from the model's own paragraph
 * content, so a name that is not a marker there extracts nothing and the switch
 * below stops compiling.
 */
const MOVE_RANGE_MARKER_TYPES = [
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
] as const;

export type MoveRangeMarker = Extract<
  ParagraphContent,
  { type: (typeof MOVE_RANGE_MARKER_TYPES)[number] }
>;

export const MOVE_RANGE_MARKER_TYPE_SET: ReadonlySet<string> = new Set(MOVE_RANGE_MARKER_TYPES);

const isMoveRangeMarker = (item: ParagraphContent): item is MoveRangeMarker =>
  MOVE_RANGE_MARKER_TYPE_SET.has(item.type);

/** Which half of a move a marker delimits; the same two values the mark's `moveKind` carries. */
type MoveKind = "moveFrom" | "moveTo";

const kindOf = (marker: MoveRangeMarker): MoveKind => {
  switch (marker.type) {
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
      return "moveFrom";
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return "moveTo";
    default: {
      const unsupported: never = marker;
      panic(`Unsupported move range marker: ${JSON.stringify(unsupported)}`);
    }
  }
};

const opensRange = (marker: MoveRangeMarker): boolean =>
  marker.type === "moveFromRangeStart" || marker.type === "moveToRangeStart";

/** The markers a paragraph holds directly, in document order. */
export const moveRangeMarkersOf = (content: readonly ParagraphContent[]): MoveRangeMarker[] =>
  content.filter(isMoveRangeMarker);

/** Where the paragraph's wrappers of one kind begin and end; `-1` when it holds none. */
const wrapperSpan = (
  content: readonly ParagraphContent[],
  kind: MoveKind,
): { first: number; last: number } => {
  let first = -1;
  let last = -1;
  for (const [index, item] of content.entries()) {
    if (item.type !== kind) {
      continue;
    }
    if (first === -1) {
      first = index;
    }
    last = index;
  }
  return { first, last };
};

/** Where a marker goes among the paragraph's rebuilt content. */
const placementOf = (content: readonly ParagraphContent[], marker: MoveRangeMarker): number => {
  const { first, last } = wrapperSpan(content, kindOf(marker));
  if (opensRange(marker)) {
    return first === -1 ? 0 : first;
  }
  return last === -1 ? content.length : last + 1;
};

/**
 * The carried markers put back around the content they delimit.
 *
 * A range opens before the first wrapper of its kind and closes after the last,
 * which is where the author put it and where Word writes it. A marker whose
 * paragraph holds no wrapper of that kind still has a place: a move spanning
 * several paragraphs opens in the first and closes in the last, so a start with
 * nothing after it opens the paragraph and an end with nothing before it closes
 * it. Dropping those would unpair the range instead of narrowing it.
 */
export const withMoveRanges = (
  content: readonly ParagraphContent[],
  markers: readonly MoveRangeMarker[],
): ParagraphContent[] => {
  if (markers.length === 0) {
    return [...content];
  }
  const placed = new Map<number, MoveRangeMarker[]>();
  for (const marker of markers) {
    const at = placementOf(content, marker);
    const alreadyThere = placed.get(at);
    if (alreadyThere === undefined) {
      placed.set(at, [marker]);
      continue;
    }
    alreadyThere.push(marker);
  }

  // The markers are cloned on the way out: the attr holds them for as long as
  // the editor state does, and the rebuilt document must not share a record
  // with it.
  const rebuilt: ParagraphContent[] = [];
  for (let index = 0; index <= content.length; index += 1) {
    for (const marker of placed.get(index) ?? []) {
      rebuilt.push({ ...marker });
    }
    const item = content[index];
    if (item !== undefined) {
      rebuilt.push(item);
    }
  }
  return rebuilt;
};
