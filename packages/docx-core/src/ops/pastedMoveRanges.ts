/**
 * One range per move-range id: what an insertion does with the move range
 * boundaries its slice carries.
 *
 * A copy of moved text is inserted text, no part of the move: a boundary that
 * stands in the document already is left out, as is one whose whole range the
 * slice holds while the document holds the other half, and one the slice
 * holds twice. A boundary completing a range that has lost it, or one of a
 * range the document does not hold, goes in as it is.
 */
import type { Document, ParagraphContent } from "../model/document";
import { copyOriginBoundaryKeysInBlocks } from "./copyOriginAncestors";
import {
  isMoveBoundaryKind,
  type MoveBoundaryKind,
  rewriteSliceMarkers,
} from "./pastedComments";
import { documentStories, storyBody } from "./stories";
import { textBoxContentsIn } from "./textBoxRouting";

type BoundaryKey = `${MoveBoundaryKind}:${number}`;

const keyOf = (kind: MoveBoundaryKind, id: number): BoundaryKey =>
  `${kind}:${id}`;

const PARTNERS = {
  moveFromRangeStart: "moveFromRangeEnd",
  moveFromRangeEnd: "moveFromRangeStart",
  moveToRangeStart: "moveToRangeEnd",
  moveToRangeEnd: "moveToRangeStart",
} as const satisfies Record<MoveBoundaryKind, MoveBoundaryKind>;

/** The keys of every bookmark and move range boundary standing in `document`. */
const standingBoundaries = (document: Document): Set<string> => {
  const standing = new Set<string>();
  for (const story of documentStories(document))
    for (const blocks of [
      storyBody(document, story).content,
      ...textBoxContentsIn(document, story),
    ])
      for (const key of copyOriginBoundaryKeysInBlocks(blocks))
        standing.add(key);
  return standing;
};

/** Whether some content holds a move range boundary, at any depth. */
export const holdsMoveRange = (
  contents: readonly (readonly ParagraphContent[])[],
): boolean => {
  const found = new Set<MoveBoundaryKind>();
  for (const content of contents)
    rewriteSliceMarkers(content, ({ id, kind }) => {
      if (isMoveBoundaryKind(kind)) found.add(kind);
      return id;
    });
  return found.size > 0;
};

/**
 * `contents` (a slice's, in order, going into `document`) without the move
 * range boundaries that repeat a range; see the module comment. Each content
 * that repeats none is the same.
 */
export const withoutRepeatedMoveRangesIn = (
  document: Document,
  contents: readonly (readonly ParagraphContent[])[],
): (readonly ParagraphContent[])[] => {
  if (!holdsMoveRange(contents)) return [...contents];
  const incoming = new Set<string>();
  for (const content of contents)
    rewriteSliceMarkers(content, ({ id, kind }) => {
      incoming.add(`${kind}:${id}`);
      return id;
    });
  const standing = standingBoundaries(document);
  const taken = new Set<string>();
  return contents.map((content) => {
    let repeats = false;
    const settled = rewriteSliceMarkers(content, ({ id, kind }) => {
      if (!isMoveBoundaryKind(kind)) return id;
      const key = keyOf(kind, id);
      const partner = keyOf(PARTNERS[kind], id);
      if (
        standing.has(key) ||
        taken.has(key) ||
        (standing.has(partner) && incoming.has(partner))
      ) {
        repeats = true;
        return null;
      }
      taken.add(key);
      return id;
    });
    return repeats ? settled : content;
  });
};

/** {@link withoutRepeatedMoveRangesIn} of one content. */
export const withoutRepeatedMoveRanges = (
  document: Document,
  content: readonly ParagraphContent[],
): readonly ParagraphContent[] =>
  withoutRepeatedMoveRangesIn(document, [content]).at(0) ?? content;
