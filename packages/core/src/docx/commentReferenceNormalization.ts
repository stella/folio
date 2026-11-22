import { panic } from "better-result";

import {
  PARSE_WARNING_CODES,
  type PreservedMarkup,
} from "@stll/docx-core/model";
import { preservedCommentRangeMarkersAt } from "@stll/docx-core/ops";

import type {
  BlockContent,
  Comment,
  DocumentBody,
  Endnote,
  Footnote,
  HeaderFooter,
  Paragraph,
  ParagraphContent,
  Table,
} from "../types/document";
import { sliceText } from "../utils/sliceText";
import { completeCommentReferences } from "./commentReferenceCompletion";
import { rebuildBodySections } from "./documentParser";
import { InlineContentRemovals } from "./inlineContentRemovals";
import {
  visitInlineContentSlots,
  visitSlotTextBoxes,
} from "./paragraphTraversal";
import { getRunText } from "./runParser";

/** The codes this normalisation is reported under, owned here, not at the caller. */
export const DANGLING_COMMENT_REFERENCE_WARNING: "dangling-comment-reference" =
  PARSE_WARNING_CODES.danglingCommentReference;
export const UNBALANCED_COMMENT_RANGE_WARNING: "unbalanced-comment-range" =
  PARSE_WARNING_CODES.unbalancedCommentRange;
export const UNPAIRED_COMMENT_RANGE_BOUNDARY_WARNING: "unpaired-comment-range-boundary" =
  PARSE_WARNING_CODES.unpairedCommentRangeBoundary;
export const DUPLICATE_COMMENT_REFERENCE_WARNING: "duplicate-comment-reference" =
  PARSE_WARNING_CODES.duplicateCommentReference;
export const DROPPED_COMMENT_RANGE_WARNING: "dropped-comment-range" =
  PARSE_WARNING_CODES.droppedCommentRange;
export const MISSING_COMMENT_REFERENCE_WARNING: "missing-comment-reference" =
  PARSE_WARNING_CODES.missingCommentReference;

type NormalizeCommentReferencesInput = {
  documentBody: DocumentBody;
  comments: readonly Comment[];
  headers?: Map<string, HeaderFooter>;
  footers?: Map<string, HeaderFooter>;
  footnotes?: readonly Footnote[];
  endnotes?: readonly Endnote[];
};

type NormalizeCommentReferencesResult = {
  removedDanglingReferences: number;
  reanchoredUnbalancedRanges: number;
  /** Boundaries with no partner, dropped because a reference anchors their comment. */
  droppedUnpairedBoundaries: number;
  /** References repeating one their comment already has, dropped. */
  droppedDuplicateReferences: number;
  /** Full ranges after a comment's first, dropped: a comment holds one range. */
  droppedRanges: DroppedCommentRange[];
  /** References added after a kept range that no reference named. */
  completedReferences: number;
};

/** A comment's later full range, dropped, and the text it covered. */
export type DroppedCommentRange = { id: number; text: string };

/** Where a range marker sits, which decides how it can be repaired. */
const COMMENT_RANGE_MARKER_SITES = {
  /** Inside a paragraph: a reference can replace it, or it can be removed. */
  inline: "inline",
  /** Between blocks: it can be removed from the blocks that hold it. */
  block: "block",
  /** Between a table's rows or a row's cells, as positioned markup: it can be removed. */
  positioned: "positioned",
} as const;

type CommentRangeMarkerSite =
  | {
      type: typeof COMMENT_RANGE_MARKER_SITES.inline;
      content: ParagraphContent[];
      index: number;
    }
  | {
      type: typeof COMMENT_RANGE_MARKER_SITES.block;
      blocks: BlockContent[];
      index: number;
    }
  | {
      type: typeof COMMENT_RANGE_MARKER_SITES.positioned;
      /** The table or row whose positioned markup holds the boundary. */
      owner: { preserved?: PreservedMarkup };
      child: PreservedChildOf;
      /** The nearest paragraph a point can take the boundary's place in. */
      nearest: () => PointPlacement | undefined;
    };

/**
 * Where a point reference goes in place of a boundary no paragraph holds:
 * the start of the next cell's first paragraph, else the end of the
 * previous cell's last one.
 */
type PointPlacement = { paragraph: Paragraph; at: "start" | "end" };

/**
 * What a range opens and closes within: the body, a header, footer, note or
 * comment, or a text box inside one.
 */
type MarkerStory = { paragraphCount: number };

/**
 * A nonempty run among the comment anchors. Paragraph ordinals count the
 * intervening breaks without adding blank paragraphs to the tape.
 */
type TextPiece = {
  order: number;
  story: MarkerStory;
  paragraphOrdinal: number;
  text: string;
};

type PreservedChildOf = NonNullable<PreservedMarkup["children"]>[number];

type CommentRangeMarkerRef = {
  site: CommentRangeMarkerSite;
  story: MarkerStory;
  id: number;
  type: "commentRangeStart" | "commentRangeEnd";
  /** Its place among every comment anchor, in document order. */
  order: number;
};

type CommentReferenceRef = {
  content: ParagraphContent[];
  index: number;
  id: number;
  order: number;
};

type BlockCommentRangeMarker = Extract<
  BlockContent,
  { type: "commentRangeStart" } | { type: "commentRangeEnd" }
>;

export const normalizeCommentReferences = ({
  documentBody,
  comments,
  headers,
  footers,
  footnotes,
  endnotes,
}: NormalizeCommentReferencesInput): NormalizeCommentReferencesResult => {
  const validCommentIds = new Set(comments.map((comment) => comment.id));
  const rangeMarkers: CommentRangeMarkerRef[] = [];
  const references: CommentReferenceRef[] = [];
  const referencedIds = new Set<number>();
  const removals = new InlineContentRemovals();
  const tape: TextPiece[] = [];
  let removedDanglingReferences = 0;
  let order = 0;
  let story: MarkerStory = { paragraphCount: 0 };
  const inOwnStory = (normalize: () => void): void => {
    const outer = story;
    story = { paragraphCount: 0 };
    normalize();
    story = outer;
  };

  // Reading, repairing and removing run as three passes because a removal
  // shifts every later index in its array: the indices collected here stay
  // valid until `removals.apply()` below.
  const normalizeParagraph = (paragraph: Paragraph): void => {
    const paragraphOrdinal = story.paragraphCount++;
    visitInlineContentSlots(paragraph, (slot) => {
      const { content, index, item } = slot;
      visitSlotTextBoxes(slot, (blocks) => {
        inOwnStory(() => normalizeBlocks(blocks));
      });
      if (item.type === "run") {
        const text = getRunText(item);
        if (text !== "") tape.push({ order, story, paragraphOrdinal, text });
        return;
      }
      if (isCommentMarker(item) && !validCommentIds.has(item.id)) {
        removals.mark({ content, index });
        removedDanglingReferences += 1;
        return;
      }
      if (item.type === "commentReference") {
        referencedIds.add(item.id);
        references.push({ content, index, id: item.id, order: order++ });
      }
      if (isCommentRangeMarker(item)) {
        rangeMarkers.push({
          site: { type: COMMENT_RANGE_MARKER_SITES.inline, content, index },
          story,
          id: item.id,
          type: item.type,
          order: order++,
        });
      }
    });
  };

  // A range may also open or close between a table's rows or a row's cells.
  // The parser keeps those markers as positioned markup, so they are read here
  // in document order; without them a range that closes after the last row
  // looks unbalanced and its paragraph-side marker would be rewritten.
  const collectPositionedMarkers = ({
    owner,
    gap,
    nearest,
  }: CollectPositionedMarkersOptions): void => {
    for (const child of owner.preserved?.children ?? []) {
      // One child at a time, so each boundary keeps the markup it came from.
      for (const { id, type } of preservedCommentRangeMarkersAt([child], gap)) {
        if (validCommentIds.has(id))
          rangeMarkers.push({
            site: {
              type: COMMENT_RANGE_MARKER_SITES.positioned,
              owner,
              child,
              nearest,
            },
            story,
            id,
            type,
            order: order++,
          });
      }
    }
  };

  const normalizeTable = (table: Table): void => {
    const cells = table.rows.flatMap((row) => row.cells);
    // Placements are read only for a boundary a gap holds, and once per
    // table, so a table with no positioned boundary costs nothing more.
    let points: ((following: number) => PointPlacement | undefined) | undefined;
    const nearestAfter = (following: number) => () =>
      (points ??= nearestPoints(cells))(following);
    // The cells before a gap, counted in document order.
    let passed = 0;
    for (const [rowIndex, row] of table.rows.entries()) {
      collectPositionedMarkers({
        owner: table,
        gap: { at: rowIndex, childCount: table.rows.length },
        nearest: nearestAfter(passed),
      });
      for (const [cellIndex, cell] of row.cells.entries()) {
        collectPositionedMarkers({
          owner: row,
          gap: { at: cellIndex, childCount: row.cells.length },
          nearest: nearestAfter(passed),
        });
        normalizeBlocks(cell.content);
        passed += 1;
      }
      collectPositionedMarkers({
        owner: row,
        gap: { at: row.cells.length, childCount: row.cells.length },
        nearest: nearestAfter(passed),
      });
    }
    collectPositionedMarkers({
      owner: table,
      gap: { at: table.rows.length, childCount: table.rows.length },
      nearest: nearestAfter(passed),
    });
  };

  const normalizeBlock = (blocks: BlockContent[], index: number): void => {
    // SAFETY: called with an index of `blocks`.
    const block = blocks[index]!;
    switch (block.type) {
      // A range may open or close between blocks; it pairs with the marker
      // inside a paragraph across the block edge.
      case "commentRangeStart":
      case "commentRangeEnd":
        if (validCommentIds.has(block.id)) {
          rangeMarkers.push({
            site: { type: COMMENT_RANGE_MARKER_SITES.block, blocks, index },
            story,
            id: block.id,
            type: block.type,
            order: order++,
          });
        }
        return;
      case "paragraph":
        normalizeParagraph(block);
        return;
      case "table":
        normalizeTable(block);
        return;
      case "blockSdt":
      case "blockCustomXml":
        normalizeBlocks(block.content);
        return;
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return;
      default: {
        const unhandled: never = block;
        return unhandled;
      }
    }
  };

  const normalizeBlocks = (blocks: BlockContent[]): void => {
    for (const index of blocks.keys()) {
      normalizeBlock(blocks, index);
    }
    // A marker between blocks for a comment the package does not define names
    // no comment folio can show, so it goes back to the bytes it was, kept in
    // place: the next open reads it the same way.
    for (const [index, block] of blocks.entries()) {
      if (
        (block.type === "commentRangeStart" ||
          block.type === "commentRangeEnd") &&
        !validCommentIds.has(block.id)
      ) {
        blocks[index] = unreadBlockMarker(block);
      }
    }
  };

  for (const { content } of [
    documentBody,
    ...(headers?.values() ?? []),
    ...(footers?.values() ?? []),
    ...(footnotes ?? []),
    ...(endnotes ?? []),
  ]) {
    inOwnStory(() => normalizeBlocks(content));
  }
  for (const comment of comments) {
    inOwnStory(() => {
      for (const paragraph of comment.content) {
        normalizeParagraph(paragraph);
      }
    });
  }

  const repaired = repairUnpairedCommentRanges(rangeMarkers, referencedIds);
  for (const slot of repaired.removed) removals.mark(slot);
  const duplicates = duplicateReferences(references, repaired.kept);
  for (const slot of duplicates) removals.mark(slot);
  removals.apply();
  removeBlocks(repaired.removedBlocks, documentBody);
  // Points go in once removals have run, so no index they hold shifts.
  for (const { id, placement } of repaired.placed) {
    const reference: ParagraphContent = { type: "commentReference", id };
    if (placement.at === "start")
      placement.paragraph.content.unshift(reference);
    else placement.paragraph.content.push(reference);
  }
  // Once every range left is one the comment keeps, a range no reference
  // names gets its reference after the end, where a reference follows a range.
  const completedReferences = completeCommentReferences(
    documentBody.content,
    ...[...(headers?.values() ?? [])].map(({ content }) => content),
    ...[...(footers?.values() ?? [])].map(({ content }) => content),
    ...(footnotes ?? []).map(({ content }) => content),
    ...(endnotes ?? []).map(({ content }) => content),
  );

  return {
    removedDanglingReferences,
    reanchoredUnbalancedRanges: repaired.reanchored,
    droppedUnpairedBoundaries: repaired.droppedBoundaries,
    droppedDuplicateReferences: duplicates.length,
    droppedRanges: repaired.droppedRanges.map(({ id, start, end }) => ({
      id,
      text: textBetween(tape, start, end),
    })),
    completedReferences,
  };
};

type CollectPositionedMarkersOptions = {
  owner: { preserved?: PreservedMarkup };
  gap: { at: number; childCount: number };
  nearest: () => PointPlacement | undefined;
};

/**
 * The point placement for a gap before cell `following` of `cells` (in
 * document order): the start of the first paragraph from that cell on, else
 * the end of the last paragraph before it. Every gap's answer is read off two
 * passes over the cells.
 */
const nearestPoints = (
  cells: readonly Table["rows"][number]["cells"][number][],
): ((following: number) => PointPlacement | undefined) => {
  const firstFrom: (Paragraph | undefined)[] = [];
  firstFrom[cells.length] = undefined;
  for (let index = cells.length - 1; index >= 0; index -= 1)
    firstFrom[index] =
      edgeParagraph(cells[index]?.content ?? [], "first") ??
      firstFrom[index + 1];
  const lastBefore: (Paragraph | undefined)[] = [undefined];
  for (const [index, cell] of cells.entries())
    lastBefore[index + 1] =
      edgeParagraph(cell.content, "last") ?? lastBefore[index];
  return (following) => {
    const first = firstFrom[following];
    if (first !== undefined) return { paragraph: first, at: "start" };
    const last = lastBefore[following];
    return last === undefined ? undefined : { paragraph: last, at: "end" };
  };
};

/** The first or last paragraph of `blocks`, looking inside tables and controls. */
const edgeParagraph = (
  blocks: readonly BlockContent[],
  edge: "first" | "last",
): Paragraph | undefined => {
  const ordered = edge === "first" ? blocks : blocks.toReversed();
  for (const block of ordered) {
    const found = paragraphAtEdge(block, edge);
    if (found !== undefined) return found;
  }
  return undefined;
};

const paragraphAtEdge = (
  block: BlockContent,
  edge: "first" | "last",
): Paragraph | undefined => {
  switch (block.type) {
    case "paragraph":
      return block;
    case "table": {
      const cells = block.rows.flatMap((row) => row.cells);
      const ordered = edge === "first" ? cells : cells.toReversed();
      for (const cell of ordered) {
        const found = edgeParagraph(cell.content, edge);
        if (found !== undefined) return found;
      }
      return undefined;
    }
    case "blockSdt":
    case "blockCustomXml":
      return edgeParagraph(block.content, edge);
    case "preservedBlock":
    case "commentRangeStart":
    case "commentRangeEnd":
    case "bookmarkStart":
    case "bookmarkEnd":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return undefined;
    default: {
      const unhandled: never = block;
      return unhandled;
    }
  }
};

/** At most this many UTF-16 code units of a dropped range are reported. */
const DROPPED_RANGE_TEXT_LIMIT = 200;

/**
 * Read the story's runs as written, joining paragraphs with line breaks
 * between them. Binary search locates the first piece, and only the bounded
 * prefix is read.
 */
const textBetween = (
  tape: readonly TextPiece[],
  start: CommentRangeMarkerRef,
  end: CommentRangeMarkerRef,
): string => {
  let low = 0;
  let high = tape.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if ((tape[middle]?.order ?? Number.POSITIVE_INFINITY) > start.order)
      high = middle;
    else low = middle + 1;
  }
  let text = "";
  let previousParagraph: number | undefined;
  for (let index = low; index < tape.length; index += 1) {
    const piece = tape[index];
    if (piece === undefined || piece.order > end.order) break;
    if (piece.story !== start.story) continue;
    const breaks =
      piece.paragraphOrdinal - (previousParagraph ?? piece.paragraphOrdinal);
    text += "\n".repeat(
      Math.min(breaks, DROPPED_RANGE_TEXT_LIMIT - text.length),
    );
    text += piece.text;
    if (text.length >= DROPPED_RANGE_TEXT_LIMIT) break;
    previousParagraph = piece.paragraphOrdinal;
  }
  return sliceText({ text, units: "codeUnits", end: DROPPED_RANGE_TEXT_LIMIT });
};

type BlockRemoval = { blocks: BlockContent[]; index: number };

/**
 * Remove each listed block once every marker has been read. The body's
 * sections partition its blocks, so they are rebuilt over what remains.
 */
const removeBlocks = (
  removals: readonly BlockRemoval[],
  documentBody: DocumentBody,
): void => {
  const byBlocks = new Map<BlockContent[], Set<number>>();
  for (const { blocks, index } of removals) {
    const indexes = byBlocks.get(blocks);
    if (indexes === undefined) byBlocks.set(blocks, new Set([index]));
    else indexes.add(index);
  }
  for (const [blocks, indexes] of byBlocks) {
    const kept = blocks.filter((_, index) => !indexes.has(index));
    blocks.length = 0;
    for (const block of kept) blocks.push(block);
  }
  if (byBlocks.has(documentBody.content)) rebuildBodySections(documentBody);
};

/**
 * A comment's references after its first: each comment keeps one. Where it
 * keeps one paired range, the reference it keeps is the first after that
 * range's end, as a reference follows its range; otherwise its first.
 */
const duplicateReferences = (
  references: readonly CommentReferenceRef[],
  kept: readonly CommentRangeMarkerRef[],
): CommentReferenceRef[] => {
  const byId = new Map<number, CommentReferenceRef[]>();
  for (const reference of references) {
    const own = byId.get(reference.id);
    if (own === undefined) byId.set(reference.id, [reference]);
    else own.push(reference);
  }
  const ends = new Map<number, number[]>();
  for (const marker of kept) {
    if (marker.type !== "commentRangeEnd") continue;
    const own = ends.get(marker.id);
    if (own === undefined) ends.set(marker.id, [marker.order]);
    else own.push(marker.order);
  }
  const dropped: CommentReferenceRef[] = [];
  for (const [id, own] of byId) {
    if (own.length < 2) continue;
    const [end, ...more] = ends.get(id) ?? [];
    const keeper =
      (end !== undefined && more.length === 0
        ? own.find(({ order: at }) => at > end)
        : undefined) ?? own.at(0);
    for (const reference of own)
      if (reference !== keeper) dropped.push(reference);
  }
  return dropped;
};

/** A block-level range marker of an undefined comment, as the bytes it was. */
const unreadBlockMarker = (block: BlockCommentRangeMarker): BlockContent => ({
  type: "preservedBlock",
  xml: `<w:${block.type} w:id="${String(block.id)}"${
    block.displacedByCustomXml === undefined
      ? ""
      : ` w:displacedByCustomXml="${block.displacedByCustomXml}"`
  }/>`,
});

type RepairUnpairedCommentRangesResult = {
  reanchored: number;
  /** Point references to insert where a boundary no paragraph holds was. */
  placed: { id: number; placement: PointPlacement }[];
  /** Full ranges after a comment's first. */
  droppedRanges: {
    id: number;
    start: CommentRangeMarkerRef;
    end: CommentRangeMarkerRef;
  }[];
  /** Boundaries with no partner, dropped. */
  droppedBoundaries: number;
  /** Inline boundaries to remove once every marker has been read. */
  removed: { content: ParagraphContent[]; index: number }[];
  /** Boundaries between blocks to remove once every marker has been read. */
  removedBlocks: BlockRemoval[];
  /** The boundaries the repair leaves in place, in document order. */
  kept: CommentRangeMarkerRef[];
};

/**
 * Leave each comment one reference and one full range. A comment's first
 * full range is kept; each later one is dropped and reported with its text,
 * since a comment holds one range. A boundary with no partner anchors its
 * comment as a point when nothing else does: the first inline one becomes
 * the comment's reference, and one between a table's rows or a row's cells
 * becomes a reference in the nearest paragraph (see {@link nearestPoints}).
 * Every other such boundary is dropped, the comment staying anchored by its
 * reference. A comment left with no anchor keeps its boundaries, so no
 * comment loses its last anchor.
 */
const repairUnpairedCommentRanges = (
  rangeMarkers: readonly CommentRangeMarkerRef[],
  referencedIds: ReadonlySet<number>,
): RepairUnpairedCommentRangesResult => {
  const byId = new Map<number, CommentRangeMarkerRef[]>();
  for (const marker of rangeMarkers) {
    const markers = byId.get(marker.id);
    if (markers) {
      markers.push(marker);
      continue;
    }
    byId.set(marker.id, [marker]);
  }

  const result: RepairUnpairedCommentRangesResult = {
    reanchored: 0,
    placed: [],
    droppedRanges: [],
    droppedBoundaries: 0,
    removed: [],
    removedBlocks: [],
    kept: [],
  };
  for (const [id, markers] of byId) {
    const { ranges, lone } = pairedRanges(markers);
    const [range, ...later] = ranges;
    for (const [start, end] of later) {
      result.droppedRanges.push({ id, start, end });
      removeMarker(start.site, result);
      removeMarker(end.site, result);
    }
    let point: CommentRangeMarkerRef | undefined;
    if (!referencedIds.has(id)) {
      point = pointFor({ id, range, lone, result });
      if (point === undefined) {
        result.kept.push(...(range ?? []), ...lone);
        continue;
      }
    }
    result.kept.push(...(range ?? []));
    for (const marker of lone) {
      if (marker === point) continue;
      removeMarker(marker.site, result);
      result.droppedBoundaries += 1;
    }
  }
  return result;
};

type PointForOptions = {
  id: number;
  range: readonly [CommentRangeMarkerRef, CommentRangeMarkerRef] | undefined;
  lone: readonly CommentRangeMarkerRef[];
  result: RepairUnpairedCommentRangesResult;
};

/**
 * Anchor a comment no reference anchors at one of its lone boundaries: the
 * first inline one becomes its reference in place; failing that, when the
 * comment has no full range either, the first between a table's rows or a
 * row's cells becomes a reference in the nearest paragraph. Returns the
 * boundary that became the point, or undefined when none could.
 */
const pointFor = ({
  id,
  range,
  lone,
  result,
}: PointForOptions): CommentRangeMarkerRef | undefined => {
  const inline = lone.find(
    ({ site }) => site.type === COMMENT_RANGE_MARKER_SITES.inline,
  );
  if (inline?.site.type === COMMENT_RANGE_MARKER_SITES.inline) {
    inline.site.content[inline.site.index] = { type: "commentReference", id };
    result.reanchored += 1;
    return inline;
  }
  if (range !== undefined) return undefined;
  for (const marker of lone) {
    if (marker.site.type !== COMMENT_RANGE_MARKER_SITES.positioned) continue;
    const placement = marker.site.nearest();
    if (placement === undefined) continue;
    removeMarker(marker.site, result);
    result.placed.push({ id, placement });
    result.reanchored += 1;
    return marker;
  }
  return undefined;
};

type PairedRangesResult = {
  /** Each full range, in the document order of their starts. */
  ranges: [CommentRangeMarkerRef, CommentRangeMarkerRef][];
  /** The boundaries no unambiguous pairing claims. */
  lone: CommentRangeMarkerRef[];
};

/** Pair a comment's boundaries, story by story: no range leaves its story. */
const pairedRanges = (
  markers: readonly CommentRangeMarkerRef[],
): PairedRangesResult => {
  const result: PairedRangesResult = { ranges: [], lone: [] };
  const byStory = new Map<MarkerStory, CommentRangeMarkerRef[]>();
  for (const marker of markers) {
    const own = byStory.get(marker.story);
    if (own === undefined) byStory.set(marker.story, [marker]);
    else own.push(marker);
  }
  for (const own of byStory.values()) pairStoryRanges(own, result);
  result.ranges.sort(([first], [second]) => first.order - second.order);
  result.lone.sort((first, second) => first.order - second.order);
  return result;
};

/**
 * Pair a comment's boundaries within one story only where they all pair. The
 * ranges a stretch of as many ends as starts nests or overlaps share an id,
 * so nothing tells which end closes which: the first start opens the first
 * range and the first end closes it, as no anchor before it repeats, and so
 * on in turn. Elsewhere only a start whose very next boundary is an end pairs,
 * so a lone start before a full range never joins that range's end and widens
 * what the comment covers.
 */
const pairStoryRanges = (
  markers: readonly CommentRangeMarkerRef[],
  { ranges, lone }: PairedRangesResult,
): void => {
  const closes = balancedCloses(markers);
  for (let index = 0; index < markers.length; index += 1) {
    const close = closes.get(index);
    if (close === undefined) {
      // SAFETY: index < markers.length.
      lone.push(markers[index]!);
      continue;
    }
    const stretch = markers.slice(index, close + 1);
    const ends = stretch.filter(({ type }) => type === "commentRangeEnd");
    let ordinal = 0;
    for (const marker of stretch) {
      if (marker.type === "commentRangeEnd") continue;
      const end = ends[ordinal++];
      if (end === undefined) panic("A balanced stretch lacks an end.");
      ranges.push([marker, end]);
    }
    index = close;
  }
};

/** Where each start that opens a balanced stretch has its closing end. */
const balancedCloses = (
  markers: readonly CommentRangeMarkerRef[],
): Map<number, number> => {
  const closes = new Map<number, number>();
  const open: number[] = [];
  for (const [index, { type }] of markers.entries()) {
    if (type === "commentRangeStart") {
      open.push(index);
      continue;
    }
    const start = open.pop();
    if (start !== undefined) closes.set(start, index);
  }
  return closes;
};

/**
 * Take a boundary out: an inline one or one between blocks once every marker
 * has been read, a positioned one from its markup. A dropped boundary leaves
 * the package, so a save does not write it back for the next open to repair.
 */
const removeMarker = (
  site: CommentRangeMarkerSite,
  result: RepairUnpairedCommentRangesResult,
): void => {
  switch (site.type) {
    case COMMENT_RANGE_MARKER_SITES.inline:
      result.removed.push({ content: site.content, index: site.index });
      return;
    case COMMENT_RANGE_MARKER_SITES.block:
      result.removedBlocks.push({ blocks: site.blocks, index: site.index });
      return;
    case COMMENT_RANGE_MARKER_SITES.positioned: {
      const { owner, child } = site;
      const children = owner.preserved?.children ?? [];
      const at = children.indexOf(child);
      if (at < 0) return;
      children.splice(at, 1);
      // An empty record is never written.
      if (children.length > 0 || owner.preserved === undefined) return;
      const { children: _empty, ...rest } = owner.preserved;
      if (Reflect.ownKeys(rest).length === 0) delete owner.preserved;
      else owner.preserved = rest;
      return;
    }
    default: {
      const unhandled: never = site;
      return unhandled;
    }
  }
};

const isCommentMarker = (
  content: ParagraphContent,
): content is Extract<
  ParagraphContent,
  | { type: "commentRangeStart" }
  | { type: "commentRangeEnd" }
  | { type: "commentReference" }
> =>
  content.type === "commentRangeStart" ||
  content.type === "commentRangeEnd" ||
  content.type === "commentReference";

const isCommentRangeMarker = (
  content: ParagraphContent,
): content is Extract<
  ParagraphContent,
  { type: "commentRangeStart" } | { type: "commentRangeEnd" }
> => content.type === "commentRangeStart" || content.type === "commentRangeEnd";
