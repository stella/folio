import {
  PARSE_WARNING_CODES,
  type PreservedChild,
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
import { InlineContentRemovals } from "./inlineContentRemovals";
import { visitInlineContentSlots } from "./paragraphTraversal";

/** The codes this normalisation is reported under, owned here, not at the caller. */
export const DANGLING_COMMENT_REFERENCE_WARNING: "dangling-comment-reference" =
  PARSE_WARNING_CODES.danglingCommentReference;
export const UNBALANCED_COMMENT_RANGE_WARNING: "unbalanced-comment-range" =
  PARSE_WARNING_CODES.unbalancedCommentRange;
export const UNPAIRED_COMMENT_RANGE_BOUNDARY_WARNING: "unpaired-comment-range-boundary" =
  PARSE_WARNING_CODES.unpairedCommentRangeBoundary;

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
};

/** Where a range marker sits, which decides how it can be repaired. */
const COMMENT_RANGE_MARKER_SITES = {
  /** Inside a paragraph: a reference can replace it, or it can be removed. */
  inline: "inline",
  /** Between blocks: it can go back to the bytes it was, read by nobody. */
  block: "block",
  /** Between a table's rows or a row's cells, as positioned markup: kept as written. */
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
  | { type: typeof COMMENT_RANGE_MARKER_SITES.positioned };

type CommentRangeMarkerRef = {
  site: CommentRangeMarkerSite;
  id: number;
  type: "commentRangeStart" | "commentRangeEnd";
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
  const referencedIds = new Set<number>();
  const removals = new InlineContentRemovals();
  let removedDanglingReferences = 0;

  // Reading, repairing and removing run as three passes because a removal
  // shifts every later index in its array: the indices collected here stay
  // valid until `removals.apply()` below.
  const normalizeParagraph = (paragraph: Paragraph): void => {
    visitInlineContentSlots(paragraph, ({ content, index, item }) => {
      if (isCommentMarker(item) && !validCommentIds.has(item.id)) {
        removals.mark({ content, index });
        removedDanglingReferences += 1;
        return;
      }
      if (item.type === "commentReference") referencedIds.add(item.id);
      if (isCommentRangeMarker(item)) {
        rangeMarkers.push({
          site: { type: COMMENT_RANGE_MARKER_SITES.inline, content, index },
          id: item.id,
          type: item.type,
        });
      }
    });
  };

  // A range may also open or close between a table's rows or a row's cells.
  // The parser keeps those markers as positioned markup, so they are read here
  // in document order; without them a range that closes after the last row
  // looks unbalanced and its paragraph-side marker would be rewritten.
  const collectPositionedMarkers = (
    children: readonly PreservedChild[] | undefined,
    gap: { at: number; childCount: number },
  ): void => {
    for (const { id, type } of preservedCommentRangeMarkersAt(children, gap)) {
      if (validCommentIds.has(id))
        rangeMarkers.push({
          site: { type: COMMENT_RANGE_MARKER_SITES.positioned },
          id,
          type,
        });
    }
  };

  const normalizeTable = (table: Table): void => {
    const tableChildren = table.preserved?.children;
    for (const [rowIndex, row] of table.rows.entries()) {
      collectPositionedMarkers(tableChildren, {
        at: rowIndex,
        childCount: table.rows.length,
      });
      const rowChildren = row.preserved?.children;
      for (const [cellIndex, cell] of row.cells.entries()) {
        collectPositionedMarkers(rowChildren, {
          at: cellIndex,
          childCount: row.cells.length,
        });
        normalizeBlocks(cell.content);
      }
      collectPositionedMarkers(rowChildren, {
        at: row.cells.length,
        childCount: row.cells.length,
      });
    }
    collectPositionedMarkers(tableChildren, {
      at: table.rows.length,
      childCount: table.rows.length,
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
            id: block.id,
            type: block.type,
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
    // no comment folio can show, so it goes back to the bytes it was: kept in
    // place, read by nobody.
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

  normalizeBlocks(documentBody.content);
  for (const header of headers?.values() ?? []) {
    normalizeBlocks(header.content);
  }
  for (const footer of footers?.values() ?? []) {
    normalizeBlocks(footer.content);
  }
  for (const footnote of footnotes ?? []) {
    normalizeBlocks(footnote.content);
  }
  for (const endnote of endnotes ?? []) {
    normalizeBlocks(endnote.content);
  }
  for (const comment of comments) {
    for (const paragraph of comment.content) {
      normalizeParagraph(paragraph);
    }
  }

  const repaired = repairUnpairedCommentRanges(rangeMarkers, referencedIds);
  for (const slot of repaired.removed) removals.mark(slot);
  removals.apply();

  return {
    removedDanglingReferences,
    reanchoredUnbalancedRanges: repaired.reanchored,
    droppedUnpairedBoundaries: repaired.removed.length + repaired.unread,
  };
};

/** A block-level range marker as the bytes it was, which no reader reads. */
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
  /** Inline boundaries to remove once every marker has been read. */
  removed: { content: ParagraphContent[]; index: number }[];
  /** Boundaries between blocks turned back into the bytes they were. */
  unread: number;
};

/**
 * Leave each comment one reference and its paired range. A boundary with no
 * partner anchors its comment as a point when nothing else does: the first
 * inline one becomes the comment's reference. Every other such boundary is
 * dropped, the comment staying anchored by its reference; a boundary only
 * positioned markup holds stays as written. A comment no reference anchors
 * keeps its boundaries, so no comment loses its last anchor.
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
    removed: [],
    unread: 0,
  };
  for (const [id, markers] of byId) {
    const unpaired = unpairedMarkers(markers);
    let point: CommentRangeMarkerRef | undefined;
    if (!referencedIds.has(id)) {
      point = unpaired.find(
        ({ site }) => site.type === COMMENT_RANGE_MARKER_SITES.inline,
      );
      if (point?.site.type !== COMMENT_RANGE_MARKER_SITES.inline) continue;
      point.site.content[point.site.index] = { type: "commentReference", id };
      result.reanchored += 1;
    }
    for (const marker of unpaired) {
      if (marker !== point) dropUnpairedMarker(marker.site, result);
    }
  }
  return result;
};

/** The boundaries that pair with none, in document order. */
const unpairedMarkers = (
  markers: readonly CommentRangeMarkerRef[],
): CommentRangeMarkerRef[] => {
  const unpaired: CommentRangeMarkerRef[] = [];
  let openStart: CommentRangeMarkerRef | undefined;
  for (const marker of markers) {
    if (marker.type === "commentRangeStart") {
      if (openStart) unpaired.push(marker);
      else openStart = marker;
      continue;
    }
    if (openStart) openStart = undefined;
    else unpaired.push(marker);
  }
  if (openStart) unpaired.push(openStart);
  return unpaired;
};

const dropUnpairedMarker = (
  site: CommentRangeMarkerSite,
  result: RepairUnpairedCommentRangesResult,
): void => {
  switch (site.type) {
    case COMMENT_RANGE_MARKER_SITES.inline:
      result.removed.push({ content: site.content, index: site.index });
      return;
    case COMMENT_RANGE_MARKER_SITES.block: {
      const block = site.blocks[site.index];
      if (
        block?.type !== "commentRangeStart" &&
        block?.type !== "commentRangeEnd"
      )
        return;
      site.blocks[site.index] = unreadBlockMarker(block);
      result.unread += 1;
      return;
    }
    case COMMENT_RANGE_MARKER_SITES.positioned:
      return;
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
