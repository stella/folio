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
import { InlineContentRemovals } from "./inlineContentRemovals";
import { visitInlineContentSlots } from "./paragraphTraversal";

/** The codes this normalisation is reported under, owned here, not at the caller. */
export const DANGLING_COMMENT_REFERENCE_WARNING: "dangling-comment-reference" =
  PARSE_WARNING_CODES.danglingCommentReference;
export const UNBALANCED_COMMENT_RANGE_WARNING: "unbalanced-comment-range" =
  PARSE_WARNING_CODES.unbalancedCommentRange;
export const UNPAIRED_COMMENT_RANGE_BOUNDARY_WARNING: "unpaired-comment-range-boundary" =
  PARSE_WARNING_CODES.unpairedCommentRangeBoundary;
export const DUPLICATE_COMMENT_REFERENCE_WARNING: "duplicate-comment-reference" =
  PARSE_WARNING_CODES.duplicateCommentReference;

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
};

/** Where a range marker sits, which decides how it can be repaired. */
const COMMENT_RANGE_MARKER_SITES = {
  /** Inside a paragraph: a reference can replace it, or it can be removed. */
  inline: "inline",
  /** Between blocks: it can go back to the bytes it was, read by nobody. */
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
    };

type PreservedChildOf = NonNullable<PreservedMarkup["children"]>[number];

type CommentRangeMarkerRef = {
  site: CommentRangeMarkerSite;
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
  let removedDanglingReferences = 0;
  let order = 0;

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
      if (item.type === "commentReference") {
        referencedIds.add(item.id);
        references.push({ content, index, id: item.id, order: order++ });
      }
      if (isCommentRangeMarker(item)) {
        rangeMarkers.push({
          site: { type: COMMENT_RANGE_MARKER_SITES.inline, content, index },
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
  const collectPositionedMarkers = (
    owner: { preserved?: PreservedMarkup },
    gap: { at: number; childCount: number },
  ): void => {
    for (const child of owner.preserved?.children ?? []) {
      // One child at a time, so each boundary keeps the markup it came from.
      for (const { id, type } of preservedCommentRangeMarkersAt([child], gap)) {
        if (validCommentIds.has(id))
          rangeMarkers.push({
            site: { type: COMMENT_RANGE_MARKER_SITES.positioned, owner, child },
            id,
            type,
            order: order++,
          });
      }
    }
  };

  const normalizeTable = (table: Table): void => {
    for (const [rowIndex, row] of table.rows.entries()) {
      collectPositionedMarkers(table, {
        at: rowIndex,
        childCount: table.rows.length,
      });
      for (const [cellIndex, cell] of row.cells.entries()) {
        collectPositionedMarkers(row, {
          at: cellIndex,
          childCount: row.cells.length,
        });
        normalizeBlocks(cell.content);
      }
      collectPositionedMarkers(row, {
        at: row.cells.length,
        childCount: row.cells.length,
      });
    }
    collectPositionedMarkers(table, {
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
  const duplicates = duplicateReferences(references, repaired.kept);
  for (const slot of duplicates) removals.mark(slot);
  removals.apply();

  return {
    removedDanglingReferences,
    reanchoredUnbalancedRanges: repaired.reanchored,
    droppedUnpairedBoundaries:
      repaired.removed.length + repaired.unread + repaired.unpositioned,
    droppedDuplicateReferences: duplicates.length,
  };
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
  /** Boundaries between a table's rows or a row's cells, removed. */
  unpositioned: number;
  /** The boundaries the repair leaves in place, in document order. */
  kept: CommentRangeMarkerRef[];
};

/**
 * Leave each comment one reference and its paired range. A boundary with no
 * partner anchors its comment as a point when nothing else does: the first
 * inline one becomes the comment's reference. Every other such boundary is
 * dropped, the comment staying anchored by its reference, between a table's
 * rows or a row's cells as much as inside a paragraph. A comment no
 * reference anchors keeps its boundaries, so no comment loses its last
 * anchor.
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
    unpositioned: 0,
    kept: [],
  };
  for (const [id, markers] of byId) {
    const unpaired = unpairedMarkers(markers);
    let point: CommentRangeMarkerRef | undefined;
    if (!referencedIds.has(id)) {
      point = unpaired.find(
        ({ site }) => site.type === COMMENT_RANGE_MARKER_SITES.inline,
      );
      if (point?.site.type !== COMMENT_RANGE_MARKER_SITES.inline) {
        result.kept.push(...markers);
        continue;
      }
      point.site.content[point.site.index] = { type: "commentReference", id };
      result.reanchored += 1;
    }
    const gone = new Set(unpaired);
    result.kept.push(...markers.filter((marker) => !gone.has(marker)));
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
    case COMMENT_RANGE_MARKER_SITES.positioned: {
      const { owner, child } = site;
      const children = owner.preserved?.children ?? [];
      const at = children.indexOf(child);
      if (at < 0) return;
      children.splice(at, 1);
      result.unpositioned += 1;
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
