import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

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
};

type CommentRangeMarkerRef = {
  /** Absent for a marker between blocks, which a reference cannot replace. */
  slot?: { content: ParagraphContent[]; index: number };
  id: number;
  type: "commentRangeStart" | "commentRangeEnd";
};

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
  const dangling = new InlineContentRemovals();
  let removedDanglingReferences = 0;

  // Reading, reanchoring and removing run as three passes because a removal
  // shifts every later index in its array: the indices collected here stay
  // valid until `dangling.apply()` below.
  const normalizeParagraph = (paragraph: Paragraph): void => {
    visitInlineContentSlots(paragraph, ({ content, index, item }) => {
      if (isCommentMarker(item) && !validCommentIds.has(item.id)) {
        dangling.mark({ content, index });
        removedDanglingReferences += 1;
        return;
      }
      if (isCommentRangeMarker(item)) {
        rangeMarkers.push({
          slot: { content, index },
          id: item.id,
          type: item.type,
        });
      }
    });
  };

  const normalizeTable = (table: Table): void => {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        normalizeBlocks(cell.content);
      }
    }
  };

  const normalizeBlock = (block: BlockContent): void => {
    switch (block.type) {
      // A range may open or close between blocks; it pairs with the marker
      // inside a paragraph across the block edge.
      case "commentRangeStart":
      case "commentRangeEnd":
        if (validCommentIds.has(block.id)) {
          rangeMarkers.push({ id: block.id, type: block.type });
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
    for (const block of blocks) {
      normalizeBlock(block);
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
        blocks[index] = {
          type: "preservedBlock",
          xml: `<w:${block.type} w:id="${String(block.id)}"${
            block.displacedByCustomXml === undefined
              ? ""
              : ` w:displacedByCustomXml="${block.displacedByCustomXml}"`
          }/>`,
        };
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

  const reanchoredUnbalancedRanges =
    reanchorUnbalancedCommentRanges(rangeMarkers);
  dangling.apply();

  return { removedDanglingReferences, reanchoredUnbalancedRanges };
};

const reanchorUnbalancedCommentRanges = (
  rangeMarkers: readonly CommentRangeMarkerRef[],
): number => {
  const byId = new Map<number, CommentRangeMarkerRef[]>();
  for (const marker of rangeMarkers) {
    const markers = byId.get(marker.id);
    if (markers) {
      markers.push(marker);
      continue;
    }
    byId.set(marker.id, [marker]);
  }

  let reanchoredCount = 0;
  for (const markers of byId.values()) {
    let openStart: CommentRangeMarkerRef | null = null;
    for (const marker of markers) {
      if (marker.type === "commentRangeStart") {
        if (openStart) {
          if (reanchorCommentRangeMarker(marker)) reanchoredCount += 1;
          continue;
        }
        openStart = marker;
        continue;
      }

      if (!openStart) {
        if (reanchorCommentRangeMarker(marker)) reanchoredCount += 1;
        continue;
      }
      openStart = null;
    }

    if (openStart) {
      if (reanchorCommentRangeMarker(openStart)) reanchoredCount += 1;
    }
  }
  return reanchoredCount;
};

const reanchorCommentRangeMarker = (marker: CommentRangeMarkerRef): boolean => {
  // A marker between blocks stays as written: no run there can hold a reference.
  if (!marker.slot) return false;
  marker.slot.content[marker.slot.index] = {
    type: "commentReference",
    id: marker.id,
  };
  return true;
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
