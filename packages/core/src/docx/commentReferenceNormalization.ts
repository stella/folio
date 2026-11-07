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
import {
  findAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

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
  content: ParagraphContent[];
  index: number;
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
        rangeMarkers.push({ content, index, id: item.id, type: item.type });
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
    if (block.type === "paragraph") {
      normalizeParagraph(block);
      return;
    }
    if (block.type === "table") {
      normalizeTable(block);
      return;
    }
    if (block.type !== "blockSdt" && block.type !== "blockCustomXml") {
      return;
    }
    normalizeBlocks(block.content);
  };

  const normalizeBlocks = (blocks: BlockContent[]): void => {
    moveCommentBoundariesIntoText(blocks);
    for (const block of blocks) {
      normalizeBlock(block);
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
          reanchorCommentRangeMarker(marker);
          reanchoredCount += 1;
          continue;
        }
        openStart = marker;
        continue;
      }

      if (!openStart) {
        reanchorCommentRangeMarker(marker);
        reanchoredCount += 1;
        continue;
      }
      openStart = null;
    }

    if (openStart) {
      reanchorCommentRangeMarker(openStart);
      reanchoredCount += 1;
    }
  }
  return reanchoredCount;
};

const reanchorCommentRangeMarker = (marker: CommentRangeMarkerRef): void => {
  marker.content[marker.index] = {
    type: "commentReference",
    id: marker.id,
  };
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

type CapturedCommentBoundary = {
  type: "commentRangeStart" | "commentRangeEnd";
  id: number;
};

/** The comment range boundary a captured block is, if it is one. */
const capturedCommentBoundary = (
  block: BlockContent,
): CapturedCommentBoundary | null => {
  if (
    block.type !== "preservedBlock" ||
    !/^<[^>]*commentRange(?:Start|End)\b/u.test(block.xml)
  )
    return null;
  let element;
  try {
    // Captured bytes rely on the part's bindings, so they are read under them.
    element = parseXml(block.xml, OOXML_NAMESPACE_SCOPE).elements?.find(
      (child) => child.type === "element",
    );
  } catch {
    return null;
  }
  if (
    !element ||
    !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "")
  )
    return null;
  const name = getLocalName(element.name ?? "");
  const id = Number(
    findAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_URIS, "id")
      ?.value,
  );
  if (
    (name !== "commentRangeStart" && name !== "commentRangeEnd") ||
    !Number.isSafeInteger(id)
  )
    return null;
  return { type: name, id };
};

/** The first or last paragraph a block holds, looking inside its wrappers. */
const edgeParagraph = (
  block: BlockContent | undefined,
  edge: "first" | "last",
): Paragraph | undefined => {
  if (!block) return undefined;
  if (block.type === "paragraph") return block;
  let children: readonly BlockContent[] | undefined;
  if (block.type === "blockSdt" || block.type === "blockCustomXml") {
    children = block.content;
  } else if (block.type === "table") {
    const cell =
      edge === "first"
        ? block.rows.at(0)?.cells.at(0)
        : block.rows.at(-1)?.cells.at(-1);
    children = cell?.content;
  }
  if (!children) return undefined;
  const ordered = edge === "first" ? children : children.toReversed();
  for (const child of ordered) {
    const found = edgeParagraph(child, edge);
    if (found) return found;
  }
  return undefined;
};

/**
 * A comment range boundary written between blocks, beside a content control
 * (`w:displacedByCustomXml`) say, opens or closes the range at the text next
 * to it. The reader keeps it as captured bytes, which every comment reader
 * then misses: the range looks unbalanced and the comment turned into a point
 * comment. A start goes to the front of the next paragraph and an end to the
 * back of the one before, which bounds the same text, so the range is one
 * that every reader sees.
 */
const moveCommentBoundariesIntoText = (blocks: BlockContent[]): void => {
  for (let index = 0; index < blocks.length; index++) {
    const block = blocks[index];
    const boundary = block ? capturedCommentBoundary(block) : null;
    if (!boundary) continue;
    const marker = { type: boundary.type, id: boundary.id } as const;
    if (boundary.type === "commentRangeStart") {
      const next = blocks
        .slice(index + 1)
        .map((candidate) => edgeParagraph(candidate, "first"))
        .find((paragraph) => paragraph !== undefined);
      if (!next) continue;
      next.content.unshift(marker);
    } else {
      const previous = blocks
        .slice(0, index)
        .toReversed()
        .map((candidate) => edgeParagraph(candidate, "last"))
        .find((paragraph) => paragraph !== undefined);
      if (!previous) continue;
      previous.content.push(marker);
    }
    blocks.splice(index, 1);
    index -= 1;
  }
};
