/**
 * One traversal over the comment markers a story's blocks hold.
 *
 * Both sides of the editor round trip need it: the conversion has to know
 * which comments own a range before it decides a reference is a point comment,
 * and the save path has to know which comments still lack a reference. Two
 * walks would drift.
 */

import type {
  BlockContent,
  Hyperlink,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  SimpleField,
  Table,
  TableCell,
  TableRow,
  TrackedRunContent,
} from "../types/document";
import { cloneParagraphWithPropertySource } from "./paragraphPropertySource";
import { visitInlineContentSlots } from "./paragraphTraversal";

type CommentMarker = Extract<
  ParagraphContent,
  | { type: "commentRangeStart" }
  | { type: "commentRangeEnd" }
  | { type: "commentReference" }
>;

/** A marker inside a paragraph, in the inline-content array it sits in. */
export type InlineCommentMarkerSlot = {
  level: "inline";
  /** The inline-content array the marker sits in, so a caller can splice it. */
  content: ParagraphContent[];
  index: number;
  item: CommentMarker;
};

/** A range marker between blocks, in the block array it sits in. */
export type BlockCommentMarkerSlot = {
  level: "block";
  content: BlockContent[];
  index: number;
  item: Extract<
    BlockContent,
    { type: "commentRangeStart" } | { type: "commentRangeEnd" }
  >;
};

export type CommentMarkerSlot =
  InlineCommentMarkerSlot | BlockCommentMarkerSlot;

export const visitCommentMarkers = (
  blocks: readonly BlockContent[],
  visit: (slot: CommentMarkerSlot) => void,
): void => {
  const visitParagraph = (paragraph: Paragraph): void => {
    visitInlineContentSlots(paragraph, ({ content, index, item }) => {
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd" ||
        item.type === "commentReference"
      ) {
        visit({ level: "inline", content, index, item });
      }
    });
  };

  const visitBlock = (block: BlockContent): void => {
    switch (block.type) {
      case "paragraph":
        visitParagraph(block);
        return;
      case "table":
        for (const row of block.rows) {
          for (const cell of row.cells) {
            visitBlocks(cell.content);
          }
        }
        return;
      case "blockSdt":
      case "blockCustomXml":
        visitBlocks(block.content);
        return;
      // Reported by `visitBlocks`, which has the array a caller splices.
      case "commentRangeStart":
      case "commentRangeEnd":
        return;
      // Captured bytes hold no comment boundary folio read, and a bookmark or
      // move marker anchors none.
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

  const visitBlocks = (nested: readonly BlockContent[]): void => {
    for (const [index, block] of nested.entries()) {
      // A range may open or close between blocks: the marker is the story's
      // as much as one inside a paragraph is.
      if (
        block.type === "commentRangeStart" ||
        block.type === "commentRangeEnd"
      ) {
        // SAFETY: the walk never writes through `content`; callers splice it.
        visit({
          level: "block",
          content: nested as BlockContent[],
          index,
          item: block,
        });
        continue;
      }
      visitBlock(block);
    }
  };

  visitBlocks(blocks);
};

/** A paragraph's comment markers, in document order. */
export type ParagraphCommentMarkers = (
  paragraph: Paragraph,
) => readonly CommentMarker[];

/** A paragraph's comment markers, read from all of its content. */
export const paragraphCommentMarkers: ParagraphCommentMarkers = (paragraph) => {
  const markers: CommentMarker[] = [];
  visitInlineContentSlots(paragraph, ({ item }) => {
    if (
      item.type === "commentRangeStart" ||
      item.type === "commentRangeEnd" ||
      item.type === "commentReference"
    )
      markers.push(item);
  });
  return markers;
};

/**
 * Each paragraph's comment markers, read once per paragraph object. Only for
 * documents whose paragraphs are never changed in place, as a session's
 * versions are: an edit replaces the paragraphs it changes.
 */
export class CommentMarkerCache {
  private readonly known = new WeakMap<Paragraph, readonly CommentMarker[]>();

  readonly read: ParagraphCommentMarkers = (paragraph) => {
    const kept = this.known.get(paragraph);
    if (kept !== undefined) return kept;
    const markers = paragraphCommentMarkers(paragraph);
    this.known.set(paragraph, markers);
    return markers;
  };
}

/**
 * Every comment marker of a story's blocks in the order
 * {@link visitCommentMarkers} visits them, each paragraph's read through
 * `markersOf`.
 */
export const visitCommentMarkerItems = (
  blocks: readonly BlockContent[],
  markersOf: ParagraphCommentMarkers,
  visit: (item: CommentMarker | BlockCommentMarkerSlot["item"]) => void,
): void => {
  for (const block of blocks) {
    switch (block.type) {
      case "commentRangeStart":
      case "commentRangeEnd":
        visit(block);
        break;
      case "paragraph":
        for (const marker of markersOf(block)) visit(marker);
        break;
      case "table":
        for (const row of block.rows)
          for (const cell of row.cells)
            visitCommentMarkerItems(cell.content, markersOf, visit);
        break;
      case "blockSdt":
      case "blockCustomXml":
        visitCommentMarkerItems(block.content, markersOf, visit);
        break;
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        break;
      default: {
        const unhandled: never = block;
        return unhandled;
      }
    }
  }
};

/** Remove every serialized anchor record for comments whose definitions are retired. */
export const removeCommentMarkers = (
  blocks: readonly BlockContent[],
  commentIds: ReadonlySet<number>,
): BlockContent[] => {
  if (commentIds.size === 0) return blocks.slice();
  const removals = new Map<readonly unknown[], number[]>();
  visitCommentMarkers(blocks, ({ content, index, item }) => {
    if (!commentIds.has(item.id)) return;
    const indexes = removals.get(content) ?? [];
    indexes.push(index);
    removals.set(content, indexes);
  });
  if (removals.size === 0) return blocks.slice();

  const rebuildArray = <T>(content: T[], rebuildItem: (item: T) => T): T[] => {
    const removed = removals.get(content);
    let changed = removed !== undefined;
    const result: T[] = [];
    for (const [index, item] of content.entries()) {
      if (removed?.includes(index) === true) continue;
      const rebuilt = rebuildItem(item);
      changed ||= rebuilt !== item;
      result.push(rebuilt);
    }
    return changed ? result : content;
  };

  type SimpleFieldContent = SimpleField["content"][number];
  type InlineSdtContent = InlineSdt["content"][number];
  type HyperlinkContent = Hyperlink["children"][number];

  const rebuildHyperlink = (hyperlink: Hyperlink): Hyperlink => {
    const children = rebuildArray(hyperlink.children, rebuildHyperlinkContent);
    return children === hyperlink.children
      ? hyperlink
      : { ...hyperlink, children };
  };
  const rebuildSimpleFieldContent = (
    item: SimpleFieldContent,
  ): SimpleFieldContent => {
    switch (item.type) {
      case "hyperlink":
        return rebuildHyperlink(item);
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        return rebuildTrackedChange(item);
      case "inlineWrapper": {
        const content = rebuildParagraphContent(item.content);
        return content === item.content ? item : { ...item, content };
      }
      case "run":
      case "preservedInline":
        return item;
    }
  };
  const rebuildSimpleField = (field: SimpleField): SimpleField => {
    const content = rebuildArray(field.content, rebuildSimpleFieldContent);
    return content === field.content ? field : { ...field, content };
  };
  const rebuildInlineSdtContent = (
    item: InlineSdtContent,
  ): InlineSdtContent => {
    switch (item.type) {
      case "hyperlink":
        return rebuildHyperlink(item);
      case "simpleField":
        return rebuildSimpleField(item);
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        return rebuildTrackedChange(item);
      case "inlineSdt":
        return rebuildInlineSdt(item);
      case "inlineWrapper": {
        const content = rebuildParagraphContent(item.content);
        return content === item.content ? item : { ...item, content };
      }
      case "run":
      case "complexField":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "mathEquation":
      case "preservedInline":
        return item;
    }
  };
  const rebuildInlineSdt = (sdt: InlineSdt): InlineSdt => {
    const content = rebuildArray(sdt.content, rebuildInlineSdtContent);
    return content === sdt.content ? sdt : { ...sdt, content };
  };
  const rebuildTrackedContent = (
    item: TrackedRunContent,
  ): TrackedRunContent => {
    switch (item.type) {
      case "hyperlink":
        return rebuildHyperlink(item);
      case "simpleField":
        return rebuildSimpleField(item);
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        return rebuildTrackedChange(item);
      case "inlineSdt":
        return rebuildInlineSdt(item);
      case "inlineWrapper": {
        const content = rebuildParagraphContent(item.content);
        return content === item.content ? item : { ...item, content };
      }
      case "run":
      case "complexField":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "mathEquation":
      case "preservedInline":
        return item;
    }
  };
  const rebuildTrackedChange = (
    change: Extract<
      ParagraphContent,
      { type: "insertion" | "deletion" | "moveFrom" | "moveTo" }
    >,
  ): typeof change => {
    const content = rebuildArray(change.content, rebuildTrackedContent);
    return content === change.content ? change : { ...change, content };
  };
  const rebuildParagraphItem = (item: ParagraphContent): ParagraphContent => {
    switch (item.type) {
      case "hyperlink":
        return rebuildHyperlink(item);
      case "simpleField":
        return rebuildSimpleField(item);
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
        return rebuildTrackedChange(item);
      case "inlineSdt":
        return rebuildInlineSdt(item);
      case "inlineWrapper": {
        const content = rebuildParagraphContent(item.content);
        return content === item.content ? item : { ...item, content };
      }
      case "run":
      case "complexField":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "mathEquation":
      case "preservedInline":
        return item;
    }
  };
  const rebuildHyperlinkContent = (
    item: HyperlinkContent,
  ): HyperlinkContent => {
    switch (item.type) {
      case "inlineWrapper": {
        const content = rebuildParagraphContent(item.content);
        return content === item.content ? item : { ...item, content };
      }
      case "run":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "complexField":
      case "preservedInline":
        return item;
    }
  };
  const rebuildParagraphContent = (
    content: ParagraphContent[],
  ): ParagraphContent[] => rebuildArray(content, rebuildParagraphItem);

  const rebuildParagraph = (paragraph: Paragraph): Paragraph => {
    const content = rebuildParagraphContent(paragraph.content);
    return content === paragraph.content
      ? paragraph
      : cloneParagraphWithPropertySource(paragraph, { content });
  };
  const rebuildCell = (cell: TableCell): TableCell => {
    const content = rebuildBlocks(cell.content);
    return content === cell.content ? cell : { ...cell, content };
  };
  const rebuildRow = (row: TableRow): TableRow => {
    const cells = rebuildArray(row.cells, rebuildCell);
    return cells === row.cells ? row : { ...row, cells };
  };
  const rebuildTable = (table: Table): Table => {
    const rows = rebuildArray(table.rows, rebuildRow);
    return rows === table.rows ? table : { ...table, rows };
  };
  const rebuildBlock = (block: BlockContent): BlockContent => {
    switch (block.type) {
      case "paragraph":
        return rebuildParagraph(block);
      case "table":
        return rebuildTable(block);
      case "blockSdt":
      case "blockCustomXml": {
        const content = rebuildBlocks(block.content);
        return content === block.content ? block : { ...block, content };
      }
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        return block;
    }
  };
  const rebuildBlocks = (content: BlockContent[]): BlockContent[] =>
    rebuildArray(content, rebuildBlock);

  // The walk keyed markers between blocks by this very array, so it is
  // rebuilt itself rather than a copy of it.
  // SAFETY: rebuildArray never writes through the array it reads.
  const rebuilt = rebuildBlocks(blocks as BlockContent[]);
  return rebuilt === blocks ? blocks.slice() : rebuilt;
};

/**
 * Comments that open a range anywhere in these blocks.
 *
 * A `commentReference` whose id is absent has no highlight of its own — the
 * point-comment shape the model's `CommentReference` documents — and only that
 * one needs its mark anchored onto neighbouring text. Anchoring one whose range
 * exists stretches that range to wherever the reference sits, which is how a
 * comment on one word came back covering the rest of the sentence. The range
 * and the reference need not share a paragraph, so the question is asked of the
 * whole story: a reference that trails the end of a range three paragraphs back
 * is still not a point comment.
 */
export const rangedCommentIds = (
  blocks: readonly BlockContent[],
): ReadonlySet<number> => {
  const ids = new Set<number>();
  visitCommentMarkers(blocks, ({ item }) => {
    if (item.type === "commentRangeStart") {
      ids.add(item.id);
    }
  });
  return ids;
};
