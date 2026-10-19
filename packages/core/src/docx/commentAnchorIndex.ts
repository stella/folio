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
import { visitInlineContentSlots } from "./paragraphTraversal";

export type CommentMarkerSlot = {
  /** The inline-content array the marker sits in, so a caller can splice it. */
  content: ParagraphContent[];
  index: number;
  item: Extract<
    ParagraphContent,
    { type: "commentRangeStart" } | { type: "commentRangeEnd" } | { type: "commentReference" }
  >;
};

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
        visit({ content, index, item });
      }
    });
  };

  const visitBlock = (block: BlockContent): void => {
    if (block.type === "paragraph") {
      visitParagraph(block);
      return;
    }
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          visitBlocks(cell.content);
        }
      }
      return;
    }
    if (block.type !== "blockSdt" && block.type !== "blockCustomXml") {
      // Captured bytes hold no comment boundary folio read, and a bookmark
      // marker anchors none. Only wrappers with parsed children recurse.
      return;
    }
    visitBlocks(block.content);
  };

  const visitBlocks = (nested: readonly BlockContent[]): void => {
    for (const block of nested) {
      visitBlock(block);
    }
  };

  visitBlocks(blocks);
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
    return children === hyperlink.children ? hyperlink : { ...hyperlink, children };
  };
  const rebuildSimpleFieldContent = (item: SimpleFieldContent): SimpleFieldContent => {
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
  const rebuildInlineSdtContent = (item: InlineSdtContent): InlineSdtContent => {
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
  const rebuildTrackedContent = (item: TrackedRunContent): TrackedRunContent => {
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
    change: Extract<ParagraphContent, { type: "insertion" | "deletion" | "moveFrom" | "moveTo" }>,
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
  const rebuildHyperlinkContent = (item: HyperlinkContent): HyperlinkContent => {
    switch (item.type) {
      case "inlineWrapper": {
        const content = rebuildParagraphContent(item.content);
        return content === item.content ? item : { ...item, content };
      }
      case "run":
      case "bookmarkStart":
      case "bookmarkEnd":
      case "preservedInline":
        return item;
    }
  };
  const rebuildParagraphContent = (content: ParagraphContent[]): ParagraphContent[] =>
    rebuildArray(content, rebuildParagraphItem);

  const rebuildParagraph = (paragraph: Paragraph): Paragraph => {
    const content = rebuildParagraphContent(paragraph.content);
    return content === paragraph.content ? paragraph : { ...paragraph, content };
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
        return block;
    }
  };
  const rebuildBlocks = (content: BlockContent[]): BlockContent[] =>
    rebuildArray(content, rebuildBlock);

  return rebuildBlocks(blocks.slice());
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
export const rangedCommentIds = (blocks: readonly BlockContent[]): ReadonlySet<number> => {
  const ids = new Set<number>();
  visitCommentMarkers(blocks, ({ item }) => {
    if (item.type === "commentRangeStart") {
      ids.add(item.id);
    }
  });
  return ids;
};
