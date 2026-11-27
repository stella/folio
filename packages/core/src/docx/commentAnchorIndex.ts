/**
 * One traversal over the comment markers a story's blocks hold.
 *
 * Both sides of the editor round trip need it: the conversion has to know
 * which comments own a range before it decides a reference is a point comment,
 * and the save path has to know which comments still lack a reference. Two
 * walks would drift.
 */

import type { PreservedMarkup } from "@stll/docx-core/model";
import {
  type PreservedCommentRangeMarker,
  preservedCommentRangeMarkersAt,
} from "@stll/docx-core/ops";

import type {
  BlockContent,
  Hyperlink,
  InlineSdt,
  Paragraph,
  ParagraphContent,
  ShapeContent,
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

/**
 * A range marker a table keeps between its rows, or a row between its cells,
 * as positioned markup. It is written back where it stood, so a caller reads
 * it as one of the range's markers but has nothing to splice.
 */
export type PositionedCommentMarkerSlot = {
  level: "positioned";
  item: PreservedCommentRangeMarker;
};

export type CommentMarkerSlot =
  | InlineCommentMarkerSlot
  | BlockCommentMarkerSlot
  | PositionedCommentMarkerSlot;

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

  const visitPositioned = (
    preserved: PreservedMarkup | undefined,
    gap: { at: number; childCount: number },
  ): void => {
    for (const item of preservedCommentRangeMarkersAt(
      preserved?.children,
      gap,
    )) {
      visit({ level: "positioned", item });
    }
  };

  const visitBlock = (block: BlockContent): void => {
    switch (block.type) {
      case "paragraph":
        visitParagraph(block);
        return;
      case "table":
        // In document order: a positioned marker stands before the row or
        // cell its index names.
        for (const [rowIndex, row] of block.rows.entries()) {
          visitPositioned(block.preserved, {
            at: rowIndex,
            childCount: block.rows.length,
          });
          for (const [cellIndex, cell] of row.cells.entries()) {
            visitPositioned(row.preserved, {
              at: cellIndex,
              childCount: row.cells.length,
            });
            visitBlocks(cell.content);
          }
          visitPositioned(row.preserved, {
            at: row.cells.length,
            childCount: row.cells.length,
          });
        }
        visitPositioned(block.preserved, {
          at: block.rows.length,
          childCount: block.rows.length,
        });
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

type CommentMarkerWalk = {
  markersOf: ParagraphCommentMarkers;
  visit: (item: CommentMarker, position: "content" | "positioned") => void;
  enterParagraph?: (paragraph: Paragraph) => void;
};

/** Every source comment marker in story order, with paragraph reads cached by the caller. */
export const visitCommentMarkerItems = (
  blocks: readonly BlockContent[],
  { markersOf, visit, enterParagraph }: CommentMarkerWalk,
): void => {
  const at = (
    preserved: PreservedMarkup | undefined,
    gap: { at: number; childCount: number },
  ): void => {
    for (const marker of preservedCommentRangeMarkersAt(
      preserved?.children,
      gap,
    ))
      visit(marker, "positioned");
  };
  const walk = (nested: readonly BlockContent[]): void => {
    for (const block of nested) {
      switch (block.type) {
        case "commentRangeStart":
        case "commentRangeEnd":
          visit(block, "content");
          break;
        case "paragraph":
          enterParagraph?.(block);
          for (const marker of markersOf(block)) visit(marker, "content");
          break;
        case "table":
          for (const [rowIndex, row] of block.rows.entries()) {
            at(block.preserved, {
              at: rowIndex,
              childCount: block.rows.length,
            });
            for (const [cellIndex, cell] of row.cells.entries()) {
              at(row.preserved, {
                at: cellIndex,
                childCount: row.cells.length,
              });
              walk(cell.content);
            }
            at(row.preserved, {
              at: row.cells.length,
              childCount: row.cells.length,
            });
          }
          at(block.preserved, {
            at: block.rows.length,
            childCount: block.rows.length,
          });
          break;
        case "blockSdt":
        case "blockCustomXml":
          walk(block.content);
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
  walk(blocks);
};

/** Gaps at which positioned table or row markers enter canonical paragraph traversal. */
export const positionedCommentGaps = (
  blocks: readonly BlockContent[],
  skippedContinuations: ReadonlySet<Paragraph>,
) => {
  const before = new Map<Paragraph, CommentMarker[]>();
  const after = new Map<Paragraph, CommentMarker[]>();
  for (const block of blocks) {
    if (
      block.type !== "table" &&
      block.type !== "blockSdt" &&
      block.type !== "blockCustomXml"
    )
      continue;
    let previous: Paragraph | undefined;
    let waiting: CommentMarker[] = [];
    visitCommentMarkerItems([block], {
      markersOf: () => [],
      enterParagraph(paragraph) {
        // A collapsed continuation has no native paragraph on which to pass
        // the following marker. Use the last rendered paragraph instead.
        if (skippedContinuations.has(paragraph)) return;
        if (waiting.length > 0) {
          before.set(paragraph, waiting);
          waiting = [];
        }
        previous = paragraph;
      },
      visit(marker, position) {
        if (position !== "positioned") return;
        if (previous === undefined) waiting.push(marker);
        else {
          const list = after.get(previous);
          if (list === undefined) after.set(previous, [marker]);
          else list.push(marker);
        }
      },
    });
  }
  return { before, after };
};

/** `removeCommentMarkers`, in the text boxes the blocks draw as well. */
export const removeCommentMarkersThroughTextBoxes = (
  blocks: readonly BlockContent[],
  commentIds: ReadonlySet<number>,
): BlockContent[] => {
  if (commentIds.size === 0) return blocks.slice();
  const strip = (value: unknown): unknown => {
    if (Array.isArray(value)) {
      let changed = false;
      const items = value.map((item: unknown) => {
        const next = strip(item);
        changed ||= next !== item;
        return next;
      });
      return changed ? items : value;
    }
    if (typeof value !== "object" || value === null) return value;
    const record = value as Record<string, unknown>;
    if (record["type"] === "shape") {
      const content = value as ShapeContent;
      const textBody = content.shape.textBody;
      if (textBody === undefined) return value;
      const stripped = removeCommentMarkersThroughTextBoxes(
        textBody.content,
        commentIds,
      );
      return {
        ...content,
        shape: {
          ...content.shape,
          textBody: { ...textBody, content: stripped },
        },
      };
    }
    let changed = false;
    const out: Record<string, unknown> = {};
    for (const [key, child] of Object.entries(record)) {
      const next = strip(child);
      changed ||= next !== child;
      out[key] = next;
    }
    return changed ? out : value;
  };
  return removeCommentMarkers(strip(blocks) as BlockContent[], commentIds);
};

/** Remove every serialized anchor record for comments whose definitions are retired. */
export const removeCommentMarkers = (
  blocks: readonly BlockContent[],
  commentIds: ReadonlySet<number>,
): BlockContent[] => {
  if (commentIds.size === 0) return blocks.slice();
  const removals = new Map<readonly unknown[], number[]>();
  visitCommentMarkers(blocks, (slot) => {
    // A table's or a row's own marker stays the positioned markup it was.
    if (slot.level === "positioned" || !commentIds.has(slot.item.id)) return;
    const indexes = removals.get(slot.content) ?? [];
    indexes.push(slot.index);
    removals.set(slot.content, indexes);
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

  type InlineSdtContent = InlineSdt["content"][number];
  type HyperlinkContent = Hyperlink["children"][number];

  const rebuildHyperlink = (hyperlink: Hyperlink): Hyperlink => {
    const children = rebuildArray(hyperlink.children, rebuildHyperlinkContent);
    return children === hyperlink.children
      ? hyperlink
      : { ...hyperlink, children };
  };
  const rebuildSimpleField = (field: SimpleField): SimpleField => {
    // A field holds what a tracked change holds.
    const content = rebuildArray(field.content, rebuildTrackedContent);
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
