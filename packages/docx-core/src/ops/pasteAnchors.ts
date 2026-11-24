/**
 * Where pasted note references and comment markers land.
 *
 * A paste places its blocks through splits, joins and insertions; rather than
 * follow each one, every reference and marker is swapped for a uniquely named
 * bookmark before the paste and found again after it. Bookmarks are zero
 * width, so taking them out of the paste's operations moves no offset the
 * operations or the found places depend on.
 */
import type {
  Comment,
  Endnote,
  Footnote,
  Paragraph,
  ParagraphContent,
  RunContent,
  Table,
} from "../model/document";
import {
  childrenOf,
  contentWidth,
  isInlineContainer,
  withChildren,
} from "./offsets";

/** The notes and comments pasted references and markers name, by those ids. */
export type PasteAnnotations = {
  notes: readonly (Footnote | Endnote)[];
  comments: readonly Comment[];
};

/** What a placeholder stood for. */
export type PasteAnchor =
  | { kind: "note"; note: "footnote" | "endnote"; id: number }
  | {
      kind: "commentRangeStart" | "commentRangeEnd" | "commentReference";
      id: number;
    };

const PLACEHOLDER_PREFIX = "_folioPasteAnchor";
/** Above any id a package writes, so a placeholder never meets a bookmark. */
const PLACEHOLDER_FIRST_ID = 0x7f_00_00_00;

const placeholderIndex = (name: string): number | undefined => {
  if (!name.startsWith(PLACEHOLDER_PREFIX)) return undefined;
  const index = Number(name.slice(PLACEHOLDER_PREFIX.length));
  return Number.isInteger(index) ? index : undefined;
};

/** `value` with every placeholder bookmark in it removed. */
const withoutPasteAnchors = <T>(value: T, ids: ReadonlySet<number>): T => {
  const isPlaceholder = (item: unknown): boolean => {
    if (typeof item !== "object" || item === null) return false;
    const type: unknown = Reflect.get(item, "type");
    const id: unknown = Reflect.get(item, "id");
    return (
      (type === "bookmarkStart" || type === "bookmarkEnd") &&
      typeof id === "number" &&
      ids.has(id)
    );
  };
  const strip = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (let index = node.length - 1; index >= 0; index -= 1)
        if (isPlaceholder(node[index])) node.splice(index, 1);
      for (const item of node) strip(item);
      return;
    }
    if (typeof node !== "object" || node === null) return;
    for (const field of Object.values(node)) strip(field);
  };
  const copy = structuredClone(value);
  strip(copy);
  return copy;
};

type AnchoredBlocks = {
  /** The blocks with a placeholder for each reference and marker. */
  blocks: (Paragraph | Table)[];
  /** The blocks without the references and markers. */
  plain: (Paragraph | Table)[];
  anchors: PasteAnchor[];
};

/** The blocks with each note reference and comment marker a placeholder. */
export const withPasteAnchors = (
  blocks: readonly (Paragraph | Table)[],
): AnchoredBlocks => {
  const anchors: PasteAnchor[] = [];
  const placeholder = (anchor: PasteAnchor): ParagraphContent[] => {
    const index = anchors.length;
    anchors.push(anchor);
    const id = PLACEHOLDER_FIRST_ID + index;
    return [
      { type: "bookmarkStart", id, name: `${PLACEHOLDER_PREFIX}${index}` },
      { type: "bookmarkEnd", id },
    ];
  };
  const inline = (content: readonly ParagraphContent[]): ParagraphContent[] => {
    const out: ParagraphContent[] = [];
    for (const item of content) {
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd" ||
        item.type === "commentReference"
      ) {
        out.push(...placeholder({ kind: item.type, id: item.id }));
        continue;
      }
      if (item.type === "run") {
        if (
          !item.content.some(
            (piece) =>
              piece.type === "footnoteRef" || piece.type === "endnoteRef",
          )
        ) {
          out.push(item);
          continue;
        }
        let pending: RunContent[] = [];
        const flush = () => {
          if (pending.length > 0) out.push({ ...item, content: pending });
          pending = [];
        };
        for (const piece of item.content) {
          if (piece.type !== "footnoteRef" && piece.type !== "endnoteRef") {
            pending.push(piece);
            continue;
          }
          flush();
          out.push(
            ...placeholder({
              kind: "note",
              note: piece.type === "footnoteRef" ? "footnote" : "endnote",
              id: piece.id,
            }),
          );
        }
        flush();
        continue;
      }
      out.push(
        isInlineContainer(item)
          ? withChildren(item, inline(childrenOf(item)))
          : item,
      );
    }
    return out;
  };
  const block = (each: Paragraph | Table): Paragraph | Table =>
    each.type === "paragraph"
      ? { ...each, content: inline(each.content) }
      : {
          ...each,
          rows: each.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: cell.content.map((inner) =>
                inner.type === "paragraph" || inner.type === "table"
                  ? block(inner)
                  : inner,
              ),
            })),
          })),
        };
  const anchoredBlocks = blocks.map(block);
  const ids = new Set(anchors.map((_, index) => PLACEHOLDER_FIRST_ID + index));
  return {
    blocks: anchoredBlocks,
    plain: withoutPasteAnchors(anchoredBlocks, ids),
    anchors,
  };
};

/** Whether `paragraph` holds placeholders and nothing else. */
export const holdsOnlyPasteAnchors = (paragraph: Paragraph): boolean =>
  paragraph.content.length > 0 &&
  paragraph.content.every(
    (item) =>
      (item.type === "bookmarkStart" || item.type === "bookmarkEnd") &&
      item.id >= PLACEHOLDER_FIRST_ID,
  );

/** Where a placeholder landed: its paragraph, offset and paragraph order. */
export type PlacedAnchor = { blockId: string; offset: number; order: number };

/** Each placeholder's place among `paragraphs`, by anchor index. */
export const placedPasteAnchors = (
  paragraphs: readonly Paragraph[],
): Map<number, PlacedAnchor> => {
  const placed = new Map<number, PlacedAnchor>();
  for (const [order, paragraph] of paragraphs.entries()) {
    const blockId = paragraph.paraId;
    if (blockId === undefined) continue;
    const visit = (items: readonly ParagraphContent[], start: number) => {
      let offset = start;
      for (const item of items) {
        if (item.type === "bookmarkStart") {
          const index = placeholderIndex(item.name);
          if (index !== undefined)
            placed.set(index, { blockId, offset, order });
        }
        if (isInlineContainer(item)) visit(childrenOf(item), offset);
        offset += contentWidth([item]);
      }
    };
    visit(paragraph.content, 0);
  }
  return placed;
};
