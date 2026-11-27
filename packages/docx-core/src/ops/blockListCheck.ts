import type { BlockContent, Paragraph, Table } from "../model/document";
import { listKey, sourceSplice } from "./listSplice";

/** Whether a block list holds what a check looks for. */
export type BlockListCheck = (blocks: readonly BlockContent[]) => boolean;

type BlockListCheckOptions = {
  /** Whether a paragraph directly in the list is what the check looks for. */
  paragraph: (paragraph: Paragraph) => boolean;
  /** Whether the table itself is what the check looks for. */
  table?: (table: Table) => boolean;
  /** How a table's cells are checked; as the list itself by default. */
  cell?: BlockListCheck;
};

/**
 * Whether a block list or a container in it holds what the checks look for,
 * reusing unchanged lists: operation documents and their block lists are
 * immutable once validated.
 */
export const blockListCheck = ({
  paragraph,
  table,
  cell,
}: BlockListCheckOptions): BlockListCheck => {
  const checked = new WeakMap<object, boolean>();
  const holds: BlockListCheck = (blocks) => {
    const key = listKey(blocks);
    const cached = checked.get(key);
    if (cached !== undefined) return cached;
    const splice = sourceSplice(blocks);
    // A previously clean list can only acquire a violation in the replacement.
    const candidates =
      splice !== undefined && checked.get(splice.previous) === false
        ? blocks.slice(splice.start, splice.end)
        : blocks;
    const found = candidates.some((block) => {
      switch (block.type) {
        case "paragraph":
          return paragraph(block);
        case "table":
          return (
            table?.(block) === true ||
            block.rows.some((row) =>
              row.cells.some(({ content }) => (cell ?? holds)(content)),
            )
          );
        case "blockSdt":
        case "blockCustomXml":
          return holds(block.content);
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
          return false;
        default: {
          const unreachable: never = block;
          return unreachable;
        }
      }
    });
    checked.set(key, found);
    return found;
  };
  return holds;
};
