import type { FolioDocumentOperation } from "../document-operations";
import type { FolioAIBlock } from "./types";

type BlockIdentity = Pick<FolioAIBlock, "id"> & { table?: FolioAIBlock["table"] };

/**
 * The pre-batch block identities an operation removes from the accepted result.
 *
 * Batch operations all address this same block list. Keeping removal in block-id
 * coordinates prevents a structural operation from moving a later operation's
 * numeric position onto a different block.
 */
export const getFolioDocumentOperationRemovedBlockIds = (
  blocks: readonly BlockIdentity[],
  operation: FolioDocumentOperation,
): ReadonlySet<string> => {
  const removed = new Set<string>();
  const blockId = "range" in operation ? operation.range.blockId : operation.blockId;
  const anchorIndex = blocks.findIndex(({ id }) => id === blockId);
  const anchor = blocks[anchorIndex];

  switch (operation.type) {
    case "deleteBlock":
      removed.add(operation.blockId);
      break;
    case "mergeBlockWithNext": {
      // An unknown anchor removes nothing; index -1 must not select the first block.
      const next = anchorIndex < 0 ? undefined : blocks[anchorIndex + 1];
      if (next !== undefined) removed.add(next.id);
      break;
    }
    case "deleteTable":
      if (anchor?.table !== undefined) {
        for (const block of blocks) {
          if (block.table?.tableIndex === anchor.table.tableIndex) removed.add(block.id);
        }
      }
      break;
    case "deleteTableRow":
      if (anchor?.table !== undefined) {
        for (const block of blocks) {
          if (
            block.table?.tableIndex === anchor.table.tableIndex &&
            block.table.rowIndex === anchor.table.rowIndex
          ) {
            removed.add(block.id);
          }
        }
      }
      break;
    case "deleteTableColumn":
      if (anchor?.table !== undefined) {
        for (const block of blocks) {
          const table = block.table;
          if (
            table?.tableIndex === anchor.table.tableIndex &&
            table.gridColumnIndex === anchor.table.gridColumnIndex &&
            table.columnSpan === 1
          ) {
            removed.add(block.id);
          }
        }
      }
      break;
    default:
      break;
  }
  return removed;
};
