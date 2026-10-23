import type { FolioDocumentOperation } from "../document-operations";
import type { FolioAIBlock } from "./types";

type BlockIdentity = Pick<FolioAIBlock, "id"> & {
  table?: FolioAIBlock["table"];
};

const blockContainerKey = ({ table }: BlockIdentity): string =>
  table === undefined
    ? "body"
    : `${String(table.tableIndex)}:${String(table.rowIndex)}:${String(table.cellIndex)}`;

const isContainerCarrier = (
  blocks: readonly BlockIdentity[],
  anchorIndex: number,
): boolean => {
  const anchor = blocks[anchorIndex];
  if (anchor === undefined) return false;
  const containerKey = blockContainerKey(anchor);
  return blocks
    .slice(anchorIndex + 1)
    .every((block) => blockContainerKey(block) !== containerKey);
};

/**
 * The pre-batch block identities an operation removes from the accepted result.
 *
 * Batch operations all address this same block list. Keeping removal in block-id
 * coordinates prevents a structural operation from moving a later operation's
 * numeric position onto a different block.
 */
const removedBlockIds = (
  blocks: readonly BlockIdentity[],
  operation: FolioDocumentOperation,
): ReadonlySet<string> => {
  const removed = new Set<string>();
  const blockId =
    "range" in operation ? operation.range.blockId : operation.blockId;
  const anchorIndex = blocks.findIndex(({ id }) => id === blockId);
  const anchor = blocks[anchorIndex];

  switch (operation.type) {
    case "deleteBlock":
      // A deletion alone removes this target identity. A later tracked merge
      // may establish the container's final mark as its surviving carrier.
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
          if (block.table?.tableIndex === anchor.table.tableIndex)
            removed.add(block.id);
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
): ReadonlySet<string> => removedBlockIds(blocks, operation);

export const getTrackedFolioDocumentOperationRemovedBlockIds = (
  blocks: readonly BlockIdentity[],
  operation: FolioDocumentOperation,
): ReadonlySet<string> =>
  operation.type === "mergeBlockWithNext"
    ? new Set()
    : removedBlockIds(blocks, operation);

export const getTrackedFolioDocumentOperationCarrierBlockId = (
  blocks: readonly BlockIdentity[],
  operation: FolioDocumentOperation,
  previouslyRemovedBlockIds: { has(blockId: string): boolean },
): string | undefined => {
  if (operation.type !== "mergeBlockWithNext") return undefined;
  const anchorIndex = blocks.findIndex(({ id }) => id === operation.blockId);
  const anchor = blocks[anchorIndex];
  if (anchor === undefined) return undefined;
  const containerKey = blockContainerKey(anchor);
  let carrierIndex = anchorIndex + 1;
  while (
    blocks[carrierIndex + 1] !== undefined &&
    blockContainerKey(blocks[carrierIndex + 1]!) === containerKey
  ) {
    carrierIndex++;
  }
  const carrier = blocks[carrierIndex];
  if (carrier === undefined || !isContainerCarrier(blocks, carrierIndex))
    return undefined;
  return blocks
    .slice(anchorIndex + 1, carrierIndex + 1)
    .every(({ id }) => previouslyRemovedBlockIds.has(id))
    ? carrier.id
    : undefined;
};
