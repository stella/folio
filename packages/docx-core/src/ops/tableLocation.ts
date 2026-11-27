/** Shared addressing for row and whole-table operations. */
import { panic, Result } from "better-result";

import type {
  BlockContent,
  Document,
  DocumentBody,
  Paragraph,
  Table,
  TableRow,
} from "../model/document";
import {
  type BlockListAnchor,
  blockListAt,
  type ParagraphLocation,
  storyBody,
  storyParagraphs,
  survivingBlockListAnchor,
  updateBlockList,
} from "./blocks";
import { idKey } from "./ids";
import { itemBefore } from "./listItems";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "./types";
import type {
  DeleteRowOp,
  DeleteTableOp,
  InsertBlocksOp,
  InsertRowOp,
  InsertTableOp,
  OpStory,
  SetContainerBlocksOp,
  SetTableOp,
  SetTableRowsOp,
  TableEditOp,
  TableInsertionPoint,
  TableIntentOperation,
} from "./types";

type TableTarget =
  | Pick<
      | DeleteRowOp
      | DeleteTableOp
      | InsertRowOp
      | SetTableRowsOp
      | TableEditOp
      | TableIntentOperation
      | SetTableOp,
      "type" | "story" | "blockId"
    >
  // A table inserted beside the table holding `blockId`.
  | {
      type: InsertTableOp["type"] | InsertBlocksOp["type"];
      story: OpStory;
      blockId: string;
    };

type TableBlockLocation = {
  list: ParagraphLocation["list"];
  index: number;
};

export type TableRowLocation = TableBlockLocation & {
  rowIndex: number;
  cellIndex: number;
  table: Table;
};

type TableRemovalAddressOptions = {
  body: DocumentBody;
  story: OpStory;
  location: TableRowLocation;
};
type TableRemovalAddress = { anchor: BlockListAnchor; terminal?: Paragraph };

/** The stable paragraph a table's removal can address in the surrounding tree. */
export const tableRemovalAddress = ({
  body,
  story,
  location: { list, index, table },
}: TableRemovalAddressOptions): TableRemovalAddress | undefined => {
  const remaining = blockListAt(body.content, list).filter(
    (_, position) => position !== index,
  );
  // Keep the first direct paragraph used by existing structural inverses.
  const direct = remaining.find((block) => block.type === "paragraph")?.paraId;
  if (direct !== undefined) return { anchor: { list, blockId: direct } };
  if (
    remaining.length === 0 &&
    list.length === 0 &&
    story === OP_STORIES.MAIN
  ) {
    // A table-only story leaves its final cell's paragraph identity on the empty story paragraph.
    const paraId = storyParagraphs({ content: [table] }).at(-1)?.paragraph
      .paraId;
    if (paraId !== undefined)
      return {
        anchor: { list, blockId: paraId },
        terminal: { type: "paragraph", paraId, content: [] },
      };
  }
  const content = updateBlockList(body.content, list, () => remaining);
  const anchor = survivingBlockListAnchor({ ...body, content }, list);
  return anchor === undefined ? undefined : { anchor };
};

/** Find a table that may have no paragraph anchor, such as an omitted row. */
export const tableBlockLocation = (
  body: DocumentBody,
  target: Table,
): TableBlockLocation | undefined => {
  const visit = (
    blocks: readonly BlockContent[],
    list: ParagraphLocation["list"],
  ): TableBlockLocation | undefined => {
    for (const [index, block] of blocks.entries()) {
      if (block === target) return { list, index };
      if (block.type === "table") {
        for (const [row, { cells }] of block.rows.entries()) {
          for (const [cell, item] of cells.entries()) {
            const result = visit(item.content, [
              ...list,
              { kind: "tableCell", block: index, row, cell },
            ]);
            if (result !== undefined) return result;
          }
        }
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml") {
        const result = visit(block.content, [
          ...list,
          { kind: block.type, block: index },
        ]);
        if (result !== undefined) return result;
      }
    }
    return undefined;
  };
  return visit(body.content, []);
};

type TableContainerEditOptions = {
  body: DocumentBody;
  story: OpStory;
  location: TableBlockLocation;
  table: Table | undefined;
};

/** Address a table replacement through a surviving surrounding list. */
export const tableContainerEdit = ({
  body,
  story,
  location: { list, index },
  table,
}: TableContainerEditOptions): SetContainerBlocksOp | undefined => {
  const content = updateBlockList(body.content, list, (blocks) => {
    const next = [...blocks];
    if (table === undefined) next.splice(index, 1);
    else next[index] = table;
    return next;
  });
  const anchor = survivingBlockListAnchor({ content }, list);
  if (anchor === undefined) return undefined;
  return {
    type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
    story,
    blockId: anchor.blockId,
    ...(anchor.ancestorDepth === undefined
      ? {}
      : { ancestorDepth: anchor.ancestorDepth }),
    expected: blockListAt(body.content, anchor.list),
    blocks: blockListAt(content, anchor.list),
  };
};

/** A paragraph selects its innermost table, even through block wrappers. */
export const locateTableRow = (
  document: Document,
  op: TableTarget,
): Result<TableRowLocation, DocumentOpRefusal> => {
  const body = storyBody(document, op.story);
  const location = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(op.blockId),
  );
  if (location === undefined) {
    return Result.err(
      new DocumentOpRefusal({
        opType: op.type,
        reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        message: `No paragraph is ${op.blockId}.`,
      }),
    );
  }
  const stepIndex = location.list.findLastIndex(
    (step) => step.kind === "tableCell",
  );
  const step = location.list[stepIndex];
  if (step?.kind !== "tableCell") {
    return Result.err(
      new DocumentOpRefusal({
        opType: op.type,
        reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        message: "The paragraph is not in a table row.",
      }),
    );
  }
  const list = location.list.slice(0, stepIndex);
  const table = blockListAt(body.content, list)[step.block];
  if (table?.type !== "table")
    return panic("A table-cell path must name a table.");
  return Result.ok({
    list,
    index: step.block,
    rowIndex: step.row,
    cellIndex: step.cell,
    table,
  });
};

/** Row operations address this table, never a nested table appearing first. */
export const tableRowAnchor = (rows: readonly TableRow[]): string | undefined =>
  storyParagraphs({ content: [{ type: "table", rows: [...rows] }] }).find(
    ({ list }) => list.filter((step) => step.kind === "tableCell").length === 1,
  )?.paragraph.paraId;

type LocateBlockInsertionOptions = Pick<
  InsertBlocksOp | InsertTableOp,
  "type" | "story" | "at"
>;

/**
 * A paragraph or table side, or the start of the block list holding a
 * paragraph, in the block list containing that block.
 */
export const locateBlockInsertion = (
  document: Document,
  op: LocateBlockInsertionOptions,
): Result<Pick<ParagraphLocation, "list" | "index">, DocumentOpRefusal> => {
  switch (op.at.type) {
    case "before":
    case "after":
    case "containerStart": {
      const { at } = op;
      const paragraph = storyParagraphs(storyBody(document, op.story)).find(
        ({ paragraph: item }) => idKey(item.paraId ?? "") === idKey(at.blockId),
      );
      if (paragraph === undefined)
        return Result.err(
          new DocumentOpRefusal({
            opType: op.type,
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "The insertion anchor does not exist.",
          }),
        );
      const index = {
        before: paragraph.index,
        after: paragraph.index + 1,
        containerStart: 0,
      } as const satisfies Record<typeof at.type, number>;
      return Result.ok({ list: paragraph.list, index: index[at.type] });
    }
    case "beforeTable":
    case "afterTable": {
      const table = locateTableRow(document, {
        type: op.type,
        story: op.story,
        blockId: op.at.blockId,
      });
      if (table.isErr()) return table;
      return Result.ok({
        list: table.value.list,
        index: table.value.index + (op.at.type === "afterTable" ? 1 : 0),
      });
    }
    default: {
      const unreachable: never = op.at;
      return unreachable;
    }
  }
};

type BlockInsertionPointAtOptions = {
  blocks: readonly BlockContent[];
  index: number;
};

/** Restore a removed paragraph run beside its immediate surviving block. */
export const blockInsertionPointAt = ({
  blocks,
  index,
}: BlockInsertionPointAtOptions): TableInsertionPoint | undefined => {
  const following = blocks.at(index);
  if (following?.type === "paragraph" && following.paraId !== undefined)
    return { type: "before", blockId: following.paraId };
  if (following?.type === "table") {
    const blockId = tableRowAnchor(following.rows);
    if (blockId !== undefined) return { type: "beforeTable", blockId };
  }
  const previous = itemBefore(blocks, index);
  if (previous?.type === "paragraph" && previous.paraId !== undefined)
    return { type: "after", blockId: previous.paraId };
  if (previous?.type === "table") {
    const blockId = tableRowAnchor(previous.rows);
    if (blockId !== undefined) return { type: "afterTable", blockId };
  }
  // At the list's start, before a block that holds no paragraph to name.
  if (index !== 0) return undefined;
  for (const block of blocks)
    if (block.type === "paragraph" && block.paraId !== undefined)
      return { type: "containerStart", blockId: block.paraId };
  return undefined;
};
