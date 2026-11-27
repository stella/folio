/** Shared addressing for row and whole-table operations. */
import { panic, Result } from "better-result";

import type {
  BlockContent,
  Document,
  DocumentBody,
  Table,
  TableRow,
} from "../model/document";
import {
  blockListAt,
  type ParagraphLocation,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { idKey } from "./ids";
import { itemBefore } from "./listItems";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type {
  DeleteRowOp,
  DeleteTableOp,
  InsertBlocksOp,
  InsertRowOp,
  InsertTableOp,
  OpStory,
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

export type TableRowLocation = {
  list: ParagraphLocation["list"];
  index: number;
  rowIndex: number;
  cellIndex: number;
  table: Table;
};

/** A paragraph beside a table or in an enclosing transparent wrapper's list. */
export const survivingTableAnchor = (
  body: DocumentBody,
  list: ParagraphLocation["list"],
): ParagraphLocation | undefined => {
  let current = list;
  for (;;) {
    const blocks = blockListAt(body.content, current);
    const index = blocks.findLastIndex(
      (block) => block.type === "paragraph" && block.paraId !== undefined,
    );
    const paragraph = blocks[index];
    if (paragraph?.type === "paragraph")
      return { list: current, index, paragraph };
    const step = current.at(-1);
    if (step === undefined || step.kind === "tableCell") return undefined;
    current = current.slice(0, -1);
  }
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
