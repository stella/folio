/** Shared addressing for row and whole-table operations. */
import { panic, Result } from "better-result";

import type {
  BlockContent,
  Document,
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
  InsertBlocksOp,
  "type" | "story" | "at"
>;

/** A paragraph or table side, in the block list containing that block. */
export const locateBlockInsertion = (
  document: Document,
  op: LocateBlockInsertionOptions,
): Result<Pick<ParagraphLocation, "list" | "index">, DocumentOpRefusal> => {
  switch (op.at.type) {
    case "before":
    case "after": {
      const paragraph = storyParagraphs(storyBody(document, op.story)).find(
        ({ paragraph: item }) =>
          idKey(item.paraId ?? "") === idKey(op.at.blockId),
      );
      if (paragraph === undefined)
        return Result.err(
          new DocumentOpRefusal({
            opType: op.type,
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "The insertion anchor does not exist.",
          }),
        );
      return Result.ok({
        list: paragraph.list,
        index: paragraph.index + (op.at.type === "after" ? 1 : 0),
      });
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
  const previous = index > 0 ? blocks.at(index - 1) : undefined;
  if (previous?.type === "paragraph" && previous.paraId !== undefined)
    return { type: "after", blockId: previous.paraId };
  if (previous?.type === "table") {
    const blockId = tableRowAnchor(previous.rows);
    if (blockId !== undefined) return { type: "afterTable", blockId };
  }
  return undefined;
};
