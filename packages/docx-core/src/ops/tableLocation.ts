/** Shared addressing for row and whole-table operations. */
import { Result, panic } from "better-result";
import type { Document, Table } from "../model/document";
import { blockListAt, storyBody, storyParagraphs, type ParagraphLocation } from "./blocks";
import { idKey } from "./ids";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { DeleteRowOp, DeleteTableOp, InsertRowOp, SetTableRowsOp } from "./types";

type TableTarget = Pick<
  DeleteRowOp | DeleteTableOp | InsertRowOp | SetTableRowsOp,
  "type" | "story" | "blockId"
>;

export type TableRowLocation = {
  list: ParagraphLocation["list"];
  index: number;
  rowIndex: number;
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
  const stepIndex = location.list.findLastIndex((step) => step.kind === "tableCell");
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
  if (table?.type !== "table") return panic("A table-cell path must name a table.");
  return Result.ok({ list, index: step.block, rowIndex: step.row, table });
};
