/** Row resolution after inline review, before paragraph-mark joins. */
import { Result } from "better-result";

import type { BlockContent, Document, Table } from "../model/document";
import { storyBody } from "./blocks";
import { combineEdits, type DocumentEdit } from "./edits";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import { locateTableRow, tableRowAnchor } from "./tableLocation";
import {
  DOCUMENT_OP_TYPES,
  REVISION_DECISIONS,
  type OpStory,
  type ResolveRevisionOp,
} from "./types";

const collectTables = (blocks: readonly BlockContent[], out: Table[]): void => {
  for (const block of blocks) {
    switch (block.type) {
      case "table":
        for (const row of block.rows) {
          for (const cell of row.cells) {
            collectTables(cell.content, out);
          }
        }
        out.push(block);
        break;
      case "blockSdt":
      case "blockCustomXml":
        collectTables(block.content, out);
        break;
      case "paragraph":
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
};
const tablesIn = (document: Document, story: OpStory): Table[] => {
  const out: Table[] = [];
  collectTables(storyBody(document, story).content, out);
  return out;
};

/** Only row insert/delete records are reachable; cell and property revisions remain refused. */
export const reachableRowIds = (document: Document, story: OpStory): number[] =>
  tablesIn(document, story).flatMap((table) =>
    table.rows.flatMap((row) => {
      const change = row.structuralChange;
      return change?.type === "tableRowInsertion" || change?.type === "tableRowDeletion"
        ? [change.info.id]
        : [];
    }),
  );

type ResolveTableRowsOptions = {
  document: Document;
  op: ResolveRevisionOp;
  applyOps: ApplyOps;
};

export const resolveTableRows = ({
  document,
  op,
  applyOps,
}: ResolveTableRowsOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const ids = new Set(op.revisionIds);
  let current = document;
  const edits: DocumentEdit[] = [];
  // Tables are visited inside out. Re-read each containing table after an
  // inner edit so the exact precondition includes the resolved inner rows.
  for (const table of tablesIn(document, op.story)) {
    if (
      !table.rows.some(
        (row) => row.structuralChange !== undefined && ids.has(row.structuralChange.info.id),
      )
    ) {
      continue;
    }
    const blockId = tableRowAnchor(table.rows);
    if (blockId === undefined) {
      return Result.err(
        new DocumentOpRefusal({
          opType: op.type,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message: "The table has no paragraph that can address its rows.",
        }),
      );
    }
    const target = {
      type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
      story: op.story,
      blockId,
      expected: table.rows,
      rows: table.rows,
    } as const;
    const located = locateTableRow(current, target);
    if (located.isErr()) return Result.err(located.error);
    const expected = located.value.table.rows;
    const rows = expected.flatMap((row) => {
      const change = row.structuralChange;
      if (change === undefined || !ids.has(change.info.id)) return [row];
      const added = change.type === "tableRowInsertion";
      if (added !== (op.decision === REVISION_DECISIONS.ACCEPT)) return [];
      const next = { ...row };
      delete next.structuralChange;
      return [next];
    });
    const applied = applyOps(current, [
      { type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS, story: op.story, blockId, expected, rows },
    ]);
    if (applied.isErr()) {
      return Result.err(applied.error);
    }
    edits.push(applied.value);
    current = applied.value.document;
  }
  return Result.ok(combineEdits(document, edits));
};
