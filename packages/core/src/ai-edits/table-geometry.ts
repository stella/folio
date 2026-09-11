import { TaggedError } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import {
  executeTableGeometryProgram,
  preflightTableGeometry,
  projectTableGeometry,
  type TableCellCoordinate,
  type TableGeometryExecutionIssue,
  type TableGeometryPairing,
  type TableGeometryUnsupportedIssue,
} from "../internal/compare/table-geometry-program";
import type { FolioStoryTable } from "./snapshot";

export { projectTableGeometry };
export type { TableCellCoordinate, TableGeometryPairing };

type MatchTableGeometryOptions = {
  readonly tr: Transaction;
  /** The base story's tables, with the positions the transaction will write at. */
  readonly baseTables: readonly FolioStoryTable[];
  /** The target story's tables, by the same index the pairings name. */
  readonly targetTables: ReadonlyMap<number, PMNode>;
  readonly pairings: readonly TableGeometryPairing[];
  readonly revision: { readonly author: string; readonly date: string; readonly idSeed: number };
};

export type MatchTableGeometryResult = {
  /** First revision id a following batch may allocate. */
  readonly nextRevisionId: number;
  /** Tables, rows and cells whose properties the match moved. */
  readonly matched: number;
};

class UnsupportedTableGeometryError extends TaggedError("UnsupportedTableGeometryError")<{
  message: string;
  issue: TableGeometryUnsupportedIssue | TableGeometryExecutionIssue;
}> {}

/**
 * Temporary composition for the existing reviewer seam. The comparison path
 * will instead compose geometry and content programs into one story
 * transaction, then this adapter can be deleted.
 */
export const matchTableGeometry = ({
  tr,
  baseTables,
  targetTables,
  pairings,
  revision,
}: MatchTableGeometryOptions): MatchTableGeometryResult => {
  const preflight = preflightTableGeometry({ baseTables, targetTables, pairings });
  if (preflight.status === "unsupported") {
    throw new UnsupportedTableGeometryError({
      message: "The paired table geometry cannot be represented exactly.",
      issue: preflight.issue,
    });
  }
  const execution = executeTableGeometryProgram({ tr, program: preflight.program, revision });
  if (execution.status === "unsupported") {
    throw new UnsupportedTableGeometryError({
      message: "The preflighted table geometry no longer matches the live transaction.",
      issue: execution.issue,
    });
  }
  return {
    nextRevisionId: execution.receipt.nextRevisionId,
    matched: execution.receipt.revisions.length,
  };
};
