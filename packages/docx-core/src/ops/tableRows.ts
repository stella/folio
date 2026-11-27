import { panic, Result } from "better-result";

import type {
  Document,
  DocumentBody,
  Paragraph,
  Table,
  TableRow,
} from "../model/document";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import {
  endsItsContainer,
  storyBody,
  storyParagraphs,
  updateBlockList,
  withBodyContent,
} from "./blocks";
import { validateOpsDocument } from "./contract";
import type { DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import {
  collides,
  countIds,
  countKeys,
  identityKeysIn,
  idKey,
  isParaId,
  packageIdentityKeys,
  packageParagraphIds,
  paragraphIdsIn,
} from "./ids";
import { textsIn } from "./leaves";
/** Row edits and their exact inverse, addressed through a paragraph in the row. */
import { cloneModel } from "./modelClone";
import {
  holdsPositionedMarkup,
  relocatePositionedMarkers,
} from "./positionedMarkers";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { WRAP_KINDS } from "./review";
import {
  isOwnRecord,
  recordedOnlyBy,
  type RevisionOwner,
} from "./revisionOwner";
import { replaceStoryBody } from "./stories";
import {
  locateTableRow,
  tableRowAnchor,
  type TableRowLocation,
} from "./tableLocation";
import {
  mergeOwnersContinuedBelow,
  removalOrphansAMerge,
  rowsWithoutRow,
  rowsWithRow,
} from "./tableRowRemoval";
import { applyTableOp } from "./tables";
import { permitsRestoredCellFinalMark, trackTableRows } from "./tableTracking";
import {
  type DeleteRowOp,
  DOCUMENT_OP_TYPES,
  type InsertRowOp,
  type OpStory,
  type SetTableRowsOp,
} from "./types";

type RowOp = InsertRowOp | DeleteRowOp | SetTableRowsOp;
type RowRefusalOptions = {
  op: RowOp;
  reason: DocumentOpRefusal["reason"];
  message: string;
};
const refused = ({ op, reason, message }: RowRefusalOptions) =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));

const paragraphsIn = (rows: readonly TableRow[]): Paragraph[] =>
  storyParagraphs({ content: [{ type: "table", rows: [...rows] }] }).map(
    ({ paragraph }) => paragraph,
  );

/**
 * Shared row wrappers need a structural table operation. Indexed markup
 * beside rows moves with them ({@link relocatePositionedMarkers}).
 */
const needsTableEdit = (table: Table): boolean =>
  table.rows.some(
    (row) =>
      row.contentControls !== undefined || row.carrierStack !== undefined,
  );

type WithTableOptions = {
  document: Document;
  story: OpStory;
  location: TableRowLocation;
  table: Table;
};
const withTable = ({
  document,
  story,
  location,
  table,
}: WithTableOptions): Document => {
  const body = storyBody(document, story);
  const content = updateBlockList(body.content, location.list, (blocks) => {
    const out = [...blocks];
    out[location.index] = table;
    return out;
  });
  return replaceStoryBody({
    document,
    story,
    body: withBodyContent(body, content),
  });
};

/** Commit a row list after checking the resulting package, before it is trusted. */
type CommitRowsOptions = {
  document: Document;
  op: RowOp;
  location: TableRowLocation;
  rows: TableRow[];
};
const commitRows = ({
  document,
  op,
  location,
  rows,
}: CommitRowsOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  if (rows.length === 0) {
    return applyTableOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: op.story,
      blockId: op.blockId,
      expected: location.table,
    }).mapError(
      (error) =>
        new DocumentOpRefusal({
          opType: op.type,
          reason: error.reason,
          message: error.message,
        }),
    );
  }
  const before = location.table.rows;
  if (structurallyEqual(before, rows)) {
    return Result.ok({
      document,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  }
  const afterParagraphs = paragraphsIn(rows);
  if (
    rows.some(
      (row) =>
        row.cells.length === 0 ||
        row.cells.some(
          (cell) => storyParagraphs({ content: cell.content }).length === 0,
        ),
    )
  ) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message: "Every row has cells, and every cell has a paragraph.",
    });
  }
  if (
    op.type !== DOCUMENT_OP_TYPES.SET_TABLE_ROWS &&
    removalOrphansAMerge(location.table, rows)
  ) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message:
        "A row edit keeps every vertical merge under the cell that starts it.",
    });
  }
  if (
    afterParagraphs.some(({ content }) =>
      textsIn(content).some(hasIllegalXmlCharacters),
    )
  ) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      message: "The row holds text that cannot be written to XML.",
    });
  }
  const body = { content: [{ type: "table", rows }] } satisfies DocumentBody;
  const originalParagraphs = new Map(
    paragraphsIn(before).map((paragraph) => [
      idKey(paragraph.paraId ?? ""),
      paragraph,
    ]),
  );
  for (const paragraphLocation of storyParagraphs(body)) {
    const mark = paragraphLocation.paragraph.pPrMark;
    if (
      mark !== undefined &&
      endsItsContainer(body, paragraphLocation) &&
      !permitsRestoredCellFinalMark({
        body,
        location: paragraphLocation,
        opType: op.type,
      }) &&
      !structurallyEqual(
        mark,
        originalParagraphs.get(idKey(paragraphLocation.paragraph.paraId ?? ""))
          ?.pPrMark,
      )
    ) {
      return refused({
        op: op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        message:
          "A cell-final mark must belong to its row structural revision.",
      });
    }
  }
  const beforeIds = new Set(paragraphIdsIn(before).map(idKey));
  if (
    afterParagraphs.some(
      ({ paraId }) =>
        paraId === undefined ||
        (!beforeIds.has(idKey(paraId)) && !isParaId(paraId)),
    )
  ) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      message: "An added paragraph has no usable id.",
    });
  }
  const remainingParagraphs = countIds(packageParagraphIds(document.package));
  for (const id of paragraphIdsIn(before)) {
    remainingParagraphs.set(
      idKey(id),
      (remainingParagraphs.get(idKey(id)) ?? 1) - 1,
    );
  }
  if (collides(remainingParagraphs, paragraphIdsIn(rows))) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      message: "A row paragraph id is already used in the package.",
    });
  }
  const remainingRecords = countKeys(packageIdentityKeys(document.package));
  for (const key of identityKeysIn(before))
    remainingRecords.set(key, (remainingRecords.get(key) ?? 1) - 1);
  const incomingRecords = identityKeysIn(rows);
  if (
    incomingRecords.some((key) => (remainingRecords.get(key) ?? 0) > 0) ||
    new Set(incomingRecords).size !== incomingRecords.length
  ) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      message:
        "A row revision or content-control id is already used in the package.",
    });
  }
  const table = relocatePositionedMarkers(location.table, {
    ...location.table,
    rows,
  });
  const next = withTable({
    document: document,
    story: op.story,
    location: location,
    table,
  });
  const valid = validateOpsDocument(next);
  if (valid.isErr()) {
    return refused({
      op: op,
      reason: valid.error.reason,
      message: valid.error.message,
    });
  }
  const afterIds = new Set(paragraphIdsIn(rows).map(idKey));
  const beforeParagraphIds = paragraphIdsIn(before);
  const afterParagraphIds = paragraphIdsIn(rows);
  const changedIds = new Set(
    rows
      .filter((row) => !before.includes(row))
      .flatMap(paragraphIdsIn)
      .map(idKey),
  );
  const anchor = tableRowAnchor(rows);
  if (anchor === undefined) {
    return refused({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "The table has no paragraph that can address its rows.",
    });
  }
  return Result.ok({
    document: next,
    // Rows put back by themselves would not put the markers beside them back
    // where they stood: the whole table restores them.
    inverse: [
      holdsPositionedMarkup(location.table)
        ? {
            type: DOCUMENT_OP_TYPES.SET_TABLE,
            story: op.story,
            blockId: anchor,
            expected: table,
            table: location.table,
          }
        : {
            type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
            story: op.story,
            blockId: anchor,
            expected: rows,
            rows: before,
          },
    ],
    touched: {
      modified: beforeParagraphIds.filter(
        (id) => afterIds.has(idKey(id)) && changedIds.has(idKey(id)),
      ),
      inserted: afterParagraphIds.filter((id) => !beforeIds.has(idKey(id))),
      removed: beforeParagraphIds.filter((id) => !afterIds.has(idKey(id))),
    },
  });
};

/** A row `owner`'s pending insertion added, holding no one else's suggestion. */
const ownInsertion = (row: TableRow, owner: RevisionOwner): boolean =>
  row.structuralChange?.type === "tableRowInsertion" &&
  isOwnRecord(row.structuralChange.info, owner) &&
  recordedOnlyBy(row, owner);

export const applyRowOp = (
  document: Document,
  op: RowOp,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const located = locateTableRow(document, op);
  if (located.isErr()) return Result.err(located.error);
  const location = located.value;
  const before = location.table.rows;
  const rows = [...before];
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_ROW: {
      if (!Number.isInteger(op.at) || op.at < 0 || op.at > rows.length) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
          message: "The row index is outside the table.",
        });
      }
      if (needsTableEdit(location.table)) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message: "Shared row wrappers require a table operation.",
        });
      }
      const incoming = cloneModel(op.row);
      if (
        paragraphIdsIn(incoming).some((id) =>
          paragraphIdsIn(before).some(
            (existing) => idKey(existing) === idKey(id),
          ),
        )
      ) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
          message: "An inserted paragraph id is already used in the table.",
        });
      }
      const tracked =
        op.revision === undefined
          ? Result.ok([incoming])
          : trackTableRows({
              document,
              op,
              rows: [incoming],
              revision: op.revision,
              newIds: op.newIds,
              kind: WRAP_KINDS.INSERTION,
            });
      if (tracked.isErr()) return Result.err(tracked.error);
      const row =
        tracked.value.at(0) ?? panic("A tracked insertion lost its row.");
      return commitRows({
        document,
        op,
        location,
        rows: rowsWithRow(rows, op.at, row),
      });
    }
    case DOCUMENT_OP_TYPES.DELETE_ROW: {
      const row = before[location.rowIndex];
      if (row === undefined) return panic("A row location must name a row.");
      if (op.expected !== undefined && !equalForStaleness(row, op.expected)) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
          message: "The row to remove has changed.",
        });
      }
      // A row the same author's pending insertion added, with nothing in it
      // anyone else suggested, goes outright, as deleting over one's own
      // insertion removes it; no deletion of an addition is recorded.
      const removesOutright =
        op.revision === undefined || ownInsertion(row, op.revision);
      if (removesOutright && before.length === 1) {
        return commitRows({ document, op, location, rows: [] });
      }
      if (needsTableEdit(location.table)) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message: "Shared row wrappers require a table operation.",
        });
      }
      if (removesOutright || op.revision === undefined) {
        return commitRows({
          document,
          op,
          location,
          rows: rowsWithoutRow(location.table, location.rowIndex),
        });
      }
      const tracked = trackTableRows({
        document,
        op,
        rows: [row],
        revision: op.revision,
        newIds: op.newIds,
        kind: WRAP_KINDS.DELETION,
        survivingCells: mergeOwnersContinuedBelow(
          location.table,
          location.rowIndex,
        ),
      });
      if (tracked.isErr()) return Result.err(tracked.error);
      rows[location.rowIndex] =
        tracked.value.at(0) ?? panic("A tracked deletion lost its row.");
      return commitRows({
        document: document,
        op: op,
        location: location,
        rows: rows,
      });
    }
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS: {
      if (!equalForStaleness(before, op.expected)) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
          message: "The table rows have changed.",
        });
      }
      if (
        op.rows.length > 0 &&
        op.rows.length !== before.length &&
        needsTableEdit(location.table)
      ) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message: "Shared row wrappers require a table operation.",
        });
      }
      const replacement = op.rows.map(
        (row) =>
          before.find((existing) => structurallyEqual(existing, row)) ??
          cloneModel(row),
      );
      return commitRows({
        document: document,
        op: op,
        location: location,
        rows: replacement,
      });
    }
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};
