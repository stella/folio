/** Row edits and their exact inverse, addressed through a paragraph in the row. */
import { Result, panic } from "better-result";
import { applyTableOp } from "./tables";
import { locateTableRow, tableRowAnchor, type TableRowLocation } from "./tableLocation";

import type {
  BlockContent,
  Document,
  DocumentBody,
  Paragraph,
  Table,
  TableRow,
} from "../model/document";
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
import { freshenIdentities } from "./identity";
import {
  IDENTITY_SPACES,
  collides,
  countIds,
  countKeys,
  identityKeysIn,
  idKey,
  isParaId,
  packageIdentityKeys,
  packageParagraphIds,
  paragraphIdsIn,
  slotKey,
} from "./ids";
import { textsIn } from "./leaves";
import { hasIllegalXmlCharacters } from "../serialize/xmlEscape";
import { paragraphLength } from "./offsets";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { stampInfo, wrapTracked, WRAP_KINDS } from "./review";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type DeleteRowOp,
  type InsertRowOp,
  type OpStory,
  type RevisionStamp,
  type SetTableRowsOp,
  type NewIds,
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

/** Rebuild only paths to changed paragraphs, retaining every other record. */
const mapParagraphs = (
  blocks: readonly BlockContent[],
  replace: (paragraph: Paragraph) => Paragraph,
): BlockContent[] =>
  blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph":
        return replace(block);
      case "blockSdt":
      case "blockCustomXml":
        return { ...block, content: mapParagraphs(block.content, replace) };
      case "table":
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: mapParagraphs(cell.content, replace),
            })),
          })),
        };
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        return block;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

const rowWithParagraphs = (
  row: TableRow,
  replace: (paragraph: Paragraph) => Paragraph,
): TableRow => ({
  ...row,
  cells: row.cells.map((cell) => ({
    ...cell,
    content: mapParagraphs(cell.content, replace),
  })),
});

/** Indexed markup and shared row wrappers need a structural table operation. */
const needsTableEdit = (table: Table): boolean =>
  table.preserved !== undefined ||
  table.bookmarks !== undefined ||
  table.carrierStack !== undefined ||
  table.rows.some((row) => row.contentControls !== undefined || row.carrierStack !== undefined);

/** Nested tables, captured blocks and cell revisions have no tracked-row construction yet. */
const trackableBlocks = (blocks: readonly BlockContent[]): boolean =>
  blocks.every((block) => {
    switch (block.type) {
      case "paragraph":
        return block.sectionProperties === undefined;
      case "blockSdt":
      case "blockCustomXml":
        return trackableBlocks(block.content);
      case "table":
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        return false;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  });

type TrackRowOptions = {
  document: Document;
  op: InsertRowOp | DeleteRowOp;
  row: TableRow;
  revision: RevisionStamp;
  newIds: NewIds | undefined;
  kind: (typeof WRAP_KINDS)[keyof typeof WRAP_KINDS];
};

const trackRow = ({
  document,
  op,
  row,
  revision,
  newIds,
  kind,
}: TrackRowOptions): Result<TableRow, DocumentOpRefusal> => {
  if (row.structuralChange !== undefined) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      message: "The row already carries a structural revision.",
    });
  }
  if (
    row.cells.some((cell) => cell.structuralChange !== undefined || !trackableBlocks(cell.content))
  ) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "The row contains an unsupported tracked structure.",
    });
  }
  const before = paragraphsIn([row]);
  const after: Paragraph[] = [];
  for (const paragraph of before) {
    const wrapped = wrapTracked({
      items: paragraph.content,
      from: { offset: 0, zeroWidthBefore: 0 },
      to: {
        offset: paragraphLength(paragraph),
        zeroWidthBefore: Number.MAX_SAFE_INTEGER,
      },
      kind,
      stamp: revision,
    });
    if (wrapped.kind === "refused") {
      return refused({
        op: op,
        reason: wrapped.reason,
        message: "The row's content cannot record this revision.",
      });
    }
    after.push(
      wrapped.kind === "unchanged" ? paragraph : { ...paragraph, content: wrapped.content },
    );
  }
  // The row gets the stamp's first id. Its inline wrappers take the supplied
  // additional ids; existing records keep theirs, including nested changes.
  const outside = new Set(packageIdentityKeys(document.package));
  if (op.type === DOCUMENT_OP_TYPES.DELETE_ROW) {
    for (const key of identityKeysIn(before)) outside.delete(key);
  }
  outside.add(slotKey({ space: IDENTITY_SPACES.REVISION, id: revision.id }));
  const freshened = freshenIdentities({
    before,
    after,
    newIds: newIds ?? {},
    usedElsewhere: () => outside,
  });
  switch (freshened.kind) {
    case "needsIds":
      return refused({
        op: op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        message: `The row needs ${freshened.missing} additional revision ids.`,
      });
    case "invalidId":
      return refused({
        op: op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        message: `${freshened.id} cannot be a new revision id.`,
      });
    case "fresh": {
      const byId = new Map(
        freshened.paragraphs.map((paragraph) => [idKey(paragraph.paraId ?? ""), paragraph]),
      );
      return Result.ok({
        ...rowWithParagraphs(
          row,
          (paragraph) =>
            byId.get(idKey(paragraph.paraId ?? "")) ?? panic("A tracked row lost a paragraph."),
        ),
        structuralChange: {
          type: kind === WRAP_KINDS.INSERTION ? "tableRowInsertion" : "tableRowDeletion",
          info: stampInfo(revision),
        },
      });
    }
    default: {
      const unreachable: never = freshened;
      return unreachable;
    }
  }
};

type WithRowsOptions = {
  document: Document;
  story: OpStory;
  location: TableRowLocation;
  rows: TableRow[];
};
const withRows = ({ document, story, location, rows }: WithRowsOptions): Document => {
  const body = storyBody(document, story);
  const content = updateBlockList(body.content, location.list, (blocks) => {
    const out = [...blocks];
    out[location.index] = {
      ...location.table,
      rows,
    };
    return out;
  });
  switch (story) {
    case OP_STORIES.MAIN:
      return {
        ...document,
        package: {
          ...document.package,
          document: withBodyContent(body, content),
        },
      };
    default: {
      const unreachable: never = story;
      return unreachable;
    }
  }
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
        row.cells.some((cell) => storyParagraphs({ content: cell.content }).length === 0),
    )
  ) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      message: "Every row has cells, and every cell has a paragraph.",
    });
  }
  if (afterParagraphs.some(({ content }) => textsIn(content).some(hasIllegalXmlCharacters))) {
    return refused({
      op: op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      message: "The row holds text that cannot be written to XML.",
    });
  }
  const body = { content: [{ type: "table", rows }] } satisfies DocumentBody;
  const originalParagraphs = new Map(
    paragraphsIn(before).map((paragraph) => [idKey(paragraph.paraId ?? ""), paragraph]),
  );
  for (const paragraphLocation of storyParagraphs(body)) {
    const mark = paragraphLocation.paragraph.pPrMark;
    if (
      mark !== undefined &&
      endsItsContainer(body, paragraphLocation) &&
      !structurallyEqual(
        mark,
        originalParagraphs.get(idKey(paragraphLocation.paragraph.paraId ?? ""))?.pPrMark,
      )
    ) {
      return refused({
        op: op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        message: "A row operation cannot add a mark to a cell's final paragraph.",
      });
    }
  }
  const beforeIds = new Set(paragraphIdsIn(before).map(idKey));
  if (
    afterParagraphs.some(
      ({ paraId }) => paraId === undefined || (!beforeIds.has(idKey(paraId)) && !isParaId(paraId)),
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
    remainingParagraphs.set(idKey(id), (remainingParagraphs.get(idKey(id)) ?? 1) - 1);
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
      message: "A row revision or content-control id is already used in the package.",
    });
  }
  const next = withRows({ document: document, story: op.story, location: location, rows: rows });
  const valid = validateOpsDocument(next);
  if (valid.isErr()) {
    return refused({ op: op, reason: valid.error.reason, message: valid.error.message });
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
    inverse: [
      {
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
          message: "Indexed table markup or shared row wrappers require a table operation.",
        });
      }
      const incoming = structuredClone(op.row);
      if (
        paragraphIdsIn(incoming).some((id) =>
          paragraphIdsIn(before).some((existing) => idKey(existing) === idKey(id)),
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
          ? Result.ok(incoming)
          : trackRow({
              document,
              op,
              row: incoming,
              revision: op.revision,
              newIds: op.newIds,
              kind: WRAP_KINDS.INSERTION,
            });
      if (tracked.isErr()) return Result.err(tracked.error);
      rows.splice(op.at, 0, tracked.value);
      return commitRows({ document: document, op: op, location: location, rows: rows });
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
      if (op.revision === undefined && before.length === 1) {
        return commitRows({ document, op, location, rows: [] });
      }
      if (needsTableEdit(location.table)) {
        return refused({
          op: op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message: "Indexed table markup or shared row wrappers require a table operation.",
        });
      }
      if (op.revision === undefined) {
        rows.splice(location.rowIndex, 1);
      } else {
        if (
          before.filter((candidate) => candidate.structuralChange?.type !== "tableRowDeletion")
            .length <= 1
        ) {
          return refused({
            op: op,
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message: "Deleting the last remaining row requires a table operation.",
          });
        }
        const tracked = trackRow({
          document,
          op,
          row,
          revision: op.revision,
          newIds: op.newIds,
          kind: WRAP_KINDS.DELETION,
        });
        if (tracked.isErr()) return Result.err(tracked.error);
        rows[location.rowIndex] = tracked.value;
      }
      return commitRows({ document: document, op: op, location: location, rows: rows });
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
          message: "Indexed table markup or shared row wrappers require a table operation.",
        });
      }
      const replacement = op.rows.map(
        (row) =>
          before.find((existing) => structurallyEqual(existing, row)) ?? structuredClone(row),
      );
      return commitRows({ document: document, op: op, location: location, rows: replacement });
    }
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};
