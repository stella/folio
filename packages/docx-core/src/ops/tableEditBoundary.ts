/** Validate untyped table edits before reading fields or revision stamps. */
import { MAX_REVISION_ID } from "../model/document";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "./types";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const tableSnapshot = (value: unknown): boolean => {
  if (!isRecord(value) || value["type"] !== "table" || !Array.isArray(value["rows"])) return false;
  return Array.from(value["rows"]).every(
    (row: unknown) =>
      isRecord(row) &&
      row["type"] === "tableRow" &&
      Array.isArray(row["cells"]) &&
      Array.from(row["cells"]).every(
        (cell: unknown) =>
          isRecord(cell) &&
          cell["type"] === "tableCell" &&
          Array.isArray(cell["content"]) &&
          Array.from(cell["content"]).every((block: unknown) => {
            if (!isRecord(block)) return false;
            if (block["type"] === "table") return tableSnapshot(block);
            return (
              typeof block["type"] === "string" &&
              (!["paragraph", "blockSdt", "blockCustomXml"].includes(block["type"]) ||
                Array.isArray(block["content"]))
            );
          }),
      ),
  );
};
const strings = (value: unknown): boolean =>
  Array.isArray(value) && Array.from(value).every((id) => typeof id === "string");

export const tableEditBoundaryRefusal = (op: DocumentOp): DocumentOpRefusal | undefined => {
  let valid = true;
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
      valid = strings(op.newBlockIds);
      if (op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN)
        valid &&= typeof op.column === "number" && typeof op.width === "number";
      if (op.type === DOCUMENT_OP_TYPES.MERGE_CELLS)
        valid &&= [op.top, op.bottom, op.left, op.right].every(
          (bound) => typeof bound === "number",
        );
      break;
    case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
      valid =
        Array.isArray(op.columnWidths) &&
        Array.from(op.columnWidths).every((width) => typeof width === "number");
      break;
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
      valid = isRecord(op.patch);
      break;
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
      valid = typeof op.column === "number";
      break;
    case DOCUMENT_OP_TYPES.SET_TABLE:
      valid = tableSnapshot(op.table) && tableSnapshot(op.expected);
      break;
    default:
      return undefined;
  }
  const fail = (reason: DocumentOpRefusal["reason"], message: string) =>
    new DocumentOpRefusal({ opType: op.type, reason, message });
  if (!valid || typeof op.blockId !== "string" || op.story !== "main")
    return fail(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "The table operation is missing required fields or has invalid field shapes.",
    );
  if (op.type === DOCUMENT_OP_TYPES.SET_TABLE) return undefined;
  if (op.revision !== undefined && !isRecord(op.revision))
    return fail(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "A revision stamp must be an object.",
    );
  if (
    op.newIds !== undefined &&
    (!isRecord(op.newIds) ||
      (op.newIds.revision !== undefined && !Array.isArray(op.newIds.revision)))
  )
    return fail(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "Revision ids must be supplied as an array.",
    );
  if (
    op.newIds?.revision !== undefined &&
    !Array.from(op.newIds.revision).every((id) => typeof id === "number")
  )
    return fail(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      "Revision id arrays must contain numbers in every slot.",
    );
  const ids = [...(op.newIds?.revision ?? [])];
  if (op.revision !== undefined) ids.push(op.revision.id);
  if (ids.some((id) => !Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID))
    return fail(
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      "Revision ids must be nonnegative 31-bit integers.",
    );
  if (new Set(ids).size !== ids.length)
    return fail(
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "Revision ids must be distinct within the operation.",
    );
  return undefined;
};
