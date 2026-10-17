/** Validate caller-supplied identities before an operation changes the model. */
import type { Document } from "../model/document";
import {
  IDENTITY_VALIDATION_REASONS,
  validateDocumentIdLease,
  validateLeasedStructuralId,
  type DocumentIdLease,
} from "../model/identity";
import { idKey, isParaId, packageParagraphIds, packageStructuralIds, walk } from "./ids";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp } from "./types";

const record = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const LEASE_FIELDS = {
  paragraph: "numeric",
  revision: "numeric",
  comment: "numeric",
  table: "structural",
  row: "structural",
  cell: "structural",
} as const satisfies Record<keyof DocumentIdLease, "numeric" | "structural">;
const leaseShape = (value: unknown): value is DocumentIdLease =>
  record(value) &&
  Object.entries(LEASE_FIELDS).every(([kind, type]) => {
    const range = value[kind];
    return (
      record(range) &&
      typeof range.start === "number" &&
      typeof range.end === "number" &&
      (type === "numeric" || typeof range.prefix === "string")
    );
  });

export const leasedOperationRefusal = (
  document: Document,
  op: DocumentOp,
): DocumentOpRefusal | undefined => {
  if (
    op.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK &&
    op.type !== DOCUMENT_OP_TYPES.INSERT_ROW &&
    op.type !== DOCUMENT_OP_TYPES.INSERT_COLUMN &&
    op.type !== DOCUMENT_OP_TYPES.MERGE_CELLS &&
    op.type !== DOCUMENT_OP_TYPES.SPLIT_CELL
  )
    return undefined;
  const fail = (reason: DocumentOpRefusal["reason"], message: string) =>
    new DocumentOpRefusal({ opType: op.type, reason, message });
  if (!leaseShape(op.idLease))
    return fail(IDENTITY_VALIDATION_REASONS.INVALID_LEASE, "The operation requires an id lease.");
  const checked = validateDocumentIdLease(op.idLease);
  if (checked.isErr()) return fail(checked.error.reason, checked.error.message);
  const lease = checked.value;
  const paragraphs = new Set(packageParagraphIds(document.package).map(idKey));
  const structures = new Set(packageStructuralIds(document.package));
  const validateParagraph = (id: unknown): DocumentOpRefusal | undefined => {
    if (typeof id !== "string" || !isParaId(id))
      return fail(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "A new block id is not a usable paragraph id.",
      );
    if (paragraphs.has(idKey(id)))
      return fail(IDENTITY_VALIDATION_REASONS.ID_COLLISION, "A new block id is already used.");
    const numeric = Number.parseInt(id, 16);
    if (numeric < lease.paragraph.start || numeric > lease.paragraph.end)
      return fail(
        IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
        "A new block id is outside the lease.",
      );
    paragraphs.add(idKey(id));
    return undefined;
  };
  type StructuralKind = "table" | "row" | "cell";
  const validateStructure = (kind: StructuralKind, id: unknown): DocumentOpRefusal | undefined => {
    if (typeof id !== "string")
      return fail(
        IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
        "Every new table record needs an id.",
      );
    const valid = validateLeasedStructuralId(kind, id, lease[kind], structures);
    if (valid.isErr()) return fail(valid.error.reason, valid.error.message);
    structures.add(id);
    return undefined;
  };
  switch (op.type) {
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return validateParagraph(op.newBlockId);
    case DOCUMENT_OP_TYPES.INSERT_ROW: {
      let failure: DocumentOpRefusal | undefined;
      walk({
        value: op.row,
        visit: (entries) => {
          if (failure !== undefined) return;
          const value = Object.fromEntries(entries);
          switch (value.type) {
            case "paragraph":
              failure = validateParagraph(value.paraId);
              break;
            case "table":
              failure = validateStructure("table", value.tableId);
              break;
            case "tableRow":
              failure = validateStructure("row", value.rowId);
              break;
            case "tableCell":
              failure = validateStructure("cell", value.cellId);
              break;
          }
        },
      });
      return failure;
    }
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
    case DOCUMENT_OP_TYPES.SPLIT_CELL: {
      if (!Array.isArray(op.newBlockIds) || !Array.isArray(op.newCellIds))
        return fail(
          IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
          "New cells require block and cell id arrays.",
        );
      if (op.newCellIds.length !== op.newBlockIds.length)
        return fail(
          IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
          "Each new cell requires one cell id.",
        );
      for (const id of op.newBlockIds) {
        const failure = validateParagraph(id);
        if (failure !== undefined) return failure;
      }
      for (const id of op.newCellIds) {
        const failure = validateStructure("cell", id);
        if (failure !== undefined) return failure;
      }
      return undefined;
    }
    default: {
      const exhaustive: never = op;
      return exhaustive;
    }
  }
};
