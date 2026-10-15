import { Result, TaggedError } from "better-result";

import type { BlockContent, Document, Table } from "./document";

export const IDENTITY_SIDECAR_VERSION = 1 as const;
export const PARA_ID_LEASE_SIZE = 4096;
export const NUMERIC_ID_LEASE_SIZE = 1024;

export type NumericIdLease = { start: number; end: number };
export type StructuralIdLease = { prefix: string; start: number; end: number };
export type DocumentIdLease = {
  paragraph: NumericIdLease;
  revision: NumericIdLease;
  comment: NumericIdLease;
  table: StructuralIdLease;
  row: StructuralIdLease;
  cell: StructuralIdLease;
};

export type TableIdentitySidecarEntry = {
  tableId: string;
  path: number[];
  rowIds: string[];
  cellIds: string[][];
};

export type DocumentIdentitySidecar = {
  version: typeof IDENTITY_SIDECAR_VERSION;
  tables: TableIdentitySidecarEntry[];
};

export const IDENTITY_VALIDATION_REASONS = Object.freeze({
  INVALID_SIDECAR: "invalidSidecar",
  STRUCTURE_MISMATCH: "structureMismatch",
  INVALID_LEASE: "invalidLease",
  OUTSIDE_LEASE: "outsideLease",
  ID_COLLISION: "idCollision",
} as const);

export type IdentityValidationReason =
  (typeof IDENTITY_VALIDATION_REASONS)[keyof typeof IDENTITY_VALIDATION_REASONS];

export class IdentityValidationError extends TaggedError("IdentityValidationError")<{
  message: string;
  reason: IdentityValidationReason;
}> {}

const PREFIX = /^[a-z2-7]+$/u;
const kindPrefix = { table: "t_", row: "r_", cell: "c_" } as const;
type StructuralKind = keyof typeof kindPrefix;

const BASE32_ALPHABET = "abcdefghijklmnopqrstuvwxyz234567";
const encodeBase32 = (value: number): string => {
  if (value === 0) return BASE32_ALPHABET[0]!;
  const digits: string[] = [];
  let remaining = value;
  while (remaining > 0) {
    digits.push(BASE32_ALPHABET[remaining % 32]!);
    remaining = Math.floor(remaining / 32);
  }
  return digits.reverse().join("");
};

const structuralId = (kind: StructuralKind, lease: StructuralIdLease, counter: number): string =>
  `${kindPrefix[kind]}${lease.prefix}_${encodeBase32(counter)}`;

const validStructuralLease = ({ prefix, start, end }: StructuralIdLease): boolean =>
  PREFIX.test(prefix) &&
  Number.isSafeInteger(start) &&
  Number.isSafeInteger(end) &&
  start >= 0 &&
  end >= start;

export const validateDocumentIdLease = (
  lease: DocumentIdLease,
): Result<DocumentIdLease, IdentityValidationError> => {
  const numeric = [lease.paragraph, lease.revision, lease.comment];
  if (
    numeric.some(
      ({ start, end }) =>
        !Number.isInteger(start) ||
        !Number.isInteger(end) ||
        start < 1 ||
        end < start ||
        end > 0x7fffffff,
    ) ||
    !validStructuralLease(lease.table) ||
    !validStructuralLease(lease.row) ||
    !validStructuralLease(lease.cell)
  ) {
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.INVALID_LEASE,
        message: "The id lease contains an invalid range or prefix.",
      }),
    );
  }
  return Result.ok(lease);
};

export const validateLeasedStructuralId = (
  kind: StructuralKind,
  id: string,
  lease: StructuralIdLease,
  used: ReadonlySet<string>,
): Result<string, IdentityValidationError> => {
  if (used.has(id)) {
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.ID_COLLISION,
        message: `The ${kind} id is already used.`,
      }),
    );
  }
  for (let counter = lease.start; counter <= lease.end; counter += 1) {
    if (id === structuralId(kind, lease, counter)) return Result.ok(id);
  }
  return Result.err(
    new IdentityValidationError({
      reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
      message: `The ${kind} id is outside the supplied lease.`,
    }),
  );
};

type Counters = { table: number; row: number; cell: number };

const assignInBlocks = (
  blocks: readonly BlockContent[],
  lease: DocumentIdLease,
  counters: Counters,
): BlockContent[] =>
  blocks.map((block) => {
    if (block.type !== "table") return block;
    const tableId = block.tableId ?? structuralId("table", lease.table, counters.table++);
    const rows = block.rows.map((row) => ({
      ...row,
      rowId: row.rowId ?? structuralId("row", lease.row, counters.row++),
      cells: row.cells.map((cell) => ({
        ...cell,
        cellId: cell.cellId ?? structuralId("cell", lease.cell, counters.cell++),
        content: assignInBlocks(cell.content, lease, counters),
      })),
    }));
    return { ...block, tableId, rows };
  });

/** Assign model-only table identities in document order using the supplied lease. */
export const assignTableIdentities = (
  document: Document,
  lease: DocumentIdLease,
): Result<Document, IdentityValidationError> => {
  const valid = validateDocumentIdLease(lease);
  if (valid.isErr()) return Result.err(valid.error);
  const counters = { table: lease.table.start, row: lease.row.start, cell: lease.cell.start };
  const content = assignInBlocks(document.package.document.content, lease, counters);
  if (
    counters.table > lease.table.end + 1 ||
    counters.row > lease.row.end + 1 ||
    counters.cell > lease.cell.end + 1
  ) {
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
        message: "The document needs more structural ids than the lease provides.",
      }),
    );
  }
  return Result.ok({
    ...document,
    package: { ...document.package, document: { ...document.package.document, content } },
  });
};

const collectTables = (
  blocks: readonly BlockContent[],
  prefix: readonly number[],
  entries: TableIdentitySidecarEntry[],
): void => {
  blocks.forEach((block, blockIndex) => {
    if (block.type !== "table") return;
    const path = [...prefix, blockIndex];
    if (
      block.tableId === undefined ||
      block.rows.some(
        (row) => row.rowId === undefined || row.cells.some((cell) => cell.cellId === undefined),
      )
    ) {
      throw new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
        message: "Every table, row and cell must have a model identity.",
      });
    }
    entries.push({
      tableId: block.tableId,
      path,
      rowIds: block.rows.map((row) => row.rowId!),
      cellIds: block.rows.map((row) => row.cells.map((cell) => cell.cellId!)),
    });
    block.rows.forEach((row, rowIndex) =>
      row.cells.forEach((cell, cellIndex) =>
        collectTables(cell.content, [...path, rowIndex, cellIndex], entries),
      ),
    );
  });
};

export const buildIdentitySidecar = (document: Document): DocumentIdentitySidecar => {
  const tables: TableIdentitySidecarEntry[] = [];
  collectTables(document.package.document.content, [], tables);
  return { version: IDENTITY_SIDECAR_VERSION, tables };
};

const isEntry = (value: unknown): value is TableIdentitySidecarEntry => {
  if (typeof value !== "object" || value === null) return false;
  const entry = value as Record<string, unknown>;
  return (
    typeof entry.tableId === "string" &&
    Array.isArray(entry.path) &&
    entry.path.every(Number.isSafeInteger) &&
    Array.isArray(entry.rowIds) &&
    entry.rowIds.every((id) => typeof id === "string") &&
    Array.isArray(entry.cellIds) &&
    entry.cellIds.every((ids) => Array.isArray(ids) && ids.every((id) => typeof id === "string"))
  );
};

export const parseIdentitySidecar = (
  json: string,
): Result<DocumentIdentitySidecar, IdentityValidationError> => {
  const parsed = Result.try(() => JSON.parse(json) as unknown);
  if (parsed.isErr() || typeof parsed.value !== "object" || parsed.value === null) {
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.INVALID_SIDECAR,
        message: "The identity sidecar is not valid JSON.",
      }),
    );
  }
  const value = parsed.value as Record<string, unknown>;
  if (
    value.version !== IDENTITY_SIDECAR_VERSION ||
    !Array.isArray(value.tables) ||
    !value.tables.every(isEntry)
  ) {
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.INVALID_SIDECAR,
        message: "The identity sidecar has an unsupported shape.",
      }),
    );
  }
  return Result.ok({ version: IDENTITY_SIDECAR_VERSION, tables: value.tables });
};

const restoreInBlocks = (
  blocks: readonly BlockContent[],
  prefix: readonly number[],
  entries: ReadonlyMap<string, TableIdentitySidecarEntry>,
): BlockContent[] =>
  blocks.map((block, blockIndex) => {
    if (block.type !== "table") return block;
    const path = [...prefix, blockIndex];
    const entry = entries.get(path.join("/"));
    if (
      entry === undefined ||
      entry.rowIds.length !== block.rows.length ||
      entry.cellIds.length !== block.rows.length
    ) {
      throw new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
        message: "The sidecar does not match the document structure.",
      });
    }
    const rows = block.rows.map((row, rowIndex) => {
      if (entry.cellIds[rowIndex]?.length !== row.cells.length) {
        throw new IdentityValidationError({
          reason: IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
          message: "The sidecar does not match the document structure.",
        });
      }
      return {
        ...row,
        rowId: entry.rowIds[rowIndex],
        cells: row.cells.map((cell, cellIndex) => ({
          ...cell,
          cellId: entry.cellIds[rowIndex]![cellIndex],
          content: restoreInBlocks(cell.content, [...path, rowIndex, cellIndex], entries),
        })),
      };
    });
    return { ...block, tableId: entry.tableId, rows } satisfies Table;
  });

export const restoreIdentitySidecar = (
  document: Document,
  sidecar: DocumentIdentitySidecar,
): Result<Document, IdentityValidationError> => {
  const entries = new Map(sidecar.tables.map((entry) => [entry.path.join("/"), entry]));
  return Result.try({
    try: () => {
      const content = restoreInBlocks(document.package.document.content, [], entries);
      const restored = {
        ...document,
        package: { ...document.package, document: { ...document.package.document, content } },
      };
      if (buildIdentitySidecar(restored).tables.length !== entries.size) {
        throw new IdentityValidationError({
          reason: IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
          message: "The sidecar contains paths absent from the document.",
        });
      }
      return restored;
    },
    catch: (error) => {
      if (error instanceof IdentityValidationError) return error;
      return new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.INVALID_SIDECAR,
        message: "The identity sidecar could not be restored.",
      });
    },
  });
};

export const stringifyIdentitySidecar = (sidecar: DocumentIdentitySidecar): string =>
  JSON.stringify(sidecar);
