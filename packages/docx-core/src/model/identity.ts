import { Result, TaggedError } from "better-result";

import type { Document } from "./document";
import {
  IDENTITY_SPACES,
  packageIdentityKeys,
  packageParagraphIds,
  storiesOf,
  walk,
} from "../ops/ids";

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
  story: string;
  path: string[];
  tableId: string;
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
  OVERLAPPING_LEASE: "overlappingLease",
  OUTSIDE_LEASE: "outsideLease",
  ID_COLLISION: "idCollision",
  DUPLICATE_PATH: "duplicatePath",
  WRONG_ID_KIND: "wrongIdKind",
} as const);
export type IdentityValidationReason =
  (typeof IDENTITY_VALIDATION_REASONS)[keyof typeof IDENTITY_VALIDATION_REASONS];
export class IdentityValidationError extends TaggedError("IdentityValidationError")<{
  message: string;
  reason: IdentityValidationReason;
}> {}

const PREFIX = /^[a-z2-7]+$/u;
const BASE32 = "abcdefghijklmnopqrstuvwxyz234567";
const kindPrefix = { table: "t_", row: "r_", cell: "c_" } as const;
type StructuralKind = keyof typeof kindPrefix;

const encodeBase32 = (value: number): string => {
  if (value === 0) return BASE32.at(0)!;
  const digits: string[] = [];
  for (let remaining = value; remaining > 0; remaining = Math.floor(remaining / 32))
    digits.push(BASE32.at(remaining % 32)!);
  return digits.reverse().join("");
};

const decodeBase32 = (encoded: string): number | undefined => {
  if (encoded.length === 0) return undefined;
  let value = 0;
  for (const digit of encoded) {
    const index = BASE32.indexOf(digit);
    if (index < 0 || value > Math.floor((Number.MAX_SAFE_INTEGER - index) / 32)) return undefined;
    value = value * 32 + index;
  }
  return encodeBase32(value) === encoded ? value : undefined;
};

const structuralId = (kind: StructuralKind, lease: StructuralIdLease, counter: number): string =>
  `${kindPrefix[kind]}${lease.prefix}_${encodeBase32(counter)}`;
const validStructuralLease = ({ prefix, start, end }: StructuralIdLease): boolean =>
  PREFIX.test(prefix) &&
  Number.isSafeInteger(start) &&
  Number.isSafeInteger(end) &&
  start >= 0 &&
  end >= start;
const rangesOverlap = (left: NumericIdLease, right: NumericIdLease): boolean =>
  left.start <= right.end && right.start <= left.end;

export const validateDocumentIdLease = (
  lease: DocumentIdLease,
  otherLeases: readonly DocumentIdLease[] = [],
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
  const overlaps = otherLeases.some(
    (other) =>
      rangesOverlap(lease.paragraph, other.paragraph) ||
      rangesOverlap(lease.revision, other.revision) ||
      rangesOverlap(lease.comment, other.comment) ||
      (["table", "row", "cell"] as const).some(
        (kind) =>
          lease[kind].prefix === other[kind].prefix && rangesOverlap(lease[kind], other[kind]),
      ),
  );
  if (overlaps)
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.OVERLAPPING_LEASE,
        message: "The id lease overlaps an existing lease.",
      }),
    );
  return Result.ok(lease);
};

export const validateLeasedStructuralId = (
  kind: StructuralKind,
  id: string,
  lease: StructuralIdLease,
  used: ReadonlySet<string>,
): Result<string, IdentityValidationError> => {
  if (!validStructuralLease(lease))
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.INVALID_LEASE,
        message: "The structural id lease is invalid.",
      }),
    );
  if (used.has(id))
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.ID_COLLISION,
        message: `The ${kind} id is already used.`,
      }),
    );
  const prefix = `${kindPrefix[kind]}${lease.prefix}_`;
  const counter = id.startsWith(prefix) ? decodeBase32(id.slice(prefix.length)) : undefined;
  if (counter === undefined || counter < lease.start || counter > lease.end)
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
        message: `The ${kind} id is outside the supplied lease.`,
      }),
    );
  return Result.ok(id);
};

const firstFreeParagraphRange = (document: Document): NumericIdLease | undefined => {
  const used = new Set(packageParagraphIds(document.package).map((id) => Number.parseInt(id, 16)));
  for (let start = 1; start <= 0x7fffffff - PARA_ID_LEASE_SIZE + 1; start += PARA_ID_LEASE_SIZE) {
    let free = true;
    for (const id of used) if (id >= start && id < start + PARA_ID_LEASE_SIZE) free = false;
    if (free) return { start, end: start + PARA_ID_LEASE_SIZE - 1 };
  }
  return undefined;
};

/** Issue deterministic, non-overlapping numeric ranges from the document's current identities. */
export const issueDocumentIdLease = (
  document: Document,
  structuralPrefix: string,
): Result<DocumentIdLease, IdentityValidationError> => {
  const paragraph = firstFreeParagraphRange(document);
  if (paragraph === undefined)
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
        message: "No paragraph id lease remains.",
      }),
    );
  let revisionMaximum = 0;
  for (const key of packageIdentityKeys(document.package)) {
    if (!key.startsWith(`${IDENTITY_SPACES.REVISION}:`)) continue;
    revisionMaximum = Math.max(revisionMaximum, Number(key.slice(key.indexOf(":") + 1)));
  }
  let commentMaximum = 0;
  for (const comment of document.package.document.comments ?? [])
    commentMaximum = Math.max(commentMaximum, comment.id);
  const lease = {
    paragraph,
    revision: {
      start: revisionMaximum + 1,
      end: revisionMaximum + NUMERIC_ID_LEASE_SIZE,
    },
    comment: {
      start: commentMaximum + 1,
      end: commentMaximum + NUMERIC_ID_LEASE_SIZE,
    },
    table: { prefix: structuralPrefix, start: 0, end: Number.MAX_SAFE_INTEGER },
    row: { prefix: structuralPrefix, start: 0, end: Number.MAX_SAFE_INTEGER },
    cell: { prefix: structuralPrefix, start: 0, end: Number.MAX_SAFE_INTEGER },
  } satisfies DocumentIdLease;
  return validateDocumentIdLease(lease);
};

type VisitTable = (
  table: Record<string, unknown>,
  story: string,
  path: string[],
) => Result<Record<string, unknown>, IdentityValidationError>;
const transformValue = (
  value: unknown,
  story: string,
  path: string[],
  visit: VisitTable,
): Result<unknown, IdentityValidationError> => {
  if (Array.isArray(value)) {
    const out: unknown[] = [];
    for (let index = 0; index < value.length; index += 1) {
      const transformed = transformValue(value[index], story, [...path, String(index)], visit);
      if (transformed.isErr()) return transformed;
      out.push(transformed.value);
    }
    return Result.ok(out);
  }
  if (value instanceof Map) {
    const out = new Map<unknown, unknown>();
    for (const [key, item] of value) {
      const transformed = transformValue(item, `${story}/${key}`, [], visit);
      if (transformed.isErr()) return transformed;
      out.set(key, transformed.value);
    }
    return Result.ok(out);
  }
  if (
    typeof value !== "object" ||
    value === null ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    return Result.ok(value);
  const record = value as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  for (const [key, item] of Object.entries(record)) {
    if (story === "main" && key === "sections") {
      out[key] = item;
      continue;
    }
    const transformed = transformValue(item, story, [...path, key], visit);
    if (transformed.isErr()) return transformed;
    out[key] = transformed.value;
  }
  return out.type === "table" ? visit(out, story, path) : Result.ok(out);
};

const transformDocument = (
  document: Document,
  visit: VisitTable,
): Result<Document, IdentityValidationError> => {
  const pkg: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(document.package)) {
    const story = key === "document" ? "main" : key;
    const transformed = transformValue(value, story, [], visit);
    if (transformed.isErr()) return Result.err(transformed.error);
    pkg[key] = transformed.value;
  }
  return Result.ok({ ...document, package: pkg as Document["package"] });
};

export const assignTableIdentities = (
  document: Document,
  lease: DocumentIdLease,
): Result<Document, IdentityValidationError> => {
  const valid = validateDocumentIdLease(lease);
  if (valid.isErr()) return Result.err(valid.error);
  const counters = {
    table: lease.table.start,
    row: lease.row.start,
    cell: lease.cell.start,
  };
  const used = {
    table: new Set<string>(),
    row: new Set<string>(),
    cell: new Set<string>(),
  };
  walk({
    value: storiesOf(document.package),
    visit: (entries) => {
      const record = Object.fromEntries(entries);
      if (record.type === "table" && typeof record.tableId === "string")
        used.table.add(record.tableId);
      if (record.type === "tableRow" && typeof record.rowId === "string")
        used.row.add(record.rowId);
      if (record.type === "tableCell" && typeof record.cellId === "string")
        used.cell.add(record.cellId);
    },
  });
  const seen = {
    table: new Set<string>(),
    row: new Set<string>(),
    cell: new Set<string>(),
  };
  const take = (
    kind: StructuralKind,
    existing: unknown,
  ): Result<string, IdentityValidationError> => {
    if (typeof existing === "string") {
      if (seen[kind].has(existing))
        return Result.err(
          new IdentityValidationError({
            reason: IDENTITY_VALIDATION_REASONS.ID_COLLISION,
            message: `The document repeats a ${kind} id.`,
          }),
        );
      seen[kind].add(existing);
      return Result.ok(existing);
    }
    while (counters[kind] <= lease[kind].end) {
      const candidate = structuralId(kind, lease[kind], counters[kind]++);
      if (!used[kind].has(candidate) && !seen[kind].has(candidate)) {
        seen[kind].add(candidate);
        return Result.ok(candidate);
      }
    }
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE,
        message: `The document needs more ${kind} ids than the lease provides.`,
      }),
    );
  };
  return transformDocument(document, (table) => {
    const tableId = take("table", table.tableId);
    if (tableId.isErr()) return tableId;
    const rows: unknown[] = [];
    for (const rowValue of table.rows as Record<string, unknown>[]) {
      const rowId = take("row", rowValue.rowId);
      if (rowId.isErr()) return rowId;
      const cells: unknown[] = [];
      for (const cellValue of rowValue.cells as Record<string, unknown>[]) {
        const cellId = take("cell", cellValue.cellId);
        if (cellId.isErr()) return cellId;
        cells.push({ ...cellValue, cellId: cellId.value });
      }
      rows.push({ ...rowValue, rowId: rowId.value, cells });
    }
    return Result.ok({ ...table, tableId: tableId.value, rows });
  });
};

const tableEntry = (
  table: Record<string, unknown>,
  story: string,
  path: string[],
): Result<TableIdentitySidecarEntry, IdentityValidationError> => {
  const rows = table.rows as Record<string, unknown>[];
  if (
    typeof table.tableId !== "string" ||
    rows.some(
      (row) =>
        typeof row.rowId !== "string" ||
        (row.cells as Record<string, unknown>[]).some((cell) => typeof cell.cellId !== "string"),
    )
  )
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
        message: "Every table, row and cell must have a model identity.",
      }),
    );
  return Result.ok({
    story,
    path,
    tableId: table.tableId,
    rowIds: rows.map((row) => row.rowId as string),
    cellIds: rows.map((row) =>
      (row.cells as Record<string, unknown>[]).map((cell) => cell.cellId as string),
    ),
  });
};

export const buildIdentitySidecar = (
  document: Document,
): Result<DocumentIdentitySidecar, IdentityValidationError> => {
  const tables: TableIdentitySidecarEntry[] = [];
  const collected = transformDocument(document, (table, story, path) => {
    const entry = tableEntry(table, story, path);
    if (entry.isErr()) return Result.err(entry.error);
    tables.push(entry.value);
    return Result.ok(table);
  });
  return collected.isErr()
    ? Result.err(collected.error)
    : Result.ok({ version: IDENTITY_SIDECAR_VERSION, tables });
};

const isEntry = (value: unknown): value is TableIdentitySidecarEntry => {
  if (typeof value !== "object" || value === null) return false;
  const entry = value as Record<string, unknown>;
  return (
    typeof entry.story === "string" &&
    Array.isArray(entry.path) &&
    entry.path.every((part) => typeof part === "string") &&
    typeof entry.tableId === "string" &&
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
  if (parsed.isErr() || typeof parsed.value !== "object" || parsed.value === null)
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.INVALID_SIDECAR,
        message: "The identity sidecar is not valid JSON.",
      }),
    );
  const value = parsed.value as Record<string, unknown>;
  if (
    value.version !== IDENTITY_SIDECAR_VERSION ||
    !Array.isArray(value.tables) ||
    !value.tables.every(isEntry)
  )
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.INVALID_SIDECAR,
        message: "The identity sidecar has an unsupported shape.",
      }),
    );
  return Result.ok({ version: IDENTITY_SIDECAR_VERSION, tables: value.tables });
};

const sidecarKey = ({ story, path }: Pick<TableIdentitySidecarEntry, "story" | "path">) =>
  `${story}:${JSON.stringify(path)}`;
export const restoreIdentitySidecar = (
  document: Document,
  sidecar: DocumentIdentitySidecar,
): Result<Document, IdentityValidationError> => {
  const entries = new Map<string, TableIdentitySidecarEntry>();
  const ids = new Set<string>();
  for (const entry of sidecar.tables) {
    const key = sidecarKey(entry);
    if (entries.has(key))
      return Result.err(
        new IdentityValidationError({
          reason: IDENTITY_VALIDATION_REASONS.DUPLICATE_PATH,
          message: "The identity sidecar repeats a table path.",
        }),
      );
    entries.set(key, entry);
    for (const [kind, values] of [
      ["table", [entry.tableId]],
      ["row", entry.rowIds],
      ["cell", entry.cellIds.flat()],
    ] as const) {
      for (const id of values) {
        if (!id.startsWith(kindPrefix[kind]))
          return Result.err(
            new IdentityValidationError({
              reason: IDENTITY_VALIDATION_REASONS.WRONG_ID_KIND,
              message: `A sidecar ${kind} id has the wrong kind prefix.`,
            }),
          );
        if (ids.has(id))
          return Result.err(
            new IdentityValidationError({
              reason: IDENTITY_VALIDATION_REASONS.ID_COLLISION,
              message: "The identity sidecar repeats an id.",
            }),
          );
        ids.add(id);
      }
    }
  }
  const used = new Set<string>();
  const restored = transformDocument(document, (table, story, path) => {
    const entry = entries.get(sidecarKey({ story, path }));
    const rows = table.rows as Record<string, unknown>[];
    if (
      entry === undefined ||
      entry.rowIds.length !== rows.length ||
      entry.cellIds.length !== rows.length ||
      rows.some((row, index) => entry.cellIds[index]?.length !== (row.cells as unknown[]).length)
    )
      return Result.err(
        new IdentityValidationError({
          reason: IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
          message: "The sidecar does not match the document structure.",
        }),
      );
    used.add(sidecarKey(entry));
    const restoredRows: Record<string, unknown>[] = [];
    for (const [rowIndex, row] of rows.entries()) {
      const restoredCells: Record<string, unknown>[] = [];
      for (const [cellIndex, cell] of (row.cells as Record<string, unknown>[]).entries())
        restoredCells.push(
          Object.assign({}, cell, { cellId: entry.cellIds[rowIndex]![cellIndex] }),
        );
      restoredRows.push(
        Object.assign({}, row, { rowId: entry.rowIds[rowIndex], cells: restoredCells }),
      );
    }
    return Result.ok(Object.assign({}, table, { tableId: entry.tableId, rows: restoredRows }));
  });
  if (restored.isErr()) return restored;
  if (used.size !== entries.size)
    return Result.err(
      new IdentityValidationError({
        reason: IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
        message: "The sidecar contains paths absent from the document.",
      }),
    );
  return restored;
};
export const stringifyIdentitySidecar = (sidecar: DocumentIdentitySidecar): string =>
  JSON.stringify(sidecar);
