import { describe, expect, test } from "bun:test";

import type { Document, Table } from "./document";
import {
  assignTableIdentities,
  buildIdentitySidecar,
  IDENTITY_VALIDATION_REASONS,
  parseIdentitySidecar,
  restoreIdentitySidecar,
  stringifyIdentitySidecar,
  validateDocumentIdLease,
  validateLeasedStructuralId,
  type DocumentIdLease,
} from "./identity";

const lease = {
  paragraph: { start: 1, end: 4096 },
  revision: { start: 1, end: 1024 },
  comment: { start: 1025, end: 2048 },
  table: { prefix: "abc", start: 0, end: 15 },
  row: { prefix: "def", start: 0, end: 15 },
  cell: { prefix: "ghi", start: 0, end: 15 },
} satisfies DocumentIdLease;

const table = (): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [
        {
          type: "tableCell",
          content: [{ type: "paragraph", paraId: "00000001", content: [] }],
        },
      ],
    },
  ],
});

const document = (): Document => ({
  package: {
    document: {
      content: [
        table(),
        {
          ...table(),
          rows: [
            {
              ...table().rows[0]!,
              cells: [{ type: "tableCell", content: [table()] }],
            },
          ],
        },
      ],
    },
  },
});

describe("document identity", () => {
  test("assigns stable ids and restores them from JSON", () => {
    const assigned = assignTableIdentities(document(), lease)._unsafeUnwrap();
    const sidecar = buildIdentitySidecar(assigned)._unsafeUnwrap();
    const parsed = parseIdentitySidecar(stringifyIdentitySidecar(sidecar))._unsafeUnwrap();
    const restored = restoreIdentitySidecar(document(), parsed)._unsafeUnwrap();

    expect(buildIdentitySidecar(restored)._unsafeUnwrap()).toEqual(sidecar);
    expect(stringifyIdentitySidecar(parsed)).toBe(stringifyIdentitySidecar(sidecar));
  });

  test("rejects structural ids outside their lease and ids already in use", () => {
    const outside = validateLeasedStructuralId("table", "t_other_a", lease.table, new Set());
    const collision = validateLeasedStructuralId(
      "table",
      "t_abc_a",
      lease.table,
      new Set(["t_abc_a"]),
    );

    expect(outside._unsafeUnwrapErr().reason).toBe(IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE);
    expect(collision._unsafeUnwrapErr().reason).toBe(IDENTITY_VALIDATION_REASONS.ID_COLLISION);
  });

  test("validates a structural id without iterating over the lease", () => {
    const huge = { prefix: "abc", start: 0, end: Number.MAX_SAFE_INTEGER };

    expect(validateLeasedStructuralId("table", "t_abc_a", huge, new Set())._unsafeUnwrap()).toBe(
      "t_abc_a",
    );
    expect(
      validateLeasedStructuralId("table", "t_abc_aa", huge, new Set())._unsafeUnwrapErr().reason,
    ).toBe(IDENTITY_VALIDATION_REASONS.OUTSIDE_LEASE);
  });

  test("returns typed errors for missing and duplicate sidecar identities", () => {
    expect(buildIdentitySidecar(document())._unsafeUnwrapErr().reason).toBe(
      IDENTITY_VALIDATION_REASONS.STRUCTURE_MISMATCH,
    );
    const assigned = assignTableIdentities(document(), lease)._unsafeUnwrap();
    const sidecar = buildIdentitySidecar(assigned)._unsafeUnwrap();
    const duplicateId = structuredClone(sidecar);
    duplicateId.tables[1]!.tableId = duplicateId.tables[0]!.tableId;
    const duplicatePath = structuredClone(sidecar);
    duplicatePath.tables[1]!.story = duplicatePath.tables[0]!.story;
    duplicatePath.tables[1]!.path = duplicatePath.tables[0]!.path;

    expect(restoreIdentitySidecar(document(), duplicateId)._unsafeUnwrapErr().reason).toBe(
      IDENTITY_VALIDATION_REASONS.ID_COLLISION,
    );
    expect(restoreIdentitySidecar(document(), duplicatePath)._unsafeUnwrapErr().reason).toBe(
      IDENTITY_VALIDATION_REASONS.DUPLICATE_PATH,
    );
  });

  test("rejects duplicate existing identities before assigning missing ones", () => {
    const source = document();
    const first = source.package.document.content[0] as Table;
    const second = source.package.document.content[1] as Table;
    first.tableId = "t_existing_a";
    second.tableId = "t_existing_a";

    expect(assignTableIdentities(source, lease)._unsafeUnwrapErr().reason).toBe(
      IDENTITY_VALIDATION_REASONS.ID_COLLISION,
    );
  });

  test("reports overlapping leases and wrong sidecar id kinds", () => {
    expect(validateDocumentIdLease(lease, [lease])._unsafeUnwrapErr().reason).toBe(
      IDENTITY_VALIDATION_REASONS.OVERLAPPING_LEASE,
    );
    const assigned = assignTableIdentities(document(), lease)._unsafeUnwrap();
    const sidecar = buildIdentitySidecar(assigned)._unsafeUnwrap();
    sidecar.tables[0]!.rowIds[0] = "c_abc_a";

    expect(restoreIdentitySidecar(document(), sidecar)._unsafeUnwrapErr().reason).toBe(
      IDENTITY_VALIDATION_REASONS.WRONG_ID_KIND,
    );
  });
});
