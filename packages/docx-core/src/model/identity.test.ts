import { describe, expect, test } from "bun:test";

import type { Document, Table } from "./document";
import {
  assignTableIdentities,
  buildIdentitySidecar,
  IDENTITY_VALIDATION_REASONS,
  parseIdentitySidecar,
  restoreIdentitySidecar,
  stringifyIdentitySidecar,
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
    const sidecar = buildIdentitySidecar(assigned);
    const parsed = parseIdentitySidecar(stringifyIdentitySidecar(sidecar))._unsafeUnwrap();
    const restored = restoreIdentitySidecar(document(), parsed)._unsafeUnwrap();

    expect(buildIdentitySidecar(restored)).toEqual(sidecar);
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
});
