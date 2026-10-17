import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { parseDocx } from "../../../core/src/docx/parser";
import { serializeDocumentToDocx } from "../serialize/docx";

import type { Document, Table } from "./document";
import {
  assignTableIdentities,
  buildIdentitySidecar,
  IDENTITY_VALIDATION_REASONS,
  IdentityValidationError,
  parseIdentitySidecar,
  restoreIdentitySidecar,
  stringifyIdentitySidecar,
  validateDocumentIdLease,
  validateLeasedStructuralId,
  type DocumentIdLease,
  type IdentityValidationReason,
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

const fixtureRoots = [
  path.resolve(import.meta.dir, "../../../core/src/docx/__tests__/__fixtures__/corpus"),
  path.resolve(import.meta.dir, "../../../../tests/visual/fixtures"),
];
const corpusFixtures = fixtureRoots.flatMap((root) =>
  readdirSync(root, { recursive: true })
    .filter((name) => name.endsWith(".docx"))
    .sort()
    .map((name) => path.join(root, name)),
);

const identityLeaseArbitrary = fc
  .record({
    prefix: fc
      .array(fc.constantFrom(..."abcdefghijklmnopqrstuvwxyz234567"), {
        minLength: 1,
        maxLength: 12,
      })
      .map((letters) => letters.join("")),
    start: fc.integer({ min: 0, max: 1_000_000 }),
  })
  .map(({ prefix, start }) =>
    Object.assign({}, lease, {
      table: { prefix, start, end: start + 1_000_000 },
      row: { prefix, start, end: start + 1_000_000 },
      cell: { prefix, start, end: start + 1_000_000 },
    }),
  );

describe("corpus table identity sidecars", () => {
  test("discovers DOCX fixtures in each corpus", () => {
    for (const root of fixtureRoots) {
      const fixtures = corpusFixtures.filter((filename) =>
        filename.startsWith(`${root}${path.sep}`),
      );
      expect(fixtures.length).toBeGreaterThan(0);
    }
  });

  test.each(corpusFixtures)(
    "%s round-trips identities without changing serialized bytes",
    async (filename) => {
      const bytes = readFileSync(filename);
      const source = await parseDocx(
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
        { preloadFonts: false },
      );
      const probe = assignTableIdentities(source, {
        ...lease,
        table: { prefix: "corpus", start: 0, end: 1_000_000 },
        row: { prefix: "corpus", start: 0, end: 1_000_000 },
        cell: { prefix: "corpus", start: 0, end: 1_000_000 },
      })._unsafeUnwrap();
      if (buildIdentitySidecar(probe)._unsafeUnwrap().tables.length === 0) return;

      // Pin metadata so equality measures identity leakage rather than the clock.
      source.package.properties = {
        ...source.package.properties,
        created: new Date("2000-01-01T00:00:00.000Z"),
        modified: new Date("2000-01-01T00:00:00.000Z"),
      };
      const serialized = await serializeDocumentToDocx(source);
      await assertProperty(
        fc.asyncProperty(identityLeaseArbitrary, async (suppliedLease) => {
          const assigned = assignTableIdentities(source, suppliedLease)._unsafeUnwrap();
          const sidecar = buildIdentitySidecar(assigned)._unsafeUnwrap();
          const parsed = parseIdentitySidecar(stringifyIdentitySidecar(sidecar))._unsafeUnwrap();
          const restored = restoreIdentitySidecar(source, parsed)._unsafeUnwrap();
          expect(buildIdentitySidecar(restored)._unsafeUnwrap()).toEqual(sidecar);
          expect(restored).toEqual(assigned);
          expect(assignTableIdentities(assigned, suppliedLease)._unsafeUnwrap()).toEqual(assigned);
          const [assignedBytes, restoredBytes] = await Promise.all([
            serializeDocumentToDocx(assigned),
            serializeDocumentToDocx(restored),
          ]);
          expect(new Uint8Array(assignedBytes)).toEqual(new Uint8Array(serialized));
          expect(new Uint8Array(restoredBytes)).toEqual(new Uint8Array(serialized));
        }),
        {
          id: `corpus-sidecar-and-serialization:${path.relative(fixtureRoots[0]!, filename)}`,
          numRuns: 5,
        },
      );
    },
    propertyTestTimeout(60_000),
  );
});

const invalidIdentityCases = {
  invalidSidecar: (variant: number) =>
    parseIdentitySidecar(
      ["{", "null", '{"version":2,"tables":[]}', '{"version":1,"tables":[{}]}'][variant % 4]!,
    )._unsafeUnwrapErr(),
  structureMismatch: (variant: number) => {
    const sidecar = buildIdentitySidecar(
      assignTableIdentities(document(), lease)._unsafeUnwrap(),
    )._unsafeUnwrap();
    if (variant % 2 === 0) sidecar.tables.pop();
    else sidecar.tables[0]!.cellIds[0]!.pop();
    return restoreIdentitySidecar(document(), sidecar)._unsafeUnwrapErr();
  },
  invalidLease: (variant: number) => {
    const invalid = structuredClone(lease);
    switch (variant % 4) {
      case 0:
        invalid.paragraph.start = 0;
        break;
      case 1:
        invalid.table.prefix = "INVALID";
        break;
      case 2:
        invalid.row.end = -1;
        break;
      case 3:
        invalid.cell.start = 0.5;
        break;
    }
    return validateDocumentIdLease(invalid)._unsafeUnwrapErr();
  },
  overlappingLease: () => validateDocumentIdLease(lease, [lease])._unsafeUnwrapErr(),
  outsideLease: (variant: number) =>
    validateLeasedStructuralId(
      "table",
      ["t_other_a", "t_abc_q", "t_abc_aa"][variant % 3]!,
      lease.table,
      new Set(),
    )._unsafeUnwrapErr(),
  idCollision: () =>
    validateLeasedStructuralId(
      "table",
      "t_abc_a",
      lease.table,
      new Set(["t_abc_a"]),
    )._unsafeUnwrapErr(),
  duplicatePath: () => {
    const sidecar = buildIdentitySidecar(
      assignTableIdentities(document(), lease)._unsafeUnwrap(),
    )._unsafeUnwrap();
    sidecar.tables.push(structuredClone(sidecar.tables[0]!));
    return restoreIdentitySidecar(document(), sidecar)._unsafeUnwrapErr();
  },
  wrongIdKind: (variant: number) => {
    const sidecar = buildIdentitySidecar(
      assignTableIdentities(document(), lease)._unsafeUnwrap(),
    )._unsafeUnwrap();
    switch (variant % 3) {
      case 0:
        sidecar.tables[0]!.tableId = "r_abc_a";
        break;
      case 1:
        sidecar.tables[0]!.rowIds[0] = "c_abc_a";
        break;
      case 2:
        sidecar.tables[0]!.cellIds[0]![0] = "t_abc_a";
        break;
    }
    return restoreIdentitySidecar(document(), sidecar)._unsafeUnwrapErr();
  },
} satisfies Record<IdentityValidationReason, (variant: number) => IdentityValidationError>;

describe("typed identity validation reasons", () => {
  test.each(Object.values(IDENTITY_VALIDATION_REASONS))(
    "%s is surfaced as an IdentityValidationError",
    (reason) => {
      assertProperty(
        fc.property(fc.nat({ max: 100 }), (variant) => {
          const error = invalidIdentityCases[reason](variant);
          expect(error).toBeInstanceOf(IdentityValidationError);
          expect(error.reason).toBe(reason);
          expect(error.message.length).toBeGreaterThan(0);
        }),
        {
          id: `identity-validation-reason:${reason}`,
          numRuns: 20,
          examples: [[0], [1], [2], [3]],
        },
      );
    },
    propertyTestTimeout(10_000),
  );
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
