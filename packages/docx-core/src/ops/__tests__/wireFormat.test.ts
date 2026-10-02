/**
 * The persisted form of operations in the current schema version.
 *
 * Journaled operations are read back years after they were written, so their
 * JSON is pinned: one envelope per operation kind and per inverse, produced
 * from a fixed document. A change to an operation's fields, or to a model
 * record an operation embeds, shows up here as a diff of the fixture; it
 * needs a new schema version and a migration, not an updated fixture.
 *
 * `bun packages/docx-core/src/ops/__tests__/wireFixtures.ts` generates the current fixture.
 */

import { expect, test } from "bun:test";
import { DOCUMENT_OP_SCHEMA_VERSION, DOCUMENT_OP_TYPES } from "../types";
import { envelopes, wireFixturePath } from "./wireFixtures";

test("every operation kind and its inverse keep their persisted JSON", async () => {
  const written = JSON.stringify(envelopes(), null, 2);
  if (process.env["UPDATE_OP_WIRE_FIXTURE"] === "1") {
    await Bun.write(wireFixturePath, `${written}\n`);
  }
  // The JSON as persisted, compared as data: the fixture's layout is the formatter's.
  const pinned: unknown = await Bun.file(wireFixturePath).json();
  expect(JSON.parse(written)).toStrictEqual(pinned);
  const kinds = new Set(envelopes().map(({ op }) => op.type));
  expect([...kinds].toSorted()).toEqual(Object.values(DOCUMENT_OP_TYPES).toSorted());
  expect(envelopes().every(({ schema }) => schema === DOCUMENT_OP_SCHEMA_VERSION)).toBe(true);
});

test("semantic table wire fixtures exercise geometry, ids, patches and tracked payloads", () => {
  const fixtures = envelopes().map(({ op }) => op);
  const fields = {
    insertColumn: {
      type: true,
      story: true,
      blockId: true,
      column: true,
      width: true,
      newBlockIds: true,
      revision: true,
      newIds: true,
    },
    deleteColumn: {
      type: true,
      story: true,
      blockId: true,
      column: true,
      revision: true,
      newIds: true,
    },
    mergeCells: {
      type: true,
      story: true,
      blockId: true,
      top: true,
      bottom: true,
      left: true,
      right: true,
      newBlockIds: true,
      revision: true,
      newIds: true,
    },
    splitCell: {
      type: true,
      story: true,
      blockId: true,
      newBlockIds: true,
      revision: true,
      newIds: true,
    },
    setTableGrid: {
      type: true,
      story: true,
      blockId: true,
      columnWidths: true,
      revision: true,
      newIds: true,
    },
    setCellProps: {
      type: true,
      story: true,
      blockId: true,
      patch: true,
      revision: true,
      newIds: true,
    },
    setRowProps: {
      type: true,
      story: true,
      blockId: true,
      patch: true,
      revision: true,
      newIds: true,
    },
    setTableProps: {
      type: true,
      story: true,
      blockId: true,
      patch: true,
      revision: true,
      newIds: true,
    },
  } as const satisfies {
    [Kind in import("../types").TableEditOp["type"]]: Record<
      keyof Extract<import("../types").TableEditOp, { type: Kind }>,
      true
    >;
  };
  for (const [kind, declared] of Object.entries(fields)) {
    const matching = fixtures.filter((op) => op.type === kind);
    expect(matching.length).toBeGreaterThan(0);
    const exercised = new Set(matching.flatMap((op) => Object.keys(op)));
    expect([...exercised].toSorted()).toEqual(Object.keys(declared).toSorted());
  }
  const restorations = fixtures.filter((op) => op.type === DOCUMENT_OP_TYPES.SET_TABLE);
  expect(restorations.length).toBeGreaterThan(0);
  for (const op of restorations) {
    expect(op.expected.type).toBe("table");
    expect(op.table.type).toBe("table");
  }
});
