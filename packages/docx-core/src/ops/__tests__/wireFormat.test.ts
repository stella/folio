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
import type { Paragraph } from "../../model/document";
import { DOCUMENT_OP_SCHEMA_VERSION, DOCUMENT_OP_TYPES } from "../types";
import { envelopes, wireFixturePath } from "./wireFixtures";
import {
  BATCH_REJECTION_REASONS,
  BATCH_WIRE_OP_TYPES,
  TABLE_EXCLUSIVE_OP_TYPES,
  validateDocumentBatch,
} from "../sequencing/envelope";
import { envelopeFixtures } from "../sequencing/__tests__/envelopeFixtures";

/** Paragraph reviews have one OOXML record; run-property review arrays have a separate contract. */
const expectParagraphReviewCardinality = (value: unknown): void => {
  if (Array.isArray(value)) {
    for (const entry of value) expectParagraphReviewCardinality(entry);
    return;
  }
  if (value === null || typeof value !== "object") return;
  if ("type" in value && value.type === "paragraph" && "propertyChanges" in value) {
    expect(Array.isArray(value.propertyChanges)).toBe(true);
    if (Array.isArray(value.propertyChanges))
      expect(value.propertyChanges.length).toBeLessThanOrEqual(1);
  }
  for (const entry of Object.values(value)) expectParagraphReviewCardinality(entry);
};

test("the wire cardinality guard rejects two property reviews on an identified paragraph", () => {
  const valid = {
    type: "paragraph",
    paraId: "00000001",
    content: [],
    propertyChanges: [
      { type: "paragraphPropertyChange", info: { id: 7, author: "Original reviewer" } },
    ],
  } satisfies Paragraph;
  expect(() => expectParagraphReviewCardinality(valid)).not.toThrow();
  const invalid = {
    ...valid,
    propertyChanges: [
      ...valid.propertyChanges,
      { type: "paragraphPropertyChange", info: { id: 8, author: "Latest reviewer" } },
    ],
  } satisfies Paragraph;
  expect(() => expectParagraphReviewCardinality(invalid)).toThrow();
});

test("every operation kind and its inverse keep their persisted JSON", async () => {
  const written = JSON.stringify(envelopes(), null, 2);
  if (process.env["UPDATE_OP_WIRE_FIXTURE"] === "1") {
    await Bun.write(wireFixturePath, `${written}\n`);
  }
  // The JSON as persisted, compared as data: the fixture's layout is the formatter's.
  const pinned: unknown = await Bun.file(wireFixturePath).json();
  expectParagraphReviewCardinality(pinned);
  const generated = envelopes();
  expectParagraphReviewCardinality(generated);
  for (const { op } of generated) {
    if (op.type !== DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW) continue;
    expect(op.expected.propertyChanges?.length ?? 0).toBeLessThanOrEqual(1);
    expect(op.review.propertyChanges?.length ?? 0).toBeLessThanOrEqual(1);
  }
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

test("historical wire journal shapes read through schemas 4 and 5 and classify every refusal", async () => {
  const journal: unknown = await Bun.file(
    new URL("./__fixtures__/ops-v4.json", import.meta.url),
  ).json();
  if (!Array.isArray(journal)) throw new Error("Historical fixture must be an envelope array.");
  const supported = new Set(BATCH_WIRE_OP_TYPES);
  const observed = {
    supported: new Set<string>(),
    unsupportedPayload: new Set<string>(),
    exclusive: new Set<string>(),
    unsupported: new Set<string>(),
  };
  let processed = 0;
  for (const entry of journal) {
    const envelope: unknown = entry;
    if (
      typeof envelope !== "object" ||
      envelope === null ||
      !("schema" in envelope) ||
      !("op" in envelope)
    )
      throw new Error("Malformed historical envelope.");
    expect(envelope.schema).toBe(4);
    const op = envelope.op;
    if (typeof op !== "object" || op === null || !("type" in op) || typeof op.type !== "string")
      throw new Error("Historical operation needs a discriminator.");
    const kind = Object.values(DOCUMENT_OP_TYPES).find((type) => type === op.type);
    if (!kind) throw new Error(`Historical fixture contains an unknown operation ${op.type}.`);
    const result = validateDocumentBatch({ ...envelopeFixtures[0], schema: 5, ops: [op] });
    const fromSchema4 = validateDocumentBatch({
      ...envelopeFixtures[0],
      schema: envelope.schema,
      ops: [op],
    });
    if (result.isOk()) {
      if (fromSchema4.isErr()) throw fromSchema4.error;
      expect(fromSchema4.value).toStrictEqual(result.value);
    } else {
      if (fromSchema4.isOk()) throw new Error("Historical schema classification must agree.");
      expect(fromSchema4.error.reason).toBe(result.error.reason);
    }
    if (supported.has(kind)) {
      if (result.isErr()) {
        // Old journals include optional shapes outside the sequenced decoder contract.
        expect(result.error.reason).toBe(BATCH_REJECTION_REASONS.INVALID_OPERATION);
        observed.unsupportedPayload.add(kind);
      } else {
        expect(result.value.schema).toBe(DOCUMENT_OP_SCHEMA_VERSION);
        expect(result.value.ops).toStrictEqual([op]);
        observed.supported.add(kind);
      }
    } else {
      if (result.isOk()) throw new Error(`Unsupported historical operation ${kind} must refuse.`);
      if (Object.hasOwn(TABLE_EXCLUSIVE_OP_TYPES, kind)) {
        expect(result.error.reason).toBe(BATCH_REJECTION_REASONS.TABLE_REQUIRES_EXCLUSIVE_EDIT);
        observed.exclusive.add(kind);
      } else {
        expect(result.error.reason).toBe(BATCH_REJECTION_REASONS.INVALID_OPERATION);
        observed.unsupported.add(kind);
      }
    }
    processed++;
  }
  expect(processed).toBe(journal.length);
  expect([...new Set([...observed.supported, ...observed.unsupportedPayload])].toSorted()).toEqual(
    [...BATCH_WIRE_OP_TYPES].toSorted(),
  );
  expect([...observed.unsupportedPayload].toSorted()).toEqual(
    ["insertContent", "joinBlocks", "splitBlock", "deleteRange", "setRunProps"].toSorted(),
  );
  expect([...observed.exclusive].toSorted()).toEqual(
    [
      "insertTable",
      "deleteTable",
      "insertRow",
      "deleteRow",
      "setTableRows",
      "setContainerBlocks",
    ].toSorted(),
  );
  expect([...observed.unsupported].toSorted()).toEqual(
    [
      "joinInline",
      "replaceBlocks",
      "replaceInline",
      "setParagraphReview",
      "splitInline",
    ].toSorted(),
  );
});
