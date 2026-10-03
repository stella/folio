import { expect, test } from "bun:test";
import { applyDocumentOps } from "../apply";
import type { Document } from "../../model/document";
import { DOCUMENT_OP_SCHEMA_VERSION, OP_STORIES, type DocumentOp } from "../types";
import { envelopes } from "../__tests__/wireFixtures";
import {
  BATCH_WIRE_OP_TYPES,
  MAX_BATCH_WIRE_BYTES,
  parseDocumentBatch,
  validateDocumentBatch,
  validateSequencedBatch,
} from "./envelope";
import { envelopeFixtures, sequencedFixture } from "./__tests__/envelopeFixtures";

const fixture = envelopeFixtures[0];
const withOp = (op: unknown) => ({ ...fixture, ops: [op] });

test("unsupported operation families refuse production-shaped wire fixtures", () => {
  const supported = new Set(BATCH_WIRE_OP_TYPES);
  let refused = 0;
  for (const { op } of envelopes()) {
    if (supported.has(op.type)) continue;
    refused += 1;
    expect(validateDocumentBatch(withOp(op)).isErr()).toBe(true);
  }
  expect(refused).toBeGreaterThan(0);
});

test("batch wire fixtures pin every supported decoder kind and JSON roundtrip", async () => {
  const pinned: unknown = await Bun.file(
    new URL("./__fixtures__/batches-v5.json", import.meta.url),
  ).json();
  expect(JSON.parse(JSON.stringify([...envelopeFixtures, sequencedFixture]))).toEqual(pinned);
  const kinds = new Set(envelopeFixtures.flatMap(({ ops }) => ops.map(({ type }) => type)));
  expect([...kinds].toSorted()).toEqual([...BATCH_WIRE_OP_TYPES].toSorted());
  for (const batch of envelopeFixtures) {
    const decoded = parseDocumentBatch(JSON.stringify(batch));
    expect(decoded.isOk()).toBe(true);
    if (decoded.isErr()) throw decoded.error;
    expect(decoded.value).toStrictEqual(batch);
    expect(JSON.stringify(decoded.value)).toBe(JSON.stringify(batch));
  }
  expect(JSON.parse(JSON.stringify(sequencedFixture))).toStrictEqual(sequencedFixture);
  expect(validateSequencedBatch(JSON.parse(JSON.stringify(sequencedFixture))).isOk()).toBe(true);
  expect(validateSequencedBatch(fixture).isErr()).toBe(true);
  expect(
    validateSequencedBatch({
      ...sequencedFixture,
      effects: sequencedFixture.effects.map((effect, index) =>
        index === 0 ? { type: "joinBlocks", firstLength: 3 } : effect,
      ),
    }).isErr(),
  ).toBe(true);
  expect(validateSequencedBatch({ ...sequencedFixture, effects: [] }).isErr()).toBe(true);
});

test("unknown envelopes refuse invalid identities, revisions, schemas and keys", () => {
  for (const batch of [
    null,
    [],
    {},
    { ...fixture, schema: 0 },
    { ...fixture, schema: DOCUMENT_OP_SCHEMA_VERSION - 1 },
    { ...fixture, actor: "" },
    { ...fixture, opId: "" },
    { ...fixture, baseRev: -1 },
    { ...fixture, baseRev: Number.MAX_SAFE_INTEGER + 1 },
    { ...fixture, revision: 1.5 },
    { ...fixture, ops: {} },
    { ...fixture, extra: true },
    { ...fixture, ops: [null] },
  ])
    expect(validateDocumentBatch(batch).isErr()).toBe(true);
  expect(parseDocumentBatch("{invalid").isErr()).toBe(true);
});

test("each supported operation validates every supplied field and nested payload", () => {
  for (const op of fixture.ops) {
    expect(validateDocumentBatch(withOp(op)).isOk()).toBe(true);
    expect(validateDocumentBatch(withOp({ ...op, extra: true })).isErr()).toBe(true);
    for (const key of Object.keys(op)) {
      expect(validateDocumentBatch(withOp({ ...op, [key]: 12345 })).isErr()).toBe(true);
    }
  }
  const malformed = [
    {
      type: "insertText",
      at: { story: "header", blockId: "00000001", offset: 0 },
      text: "A",
      runProps: "inherit",
    },
    {
      type: "insertText",
      at: { story: "main", blockId: "0000000a", offset: 0 },
      text: "A",
      runProps: "inherit",
    },
    {
      type: "insertText",
      at: { story: "main", blockId: "00000001", offset: 0, zeroWidthBefore: -1 },
      text: "A",
      runProps: "inherit",
    },
    {
      type: "insertText",
      at: { story: "main", blockId: "00000001", offset: 0 },
      text: "A",
      runProps: { bold: "true" },
    },
    {
      type: "insertText",
      at: { story: "main", blockId: "00000001", offset: 0 },
      text: "A",
      runProps: "inherit",
      revision: { id: 1, author: "A", date: 0 },
    },
    {
      type: "insertContent",
      at: { story: "main", blockId: "00000001", offset: 0 },
      slice: {
        content: [{ type: "run", content: [{ type: "text", text: 0 }] }],
        openStart: 0,
        openEnd: 0,
      },
    },
    {
      type: "insertBlocks",
      story: "main",
      at: { type: "before", blockId: "00000001" },
      blocks: [{ type: "paragraph", paraId: "00000002", content: [null] }],
    },
    { type: "unknown" },
  ];
  for (const op of malformed) expect(validateDocumentBatch(withOp(op)).isErr()).toBe(true);
});

test("wire validation refuses non-JSON data without executing getters", () => {
  const cycle: Record<string, unknown> = {};
  cycle.self = cycle;
  let getterRead = false;
  const getter = Object.defineProperty({}, "ops", {
    enumerable: true,
    get: () => {
      getterRead = true;
      return [];
    },
  });
  const sparse = Array(1);
  for (const value of [
    cycle,
    getter,
    sparse,
    { ...fixture, actor: undefined },
    { ...fixture, actor: new Date() },
    { ...fixture, baseRev: NaN },
    { ...fixture, baseRev: -0 },
    { ...fixture, baseRev: Infinity },
    { ...fixture, actor: () => "actor" },
    { ...fixture, actor: 1n },
    { ...fixture, [Symbol("hidden")]: true },
  ]) {
    expect(validateDocumentBatch(value).isErr()).toBe(true);
  }
  expect(getterRead).toBe(false);
  let deep: unknown = "leaf";
  for (let index = 0; index < 80; index += 1) deep = { child: deep };
  expect(validateDocumentBatch(deep).isErr()).toBe(true);
});

test("wire byte limit counts UTF-8 bytes and parser input", () => {
  expect(parseDocumentBatch(" ".repeat(MAX_BATCH_WIRE_BYTES + 1)).isErr()).toBe(true);
  expect(
    validateDocumentBatch({ ...fixture, actor: "😀".repeat(MAX_BATCH_WIRE_BYTES / 3) }).isErr(),
  ).toBe(true);
});

test("actual text inverses normalize absent optional fields and remain decodable", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [{ type: "run", content: [{ type: "text", text: "abc" }] }],
          },
        ],
      },
    },
  };
  const at = (offset: number) => ({ story: OP_STORIES.MAIN, blockId: "00000001", offset });
  const operations = [
    { type: "insertText", at: at(1), text: "X", runProps: "inherit" },
    { type: "deleteRange", from: at(0), to: at(2) },
  ] as const satisfies readonly DocumentOp[];
  for (const op of operations) {
    const applied = applyDocumentOps(document, [op]);
    if (applied.isErr()) throw applied.error;
    const inverseBatch = validateDocumentBatch({ ...fixture, ops: applied.value.inverse });
    expect(inverseBatch.isOk()).toBe(true);
    if (inverseBatch.isErr()) throw inverseBatch.error;
    const undone = applyDocumentOps(applied.value.document, inverseBatch.value.ops);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(document);
  }
});
