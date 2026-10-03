import { expect, test } from "bun:test";
import { deepStrictEqual, notDeepStrictEqual } from "node:assert/strict";
import type { Document, Paragraph } from "../../../model/document";
import { applyDocumentOps } from "../../apply";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  toOpEnvelope,
  type DocumentOp,
} from "../../types";
import { restoreDocumentOp } from "../../wire";
import { createClient } from "../client";
import {
  BATCH_REJECTION_REASONS,
  parseDocumentBatch,
  validateDocumentBatch,
  type DocumentBatch,
  type SequencedBatch,
} from "../envelope";
import { createSequencer } from "../sequencer";
import { transformBatch } from "../transform";

const jsonTransport = <T>(value: T): T => JSON.parse(JSON.stringify(value));
const ownUndefined = (record: object, key: string) =>
  Object.defineProperty(record, key, {
    value: undefined,
    enumerable: true,
    configurable: true,
    writable: true,
  });
const documentFixture = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          formatting: { keepNext: true },
          content: [
            {
              type: "run",
              formatting: { italic: false },
              content: [{ type: "text", text: "abc" }],
            },
          ],
        },
      ],
    },
  },
});
const batch = (ops: readonly DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "presence",
  actor: "actor",
  baseRev: 0,
  ops: ops.map((op) => toOpEnvelope(op).op),
});
const position = (offset: number) => ({ story: "main", blockId: "00000001", offset }) as const;
const insertion = {
  type: DOCUMENT_OP_TYPES.INSERT_TEXT,
  at: position(0),
  text: "X",
  runProps: "inherit",
} as const satisfies DocumentOp;

test("batch validation and parsing preserve owned undefined patches through authority and client admission", () => {
  const patch = { keepNext: false };
  ownUndefined(patch, "alignment");
  const runPatch = { italic: true };
  ownUndefined(runPatch, "bold");
  const ops = [
    { type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS, story: "main", blockId: "00000001", patch },
    { type: DOCUMENT_OP_TYPES.SET_RUN_PROPS, from: position(0), to: position(3), patch: runPatch },
  ] satisfies DocumentOp[];
  const incoming = batch(ops);
  const parsed = parseDocumentBatch(JSON.stringify(incoming)).unwrap();
  const validated = validateDocumentBatch(incoming).unwrap();
  for (const candidate of [parsed, validated])
    deepStrictEqual(
      candidate.ops.map((op) => restoreDocumentOp(op).unwrap()),
      ops,
    );
  const original = documentFixture();
  const expected = structuredClone(original);
  const paragraph = expected.package.document.content.at(0);
  const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
  if (
    paragraph?.type !== "paragraph" ||
    !paragraph.formatting ||
    run?.type !== "run" ||
    !run.formatting
  )
    throw new TypeError("Admission fixture disappeared.");
  paragraph.formatting.keepNext = false;
  ownUndefined(paragraph.formatting, "alignment");
  run.formatting.italic = true;
  ownUndefined(run.formatting, "bold");
  const client = createClient(original);
  client.enqueue(parsed).unwrap();
  deepStrictEqual(client.document, expected);
  const sequencer = createSequencer(original);
  const submission = client.nextSubmission();
  if (!submission) throw new TypeError("Admitted client batch did not submit.");
  const ack = sequencer.submit(jsonTransport(submission));
  expect(ack.type).toBe("ack");
  if (ack.type !== "ack") throw ack.reason;
  client.receiveAck(ack);
  const broadcast = sequencer.broadcasts.at(0);
  if (!broadcast) throw new TypeError("Authority did not broadcast.");
  client.receiveBroadcast(jsonTransport(broadcast)).unwrap();
  deepStrictEqual(client.document, expected);
  deepStrictEqual(sequencer.document, expected);
  expect(client.notices).toEqual([]);
  expect(client.pending).toEqual([]);
});

test("admitted paragraph snapshot arrays retain undefined formatting after JSON and exact inverse", () => {
  const paragraph = {
    type: "paragraph",
    paraId: "00000002",
    formatting: { keepNext: false },
    content: [
      { type: "run", formatting: { italic: true }, content: [{ type: "text", text: "snapshot" }] },
    ],
  } satisfies Paragraph;
  ownUndefined(paragraph.formatting, "alignment");
  const run = paragraph.content.at(0);
  if (!run) throw new TypeError("Snapshot run disappeared.");
  ownUndefined(run.formatting, "bold");
  const op = {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: "main",
    at: { type: "before", blockId: "00000001" },
    blocks: [paragraph],
  } as const satisfies DocumentOp;
  notDeepStrictEqual(jsonTransport(op), op);
  const parsed = parseDocumentBatch(JSON.stringify(batch([op]))).unwrap();
  deepStrictEqual(
    parsed.ops.map((item) => restoreDocumentOp(item).unwrap()),
    [op],
  );
  const original = documentFixture();
  const applied = applyDocumentOps(original, parsed.ops).unwrap();
  deepStrictEqual(applied.document.package.document.content.at(0), paragraph);
  const undo = applyDocumentOps(applied.document, jsonTransport(applied.inverse)).unwrap();
  deepStrictEqual(undo.document, original);
});

test.each(["paragraph", "run"] as const)(
  "transform recaptures remaining %s patch metadata after conflict removal",
  (kind) => {
    const patch = { keepNext: true };
    ownUndefined(patch, "alignment");
    const runPatch = { italic: true };
    ownUndefined(runPatch, "bold");
    ownUndefined(patch, "spaceBefore");
    ownUndefined(runPatch, "strike");
    const incomingOp: DocumentOp =
      kind === "paragraph"
        ? { type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS, story: "main", blockId: "00000001", patch }
        : {
            type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
            from: position(0),
            to: position(3),
            patch: runPatch,
          };
    const foreignOp: DocumentOp =
      kind === "paragraph"
        ? {
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: "main",
            blockId: "00000001",
            patch: { alignment: "right" },
          }
        : {
            type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
            from: position(0),
            to: position(3),
            patch: { bold: true },
          };
    const incoming = jsonTransport(batch([incomingOp]));
    const before = structuredClone(incoming);
    const foreign = {
      ...batch([foreignOp]),
      opId: "foreign",
      revision: 1,
    } satisfies SequencedBatch;
    const transformed = transformBatch(incoming, [foreign], { order: "before" }).unwrap();
    const expectedOp: DocumentOp =
      kind === "paragraph"
        ? {
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: "main",
            blockId: "00000001",
            patch: { keepNext: true },
          }
        : {
            type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
            from: position(0),
            to: position(3),
            patch: { italic: true },
          };
    if (expectedOp.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS)
      ownUndefined(expectedOp.patch, "spaceBefore");
    if (expectedOp.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS)
      ownUndefined(expectedOp.patch, "strike");
    deepStrictEqual(
      transformed.ops.map((op) => restoreDocumentOp(jsonTransport(op)).unwrap()),
      [expectedOp],
    );
    expect(
      transformed.ops.some((op) =>
        op.undefinedFields?.some((path) =>
          path.includes(kind === "paragraph" ? "alignment" : "bold"),
        ),
      ),
    ).toBe(false);
    deepStrictEqual(incoming, before);
  },
);

const unproven = [
  {
    type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
    num: { numId: 1, abstractNumId: 1 },
    abstractNum: { abstractNumId: 1, levels: [] },
  },
  { type: DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE, num: { numId: 1, abstractNumId: 1 } },
  {
    type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
    endpoint: { type: "paragraph", blockId: "00000002" },
    properties: { type: "undefined" },
  },
] as const satisfies readonly DocumentOp[];

test("unproven numbering and unrelated section transforms refuse atomically in either direction", () => {
  for (const op of unproven) {
    for (const reverse of [false, true]) {
      const incoming = batch(reverse ? [insertion] : [insertion, op]);
      const foreign = {
        ...batch(reverse ? [op] : [insertion]),
        opId: "foreign",
        revision: 1,
      } satisfies SequencedBatch;
      const before = structuredClone(incoming);
      const foreignBefore = structuredClone(foreign);
      const transformed = transformBatch(incoming, [foreign]);
      expect(transformed.isErr()).toBe(true);
      if (transformed.isOk()) throw new TypeError("Unproven transform pair unexpectedly accepted.");
      expect(transformed.error.reason).toBe(BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR);
      deepStrictEqual(incoming, before);
      deepStrictEqual(foreign, foreignBefore);
    }
  }
});
