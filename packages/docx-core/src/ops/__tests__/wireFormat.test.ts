/**
 * The persisted form of schema-version-1 operations.
 *
 * Journaled operations are read back years after they were written, so their
 * JSON is pinned: one envelope per operation kind and per inverse, produced
 * from a fixed document. A change to an operation's fields, or to a model
 * record an operation embeds, shows up here as a diff of the fixture; it
 * needs a new schema version and a migration, not an updated fixture.
 *
 * `UPDATE_OP_WIRE_FIXTURE=1` writes the fixture again.
 */

import { expect, test } from "bun:test";
import path from "node:path";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpEnvelope,
  toOpEnvelope,
  OP_STORIES,
} from "../types";

const FIXTURE = path.join(import.meta.dir, "__fixtures__", "ops-v1.json");

const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });

const first: Paragraph = {
  type: "paragraph",
  paraId: "00000001",
  formatting: { alignment: "center" },
  preservedAttributes: [{ name: "rsidR", value: "00A1B2C3" }],
  content: [
    { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "Alpha " }] },
    { type: "bookmarkStart", id: 1, name: "_Ref1" },
    {
      type: "insertion",
      info: { id: 5, author: "A", date: "2026-01-02T03:04:05Z" },
      content: [{ type: "run", content: [{ type: "text", text: "beta" }] }],
    },
    { type: "bookmarkEnd", id: 1 },
    { type: "run", content: [{ type: "text", text: " gamma" }, { type: "tab" }] },
  ],
};

const second: Paragraph = {
  type: "paragraph",
  paraId: "00000002",
  content: [{ type: "run", content: [{ type: "text", text: "Delta" }] }],
};

const document: Document = { package: { document: { content: [first, second] } } };

/** Each operation, applied in turn to what the previous one produced. */
const OPS: readonly DocumentOp[] = [
  {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: at("00000001", 3),
    text: "X",
    runProps: { italic: true },
  },
  { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at("00000001", 5), to: at("00000001", 9) },
  { type: DOCUMENT_OP_TYPES.SPLIT_INLINE, at: at("00000002", 2), depth: 2 },
  { type: DOCUMENT_OP_TYPES.JOIN_INLINE, at: at("00000002", 2), depth: 2 },
  {
    type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
    from: at("00000002", 1),
    to: at("00000002", 3),
    patch: { bold: true, italic: null },
  },
  {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: "00000002",
    patch: { alignment: "end" },
  },
  {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: at("00000001", 6),
    newBlockId: "0000000A",
    newIds: { revision: [50] },
  },
  {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "0000000A",
    nextBlockId: "00000002",
  },
];

const envelopes = (): DocumentOpEnvelope[] => {
  const out: DocumentOpEnvelope[] = [];
  let current = document;
  for (const op of OPS) {
    const result = applyDocumentOp(current, op);
    if (result.isErr()) throw result.error;
    out.push(toOpEnvelope(op));
    for (const inverse of result.value.inverse) out.push(toOpEnvelope(inverse));
    current = result.value.document;
  }
  const [paragraph] = current.package.document.content;
  if (paragraph?.type === "paragraph") {
    const replacement: Paragraph = { ...paragraph, content: [] };
    const result = applyDocumentOp(current, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: [paragraph],
      blocks: [replacement],
    });
    if (result.isErr()) throw result.error;
    out.push(
      toOpEnvelope({
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: OP_STORIES.MAIN,
        expected: [paragraph],
        blocks: [replacement],
      }),
    );
    for (const inverse of result.value.inverse) out.push(toOpEnvelope(inverse));
  }
  return out;
};

test("every operation kind and its inverse keep their persisted JSON", async () => {
  const written = `${JSON.stringify(envelopes(), null, 2)}\n`;
  if (process.env["UPDATE_OP_WIRE_FIXTURE"] === "1") {
    await Bun.write(FIXTURE, written);
  }
  expect(written).toBe(await Bun.file(FIXTURE).text());
  const kinds = new Set(envelopes().map(({ op }) => op.type));
  expect([...kinds].toSorted()).toEqual(Object.values(DOCUMENT_OP_TYPES).toSorted());
  expect(envelopes().every(({ schema }) => schema === DOCUMENT_OP_SCHEMA_VERSION)).toBe(true);
});
