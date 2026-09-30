/**
 * The persisted form of operations in the current schema version.
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

import type { Document, Paragraph, Table, TableRow } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpEnvelope,
  toOpEnvelope,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const FIXTURE = path.join(
  import.meta.dir,
  "__fixtures__",
  `ops-v${DOCUMENT_OP_SCHEMA_VERSION}.json`,
);

const stamp = (id: number) => ({
  id,
  author: "Reviewer",
  date: "2026-02-03T04:05:06Z",
  initials: "R",
});

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
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId: "00000001" },
    blocks: [{ type: "paragraph", paraId: "0000000C", content: [] }],
    revision: stamp(80),
  },
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
    blockId: "00000001",
    nextBlockId: "00000002",
  },
  {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: at("0000000A", 1),
    text: "Y",
    runProps: "inherit",
    newIds: { revision: [61] },
    revision: stamp(60),
  },
  {
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: at("0000000A", 3),
    to: at("0000000A", 5),
    newIds: { revision: [63, 64] },
    revision: stamp(62),
  },
  {
    type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
    from: at("00000002", 0),
    to: at("00000002", 2),
    patch: { underline: { style: "single" } },
    newIds: { revision: [66, 67] },
    revision: stamp(65),
  },
  {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: "00000002",
    patch: { keepNext: true },
    revision: stamp(68),
  },
  {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: at("0000000A", 2),
    newBlockId: "0000000B",
    newIds: { revision: [70, 71] },
    revision: stamp(69),
  },
  {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "0000000A",
    nextBlockId: "00000002",
    newIds: { revision: [73] },
    revision: stamp(72),
  },
  {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [60, 62, 65, 68],
    decision: REVISION_DECISIONS.ACCEPT,
  },
  {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [69, 72],
    decision: REVISION_DECISIONS.REJECT,
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
  const row = (paraId: string): TableRow => ({
    type: "tableRow",
    cells: [
      {
        type: "tableCell",
        content: [
          {
            type: "paragraph",
            paraId,
            content: [{ type: "run", content: [{ type: "text", text: "Cell" }] }],
          },
        ],
      },
    ],
  });
  let rowDocument: Document = {
    package: {
      document: {
        content: [
          { type: "table", rows: [row("00000020"), row("00000021")] },
          { type: "paragraph", paraId: "00000022", content: [] },
        ],
      },
    },
  };
  const rowOps = [
    {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000020",
      at: 1,
      row: row("00000023"),
      revision: stamp(200),
      newIds: { revision: [201] },
    },
    {
      type: DOCUMENT_OP_TYPES.DELETE_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000021",
      revision: stamp(202),
      newIds: { revision: [203] },
    },
  ] as const satisfies readonly DocumentOp[];
  for (const op of rowOps) {
    const result = applyDocumentOp(rowDocument, op);
    if (result.isErr()) throw result.error;
    out.push(toOpEnvelope(op));
    for (const inverse of result.value.inverse) out.push(toOpEnvelope(inverse));
    rowDocument = result.value.document;
  }
  const table: Table = { type: "table", rows: [row("00000031")] };
  let tableDocument: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000030", content: [] },
          table,
          { type: "paragraph", paraId: "00000032", content: [] },
        ],
      },
    },
  };
  const tableOps = [
    {
      type: DOCUMENT_OP_TYPES.INSERT_TABLE,
      story: OP_STORIES.MAIN,
      at: { type: "before", blockId: "00000030" },
      table: { type: "table", rows: [row("00000033")] },
    },
    {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: OP_STORIES.MAIN,
      blockId: "00000031",
      expected: table,
    },
  ] as const satisfies readonly DocumentOp[];
  for (const op of tableOps) {
    const result = applyDocumentOp(tableDocument, op);
    if (result.isErr()) throw result.error;
    out.push(toOpEnvelope(op));
    for (const inverse of result.value.inverse) out.push(toOpEnvelope(inverse));
    tableDocument = result.value.document;
  }
  return out;
};

test("every operation kind and its inverse keep their persisted JSON", async () => {
  const written = JSON.stringify(envelopes(), null, 2);
  if (process.env["UPDATE_OP_WIRE_FIXTURE"] === "1") {
    await Bun.write(FIXTURE, `${written}\n`);
  }
  // The JSON as persisted, compared as data: the fixture's layout is the formatter's.
  const pinned: unknown = await Bun.file(FIXTURE).json();
  expect(JSON.parse(written)).toStrictEqual(pinned);
  const kinds = new Set(envelopes().map(({ op }) => op.type));
  expect([...kinds].toSorted()).toEqual(Object.values(DOCUMENT_OP_TYPES).toSorted());
  expect(envelopes().every(({ schema }) => schema === DOCUMENT_OP_SCHEMA_VERSION)).toBe(true);
});
