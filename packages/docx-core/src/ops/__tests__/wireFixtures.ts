/** Pure wire-fixture construction, shared by snapshot tests and generation. */
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

export const wireFixturePath = path.join(
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
    blocks: [{ type: "paragraph", paraId: "0000000D", content: [] }],
  },
  {
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: ["0000000D"],
    revision: stamp(81),
  },
  {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [81],
    decision: REVISION_DECISIONS.ACCEPT,
  },
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

export const envelopes = (): DocumentOpEnvelope[] => {
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
      newIds: { revision: [201, 204] },
    },
    {
      type: DOCUMENT_OP_TYPES.DELETE_ROW,
      story: OP_STORIES.MAIN,
      blockId: "00000021",
      revision: stamp(202),
      newIds: { revision: [203, 205] },
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
  const trackedTableCases = [
    {
      document,
      op: {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story: OP_STORIES.MAIN,
        at: { type: "before", blockId: "00000002" },
        table: { type: "table", rows: [row("00000040"), row("00000041")] },
        revision: stamp(220),
        newIds: { revision: [221, 222, 223, 224, 225] },
      },
    },
    {
      document: {
        package: {
          document: {
            content: [{ type: "table", rows: [row("00000042"), row("00000043")] }, second],
          },
        },
      },
      op: {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "00000042",
        revision: stamp(230),
        newIds: { revision: [231, 232, 233, 234, 235] },
      },
    },
    {
      document: { package: { document: { content: [second] } } },
      op: {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: "00000002" },
        table: { type: "table", rows: [row("00000044")] },
        terminal: { beforeBlockId: "00000045" },
        revision: stamp(240),
        newIds: { revision: [241, 242, 243] },
      },
    },
  ] as const satisfies readonly { document: Document; op: DocumentOp }[];
  for (const { document: trackedTableDocument, op } of trackedTableCases) {
    const result = applyDocumentOp(trackedTableDocument, op);
    if (result.isErr()) throw result.error;
    out.push(toOpEnvelope(op));
    for (const inverse of result.value.inverse) out.push(toOpEnvelope(inverse));
    // Persist the physical ids minted for every row, cell mark and inline wrapper.
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const resolution = {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: result.value.revisions,
        decision,
      } as const satisfies DocumentOp;
      const resolved = applyDocumentOp(result.value.document, resolution);
      if (resolved.isErr()) throw resolved.error;
      out.push(toOpEnvelope(resolution));
      for (const inverse of resolved.value.inverse) out.push(toOpEnvelope(inverse));
    }
  }
  const terminalOp = {
    type: DOCUMENT_OP_TYPES.INSERT_TABLE,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: "00000002" },
    table: { type: "table", rows: [row("00000046")] },
    terminal: { beforeBlockId: "00000047" },
  } as const satisfies DocumentOp;
  const terminalResult = applyDocumentOp(
    { package: { document: { content: [second] } } },
    terminalOp,
  );
  if (terminalResult.isErr()) throw terminalResult.error;
  out.push(toOpEnvelope(terminalOp));
  for (const inverse of terminalResult.value.inverse) out.push(toOpEnvelope(inverse));
  const lifecycleOps = [
    {
      type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
      sectionIndex: 0,
      story: { kind: "footer", rId: "rIdFooter" },
      referenceType: "first",
      content: [{ type: "paragraph", paraId: "000000A1", content: [] }],
    },
    {
      type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
      sectionIndex: 0,
      patch: { footnotePr: { numStart: 2 } },
    },
    {
      type: DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
      sectionIndex: 0,
      story: { kind: "footer", rId: "rIdFooter" },
      referenceType: "first",
    },
    {
      type: DOCUMENT_OP_TYPES.ADD_NOTE,
      at: at("00000002", 0),
      note: {
        type: "footnote",
        id: 1,
        content: [{ type: "paragraph", paraId: "000000A2", content: [] }],
      },
    },
    {
      type: DOCUMENT_OP_TYPES.REMOVE_NOTE,
      at: at("00000002", 0),
      story: { kind: "footnote", id: 1 },
    },
  ] as const satisfies readonly DocumentOp[];
  let lifecycleDocument: Document = { package: { document: { content: [second] } } };
  for (const op of lifecycleOps) {
    const applied = applyDocumentOp(lifecycleDocument, op);
    if (applied.isErr()) throw applied.error;
    out.push(toOpEnvelope(op));
    for (const inverse of applied.value.inverse) out.push(toOpEnvelope(inverse));
    lifecycleDocument = applied.value.document;
  }
  return out;
};

if (import.meta.main) {
  await Bun.write(wireFixturePath, `${JSON.stringify(envelopes(), null, 2)}\n`);
}
