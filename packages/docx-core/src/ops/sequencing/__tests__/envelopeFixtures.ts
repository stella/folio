import path from "node:path";
import { DOCUMENT_OP_SCHEMA_VERSION, OP_STORIES } from "../../types";
import { captureDocumentOp } from "../../wire";
import type { DocumentBatch, SequencedBatch } from "../envelope";

const at = (offset: number) => ({ story: OP_STORIES.MAIN, blockId: "00000001", offset });
const revision = { id: 42, author: "Editor", date: "2026-01-01T00:00:00Z" };

export const envelopeFixtures = [
  {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "01JTESTBATCH000000000000001",
    actor: "actor-1",
    baseRev: 0,
    ops: [
      { type: "insertText", at: at(0), text: "A😀", runProps: "inherit" },
      {
        type: "insertContent",
        at: at(3),
        slice: { content: [{ type: "run", content: [{ type: "tab" }] }], openStart: 0, openEnd: 0 },
      },
      { type: "deleteRange", from: at(1), to: at(3), revision },
      {
        type: "setRunProps",
        from: at(0),
        to: at(1),
        patch: { bold: true, italic: null },
        expected: { bold: null, italic: true },
        whenEmpty: "omit",
      },
      {
        type: "setParagraphProps",
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        patch: { alignment: "center", spaceBefore: null },
      },
      {
        type: "splitBlock",
        at: at(1),
        newBlockId: "00000002",
        newHalf: "second",
        newIds: { revision: [43], control: [9] },
      },
      {
        type: "joinBlocks",
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
        survivor: "first",
      },
      {
        type: "insertBlocks",
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: "00000001" },
        blocks: [
          {
            type: "paragraph",
            paraId: "00000003",
            content: [
              { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "New" }] },
            ],
          },
        ],
      },
      { type: "deleteBlocks", story: OP_STORIES.MAIN, blockIds: ["00000003"] },
      { type: "resolveRevision", story: OP_STORIES.MAIN, revisionIds: [42], decision: "accept" },
    ],
  },
  {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "01JTESTBATCH000000000000002",
    actor: "actor-2",
    baseRev: 1,
    revision: 2,
    ops: [
      {
        type: "insertText",
        at: { ...at(1), zeroWidthBefore: 2 },
        text: "Z",
        runProps: { bold: true, language: { val: "cs-CZ" } },
        revision,
      },
    ],
  },
  {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "01JTESTBATCH000000000000004",
    actor: "actor-1",
    baseRev: 0,
    ops: [
      captureDocumentOp({
        type: "setRunProps",
        from: at(0),
        to: at(1),
        patch: { bold: undefined },
        expected: { bold: null },
      }),
      captureDocumentOp({
        type: "setParagraphProps",
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        patch: { alignment: undefined },
        expected: { alignment: null },
      }),
    ],
  },
] as const satisfies readonly DocumentBatch[];

export const sequencedFixture = {
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "01JTESTBATCH000000000000003",
  actor: "actor-1",
  baseRev: 2,
  revision: 3,
  ops: [
    { type: "splitBlock", at: at(3), newBlockId: "00000004" },
    { type: "joinBlocks", story: OP_STORIES.MAIN, blockId: "00000004", nextBlockId: "00000001" },
    { type: "deleteBlocks", story: OP_STORIES.MAIN, blockIds: ["00000003"] },
    { type: "resolveRevision", story: OP_STORIES.MAIN, revisionIds: [42], decision: "accept" },
  ],
  effects: [
    { type: "splitBlock", newHalf: "first" },
    { type: "joinBlocks", firstLength: 3 },
    { type: "touchedBlocks", blockIds: ["00000001", "00000003"] },
    { type: "touchedBlocks", blockIds: ["00000001"] },
  ],
} as const satisfies SequencedBatch;

if (import.meta.main) {
  await Bun.write(
    path.join(import.meta.dir, "../__fixtures__", `batches-v${DOCUMENT_OP_SCHEMA_VERSION}.json`),
    `${JSON.stringify([...envelopeFixtures, sequencedFixture], null, 2)}\n`,
  );
}
