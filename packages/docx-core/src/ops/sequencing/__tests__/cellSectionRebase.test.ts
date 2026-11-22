/**
 * A pending edit that took a section out of a table cell is taken back by its
 * recorded inverse when the authority rejects it or a foreign edit drops it,
 * though as an operation of its own that inverse is refused.
 */

import { panic } from "better-result";
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../../model/document";
import { hasSectionInTableCell } from "../../cellSections";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "../../types";
import { createClient } from "../client";
import type { DocumentBatch } from "../envelope";
import { createSequencer } from "../sequencer";

const story = OP_STORIES.MAIN;

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const inner = paragraph("70000001", "inner");
const ending: Paragraph = {
  ...inner,
  sectionProperties: { pageWidth: 11_906 },
};

const arrived = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          columnWidths: [2400],
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [
                    structuredClone(ending),
                    paragraph("70000002", "b"),
                  ],
                },
              ],
            },
          ],
        },
        paragraph("70000003", "after"),
      ],
    },
  },
});

const batch = (opId: string, actor: string, op: DocumentOp): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor,
  baseRev: 0,
  ops: [op],
});

const removal = batch("removal", "local", {
  type: "replaceBlocks",
  story,
  expected: [ending],
  blocks: [inner],
  sectionBoundaries: "replace",
});

const foreign = batch("typing", "foreign", {
  type: "insertText",
  at: { story, blockId: "70000001", offset: 0 },
  text: "x",
  runProps: INHERIT_RUN_PROPS,
});

test("a rejected removal is taken back", () => {
  const client = createClient(arrived());
  client.enqueue(removal).unwrap();
  const submitted = client.nextSubmission();
  if (submitted === undefined) panic("The removal is submitted.");
  expect(hasSectionInTableCell(client.document)).toBe(false);
  const sequencer = createSequencer(arrived());
  const answer = sequencer.submit({ ...submitted, baseRev: 5 });
  if (answer.type !== "reject") panic("A batch from no revision is rejected.");
  client.receiveReject(answer);
  expect(client.document).toEqual(arrived());
  expect(client.notices.map(({ opId }) => opId)).toEqual(["removal"]);
});

test("a removal a foreign edit drops is taken back", () => {
  const client = createClient(arrived());
  client.enqueue(removal).unwrap();
  const sequencer = createSequencer(arrived());
  expect(sequencer.submit(foreign).type).toBe("ack");
  const broadcast = sequencer.broadcasts.at(0);
  if (broadcast === undefined) panic("The foreign edit is broadcast.");
  client.receiveBroadcast(broadcast).unwrap();
  expect(client.notices.map(({ opId }) => opId)).toEqual(["removal"]);
  expect(client.document).toEqual(sequencer.document);
  expect(hasSectionInTableCell(client.document)).toBe(true);
});

test("a sequenced undo of the removal is refused", () => {
  const sequencer = createSequencer(arrived());
  expect(sequencer.submit(removal).type).toBe("ack");
  const undo = sequencer.submit({
    ...removal,
    opId: "undo",
    baseRev: 1,
    ops: [
      {
        type: "replaceBlocks",
        story,
        expected: [inner],
        blocks: [ending],
        sectionBoundaries: "replace",
      },
    ],
  });
  expect(undo.type).toBe("reject");
  expect(hasSectionInTableCell(sequencer.document)).toBe(false);
});
