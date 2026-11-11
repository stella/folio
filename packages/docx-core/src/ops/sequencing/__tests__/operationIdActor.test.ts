import { expect, test } from "bun:test";

import type { Document } from "../../../model/document";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "../../types";
import { createClient } from "../client";
import type { DocumentBatch } from "../envelope";
import { createSequencer } from "../sequencer";

const PARAGRAPH = "70000001";

const document = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: PARAGRAPH,
          content: [{ type: "run", content: [{ type: "text", text: "ab" }] }],
        },
      ],
    },
  },
});

const insertion = (actor: string, text: string): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: "shared-id",
  actor,
  baseRev: 0,
  ops: [
    {
      type: "insertText",
      at: { story: OP_STORIES.MAIN, blockId: PARAGRAPH, offset: 1 },
      text,
      runProps: INHERIT_RUN_PROPS,
    },
  ],
});

test("a replayed outcome answers only the actor that first used the operation id", () => {
  const sequencer = createSequencer(document());
  const first = sequencer.submit(insertion("first", "X"));
  expect(first.type).toBe("ack");

  expect(sequencer.submit(insertion("first", "X"))).toBe(first);
  const second = sequencer.submit(insertion("second", "Y"));
  expect(second.type).toBe("reject");
  expect(sequencer.headRev).toBe(1);
  // The other actor's attempt does not replace the stored outcome.
  expect(sequencer.submit(insertion("first", "X"))).toBe(first);
});

test("a broadcast under another actor does not settle a pending batch with the same id", () => {
  const sequencer = createSequencer(document());
  expect(sequencer.submit(insertion("second", "Y")).type).toBe("ack");
  const broadcast = sequencer.broadcasts.at(0);
  if (broadcast === undefined) throw new Error("Expected a broadcast");

  const client = createClient(document());
  expect(client.enqueue(insertion("first", "X")).isOk()).toBe(true);
  expect(client.receiveBroadcast(broadcast).isErr()).toBe(true);
  expect(client.pending.map(({ actor }) => actor)).toEqual(["first"]);
});
