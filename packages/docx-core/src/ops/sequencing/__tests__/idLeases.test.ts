import { describe, expect, test } from "bun:test";

import type { Document } from "../../../model/document";
import { equalForStaleness } from "../../equality";
import {
  formatParaId,
  type IdLease,
  idLeaseHolds,
  leasedIdCensus,
} from "../../idLease";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  type DocumentOp,
  OP_STORIES,
} from "../../types";
import type { DocumentBatch } from "../envelope";
import { ID_LEASE_REFUSAL_REASONS } from "../idLeases";
import { createSequencer } from "../sequencer";

const FIRST = "00000001";
const SECOND = "00000002";

const document = (): Document => ({
  package: {
    document: {
      content: [FIRST, SECOND].map((paraId) => ({
        type: "paragraph",
        paraId,
        content: [{ type: "run", content: [{ type: "text", text: "text" }] }],
      })),
    },
  },
});

const batch = (
  actor: string,
  opId: string,
  baseRev: number,
  ops: readonly DocumentOp[],
): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor,
  baseRev,
  ops,
});

const split = (newBlockId: string, blockId = FIRST): DocumentOp => ({
  type: "splitBlock",
  at: { story: OP_STORIES.MAIN, blockId, offset: 2 },
  newBlockId,
  newHalf: "second",
});

const firstOf = (lease: IdLease, kind: string): number => {
  const block = lease.blocks.find((candidate) => candidate.kind === kind);
  if (block === undefined) throw new Error(`No ${kind} block.`);
  return block.start;
};

describe("sequencer id leases", () => {
  test("refuses a batch introducing an id outside the submitter's lease", () => {
    const sequencer = createSequencer(document());
    const first = sequencer.grantIdLease("first").unwrap();
    sequencer.grantIdLease("second").unwrap();
    const before = sequencer.document;
    // The second actor mints an id only the first leases.
    const theirs = formatParaId(firstOf(first, "paragraph"));
    const refused = sequencer.submit(
      batch("second", "foreign", 0, [split(theirs)]),
    );
    expect(refused.type).toBe("reject");
    if (refused.type !== "reject") return;
    expect(refused.reason.reason).toBe("outsideIdLease");
    expect(refused.reason.idKind).toBe("paragraph");
    expect(sequencer.headRev).toBe(0);
    expect(equalForStaleness(sequencer.document, before)).toBe(true);
    // An actor with no lease may introduce no id at all.
    const unleased = sequencer.submit(
      batch("third", "unleased", 0, [split("7FFFFFF0")]),
    );
    expect(unleased.type).toBe("reject");
    if (unleased.type === "reject")
      expect(unleased.reason.reason).toBe("outsideIdLease");
    // The lessee itself may.
    expect(
      sequencer.submit(batch("first", "own", 0, [split(theirs)])).type,
    ).toBe("ack");
  });

  test("grants disjoint blocks that avoid the room's ids, and gives a lease back unchanged", () => {
    const sequencer = createSequencer(document());
    const first = sequencer.grantIdLease("first").unwrap();
    const second = sequencer.grantIdLease("second").unwrap();
    const blocks = [...first.blocks, ...second.blocks];
    const annotation = (kind: string) => kind !== "paragraph";
    for (const [index, block] of blocks.entries())
      for (const other of blocks.slice(index + 1))
        if (
          block.kind === other.kind ||
          (annotation(block.kind) && annotation(other.kind))
        )
          expect(block.start < other.end && other.start < block.end).toBe(
            false,
          );
    for (const id of [1, 2])
      expect(idLeaseHolds(first, "paragraph", id)).toBe(false);
    // A reloading actor gets its lease back, and nothing new is granted.
    expect(sequencer.grantIdLease("first").unwrap()).toEqual(first);
    expect(sequencer.idLeases.grants).toHaveLength(blocks.length);
  });

  test("an actor's remaining lease starts above what it minted, across a restart", () => {
    const sequencer = createSequencer(document());
    const lease = sequencer.grantIdLease("first").unwrap();
    const minted = firstOf(lease, "paragraph") + 5;
    expect(
      sequencer.submit(batch("first", "own", 0, [split(formatParaId(minted))]))
        .type,
    ).toBe("ack");
    const remaining = sequencer.grantIdLease("first").unwrap();
    expect(firstOf(remaining, "paragraph")).toBe(minted + 1);
    const restarted = createSequencer(sequencer.document, {
      idLeases: sequencer.idLeases,
    });
    expect(restarted.grantIdLease("first").unwrap()).toEqual(remaining);
    // Below what the lessee minted, nobody mints again: only a restoration
    // brings such an id back, whoever submits it.
    restarted.grantIdLease("second").unwrap();
    expect(
      restarted.submit(
        batch("second", "restore", 0, [split(formatParaId(minted - 1))]),
      ).type,
    ).toBe("ack");
  });

  test("accepts any actor restoring an id the room held", () => {
    const sequencer = createSequencer(document());
    sequencer.grantIdLease("first").unwrap();
    sequencer.grantIdLease("second").unwrap();
    expect(
      sequencer.submit(
        batch("first", "delete", 0, [
          { type: "deleteBlocks", story: OP_STORIES.MAIN, blockIds: [FIRST] },
        ]),
      ),
    ).toMatchObject({ type: "ack" });
    expect(leasedIdCensus(sequencer.document).paragraph.has(1)).toBe(false);
    expect(
      sequencer.submit(batch("second", "restore", 1, [split(FIRST, SECOND)])),
    ).toMatchObject({ type: "ack" });
  });

  test("renews only an actor running low, and refuses rather than wraps when the space is spent", () => {
    const sequencer = createSequencer(document(), {
      idLeases: {
        blockSizes: {
          paragraph: 0x3000_0000,
          revision: 16,
          control: 16,
          comment: 16,
        },
        grants: [],
        minted: [],
      },
    });
    const lease = sequencer.grantIdLease("first").unwrap();
    // One block of spare ids may be held ahead of what was minted.
    const spare = sequencer.renewIdLease("first", "paragraph");
    expect(spare.isOk()).toBe(true);
    const notDue = sequencer.renewIdLease("first", "paragraph");
    expect(notDue.isErr() && notDue.error.reason).toBe(
      ID_LEASE_REFUSAL_REASONS.NOT_DUE,
    );
    // Two blocks of 0x30000000 leave no third below 0x80000000.
    const before = sequencer.idLeases;
    const refused = sequencer.grantIdLease("second");
    expect(refused.isErr() && refused.error.reason).toBe(
      ID_LEASE_REFUSAL_REASONS.EXHAUSTED,
    );
    expect(sequencer.idLeases).toEqual(before);
    for (const { block } of sequencer.idLeases.grants)
      expect(block.end).toBeLessThanOrEqual(0x8000_0000);
    expect(firstOf(lease, "revision")).toBeGreaterThan(0);
  });

  test("refuses a restored lease table whose blocks overlap", () => {
    const block = { kind: "paragraph", start: 10, end: 20 } as const;
    expect(() =>
      createSequencer(document(), {
        idLeases: {
          blockSizes: { paragraph: 10, revision: 10, control: 10, comment: 10 },
          grants: [
            { actor: "first", block },
            { actor: "second", block: { ...block, start: 15, end: 25 } },
          ],
          minted: [],
        },
      }),
    ).toThrow("overlap");
  });
});
