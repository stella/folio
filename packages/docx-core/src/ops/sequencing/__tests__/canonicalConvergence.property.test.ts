/**
 * Peers editing one room concurrently with every kind of canonical edit,
 * through the real sequencer and optimistic client, with every batch crossing
 * the JSON wire: all of them end on the authority's bytes, no edit is lost
 * without a notice, and every concurrent batch the authority accepted
 * commutes with the journal tail it was rebased over.
 */
import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document } from "../../../model/document";
import { serializeDocumentToDocx } from "../../../serialize/docx";
import { createEditorIntentIdAllocator } from "../../editorIntent";
import { equalForStaleness } from "../../equality";
import { introducedIds, leasedIdCensus, trimIdLease } from "../../idLease";
import { DOCUMENT_OP_SCHEMA_VERSION, type DocumentOp } from "../../types";
import { applyBatch } from "../applyBatch";
import { createClient } from "../client";
import * as envelope from "../envelope";
import {
  BATCH_REJECTION_REASONS,
  type BatchRejectionReason,
  type DocumentBatch,
  type SequencedBatch,
  validateDocumentBatch,
} from "../envelope";
import { isBlockScoped } from "../scope";
import { type BatchSubmission, createSequencer } from "../sequencer";
import { transformBatch } from "../transform";
import {
  canonicalBase,
  compileCanonicalEdit,
  type EditAuthor,
  type EditSeed,
  editSeedArbitrary,
} from "./canonicalEdits";

const NUM_RUNS = 25;

type Step =
  | { type: "edit"; peer: number; seed: EditSeed }
  | { type: "undo"; peer: number }
  | { type: "submit"; peer: number }
  | { type: "deliver"; seed: number };

type Peer = {
  author: EditAuthor;
  client: ReturnType<typeof createClient>;
  /** The inverse of this peer's last edit, as applying it returned. */
  undo: readonly DocumentOp[] | undefined;
};

type Delivery =
  | { peer: number; type: "result"; result: BatchSubmission }
  | { peer: number; type: "broadcast"; batch: SequencedBatch };

const wire = <Value>(value: Value): Value => JSON.parse(JSON.stringify(value));

/** Refusals that would mean the wire, the base or the leases failed. */
const FORBIDDEN: ReadonlySet<BatchRejectionReason> = new Set([
  BATCH_REJECTION_REASONS.INVALID_BATCH,
  BATCH_REJECTION_REASONS.UNSUPPORTED_SCHEMA,
  BATCH_REJECTION_REASONS.STALE_BASE,
  BATCH_REJECTION_REASONS.OUTSIDE_ID_LEASE,
]);

type Outcome = {
  enqueued: string[];
  sequenced: string[];
  rejected: BatchRejectionReason[];
  dropped: string[];
  documents: Document[];
  authority: Document;
  /** Accepted batches rebased over a tail, and how many were proven to commute. */
  rebased: number;
  proven: number;
  scopedRebased: number;
  /** Accepted rebased batches whose result differs from the commuted order. */
  diverged: string[];
};

/**
 * Apply `batch` where it was authored, then each tail batch rebased over it:
 * the result must be the authority's. Undefined when a step of that order is
 * refused, which leaves nothing to compare.
 */
const commutes = (
  history: readonly Document[],
  batch: DocumentBatch,
  tail: readonly SequencedBatch[],
  result: Document,
): boolean | undefined => {
  const base = history.at(batch.baseRev);
  if (base === undefined) return undefined;
  const first = applyBatch(base, batch.ops);
  if (first.isErr()) return undefined;
  let document = first.value.document;
  let mine: SequencedBatch = {
    ...batch,
    revision: batch.baseRev + 1,
    effects: first.value.effects,
  };
  for (const remote of tail) {
    const moved = transformBatch(remote, [mine], { order: "before" });
    if (moved.isErr()) return undefined;
    const applied = applyBatch(document, moved.value.ops);
    if (applied.isErr()) return undefined;
    document = applied.value.document;
    const next = transformBatch(mine, [remote]);
    const at = history.at(remote.revision);
    if (next.isErr() || at === undefined) return undefined;
    const placed = applyBatch(at, next.value.ops);
    if (placed.isErr()) return undefined;
    mine = {
      ...next.value,
      revision: remote.revision + 1,
      effects: placed.value.effects,
    };
  }
  return equalForStaleness(document, result);
};

const simulate = (peers: number, steps: readonly Step[]): Outcome => {
  const base = canonicalBase();
  const sequencer = createSequencer(base);
  const history: Document[] = [base];
  const team: Peer[] = Array.from({ length: peers }, (_, index) => {
    const actor = `actor-${index}`;
    const granted = sequencer.grantIdLease(actor);
    if (granted.isErr()) throw granted.error;
    return {
      author: {
        actor,
        tag: String(index),
        lease: granted.value,
        allocate: createEditorIntentIdAllocator(),
      },
      client: createClient(base),
      undo: undefined,
    };
  });
  const outcome: Outcome = {
    enqueued: [],
    sequenced: [],
    rejected: [],
    dropped: [],
    documents: [],
    authority: base,
    rebased: 0,
    proven: 0,
    scopedRebased: 0,
    diverged: [],
  };
  const delivery: Delivery[] = [];
  let serial = 10;

  const enqueue = (peer: Peer, ops: readonly DocumentOp[]): void => {
    const before = peer.client.document;
    const batch: DocumentBatch = {
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: `edit-${serial++}`,
      actor: peer.author.actor,
      baseRev: peer.client.headRev,
      ops,
    };
    // Every compiled edit and every undo crosses the wire.
    const sent = validateDocumentBatch(wire(batch));
    if (sent.isErr())
      throw new Error(
        `${ops.map(({ type }) => type).join(", ")} is not sendable: ${sent.error.message}`,
      );
    const queued = peer.client.enqueue(wire(batch));
    // A stale undo is dropped here, never forced.
    if (queued.isErr()) {
      expect(queued.error.reason).toBe(
        BATCH_REJECTION_REASONS.INVALID_OPERATION,
      );
      return;
    }
    outcome.enqueued.push(batch.opId);
    const applied = applyBatch(before, queued.value.ops);
    peer.undo = applied.isOk() ? applied.value.inverse : undefined;
    const minted = introducedIds(
      leasedIdCensus(before),
      leasedIdCensus(peer.client.document),
    );
    const lease = peer.author.lease;
    if (lease !== undefined)
      peer.author = { ...peer.author, lease: trimIdLease(lease, minted) };
  };

  const submit = (index: number): void => {
    const peer = team.at(index);
    const batch = peer?.client.nextSubmission();
    if (batch === undefined) return;
    const tail = sequencer.broadcasts.slice(batch.baseRev);
    const result = sequencer.submit(wire(batch));
    delivery.push({ peer: index, type: "result", result });
    if (result.type === "reject") {
      outcome.rejected.push(result.reason.reason);
      return;
    }
    history.push(sequencer.document);
    if (tail.length > 0) {
      outcome.rebased += 1;
      if (
        [batch, ...tail].some(({ ops }) => ops.some((op) => isBlockScoped(op)))
      )
        outcome.scopedRebased += 1;
      const proven = commutes(history, batch, tail, sequencer.document);
      if (proven === true) outcome.proven += 1;
      if (proven === false) outcome.diverged.push(batch.opId);
    }
    const broadcast = sequencer.broadcasts.at(-1);
    if (broadcast === undefined) return;
    for (let target = 0; target < peers; target += 1)
      delivery.push({
        peer: target,
        type: "broadcast",
        batch: wire(broadcast),
      });
  };

  const deliver = (seed: number): void => {
    const event = delivery.splice(seed % delivery.length, 1).at(0);
    const peer = event === undefined ? undefined : team.at(event.peer);
    if (event === undefined || peer === undefined) return;
    if (event.type === "broadcast") {
      const received = peer.client.receiveBroadcast(event.batch);
      if (received.isErr()) throw received.error;
    } else if (event.result.type === "ack")
      peer.client.receiveAck(event.result);
    else peer.client.receiveReject(event.result);
  };

  for (const step of steps) {
    if (step.type === "deliver") {
      if (delivery.length > 0) deliver(step.seed);
      continue;
    }
    const index = step.peer % peers;
    const peer = team.at(index);
    if (peer === undefined) continue;
    if (step.type === "submit") {
      submit(index);
      continue;
    }
    if (step.type === "undo") {
      if (peer.undo !== undefined && peer.undo.length > 0)
        enqueue(peer, peer.undo);
      peer.undo = undefined;
      continue;
    }
    const ops = compileCanonicalEdit(
      peer.client.document,
      peer.author,
      step.seed,
      serial,
    );
    if (ops !== undefined && ops.length > 0) enqueue(peer, ops);
  }
  let rounds = 0;
  while (
    delivery.length > 0 ||
    team.some(({ client }) => client.pending.length > 0)
  ) {
    expect(rounds++).toBeLessThan(10_000);
    for (let index = 0; index < peers; index += 1) submit(index);
    if (delivery.length > 0) deliver(rounds * 7919);
  }
  return {
    ...outcome,
    sequenced: sequencer.broadcasts.map(({ opId }) => opId),
    dropped: team.flatMap(({ client }) =>
      client.notices.map(({ opId }) => opId),
    ),
    documents: team.map(({ client }) => client.document),
    authority: sequencer.document,
  };
};

const stepArbitrary = (peers: number): fc.Arbitrary<Step> =>
  fc.oneof(
    {
      weight: 5,
      arbitrary: fc.record({
        type: fc.constant("edit" as const),
        peer: fc.nat({ max: peers - 1 }),
        seed: editSeedArbitrary,
      }),
    },
    {
      weight: 1,
      arbitrary: fc.record({
        type: fc.constant("undo" as const),
        peer: fc.nat({ max: peers - 1 }),
      }),
    },
    {
      weight: 2,
      arbitrary: fc.record({
        type: fc.constant("submit" as const),
        peer: fc.nat({ max: peers - 1 }),
      }),
    },
    {
      weight: 3,
      arbitrary: fc.record({
        type: fc.constant("deliver" as const),
        seed: fc.nat({ max: 100_000 }),
      }),
    },
  );

type Scenario = [number, Step[]];

const scenarioArbitrary: fc.Arbitrary<Scenario> = fc
  .integer({ min: 3, max: 4 })
  .chain((peers) =>
    fc.tuple(
      fc.constant(peers),
      fc.array(stepArbitrary(peers), { minLength: 8, maxLength: 30 }),
    ),
  );

const bytes = async (document: Document): Promise<Uint8Array> =>
  new Uint8Array(await serializeDocumentToDocx(document));

type Counts = { rebased: number; proven: number; scopedRebased: number };

const convergence = (
  counts: Counts = { rebased: 0, proven: 0, scopedRebased: 0 },
) =>
  fc.asyncProperty(scenarioArbitrary, async ([peers, steps]) => {
    const outcome = simulate(peers, steps);
    counts.rebased += outcome.rebased;
    counts.proven += outcome.proven;
    counts.scopedRebased += outcome.scopedRebased;
    // Every refusal is typed, and none is a wire, base or lease failure.
    expect(outcome.rejected.filter((reason) => FORBIDDEN.has(reason))).toEqual(
      [],
    );
    // Concurrent batches the authority accepted commute with what they were
    // rebased over: the transform never accepts a pair it maps wrongly.
    if (outcome.diverged.length > 0)
      throw new Error(
        `Accepted batches do not commute with their tail: ${outcome.diverged.join(", ")}`,
      );
    // No acknowledged edit is lost, and no edit vanishes unreported.
    expect(new Set(outcome.sequenced).size).toBe(outcome.sequenced.length);
    expect([...outcome.sequenced, ...outcome.dropped].toSorted()).toEqual(
      outcome.enqueued.toSorted(),
    );
    const authority = await bytes(outcome.authority);
    for (const document of outcome.documents) {
      expect(equalForStaleness(document, outcome.authority)).toBe(true);
      expect(await bytes(document)).toEqual(authority);
    }
  });

const edit = (peer: number, seed: Omit<EditSeed, "suggest">): Step => ({
  type: "edit",
  peer,
  seed: { ...seed, suggest: false },
});

/** Three peers restructure one table at once: rows at different indexes. */
const concurrentRows: Scenario = [
  3,
  [
    edit(0, { kind: "insertRow", a: 1, b: 0 }),
    edit(1, { kind: "insertRow", a: 2, b: 3 }),
    edit(2, { kind: "type", a: 0, b: 0 }),
    { type: "submit", peer: 0 },
    { type: "submit", peer: 1 },
    { type: "submit", peer: 2 },
  ],
];

describe("canonical sequencing", () => {
  test(
    "concurrent canonical edits converge on the authority's bytes and commute where accepted",
    async () => {
      await assertProperty(convergence(), {
        numRuns: NUM_RUNS,
        id: "canonical sequencing convergence",
        examples: [concurrentRows],
      });
      // On a fixed sample the schedules race: accepted batches are rebased
      // over concurrent ones, block-scoped ones among them, and the commuted
      // order is proven for them.
      const counts = { rebased: 0, proven: 0, scopedRebased: 0 };
      await checkCoverageSample(convergence(counts), NUM_RUNS);
      expect(counts.rebased).toBeGreaterThan(0);
      expect(counts.scopedRebased).toBeGreaterThan(0);
      expect(counts.proven).toBeGreaterThan(0);
    },
    propertyTestTimeout(180_000),
  );

  test(
    "a transform that skips the table structure check is caught",
    async () => {
      // The test-only break: no operation counts as a table structure edit,
      // so the footprint check lets two edits of one table's rows through.
      const original = { ...envelope };
      mock.module("../envelope", () => ({
        ...original,
        TABLE_EXCLUSIVE_OP_TYPES: {},
      }));
      try {
        const result = await fc.check(
          convergence(),
          // The break is a counterexample on purpose: report nothing.
          propertyConfig({
            numRuns: NUM_RUNS,
            examples: [concurrentRows],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
        // Caught by the commutation check, not by anything else.
        expect(String(result.errorInstance)).toMatch(/do not commute/u);
      } finally {
        mock.module("../envelope", () => original);
      }
      // Without the break the same scenario passes.
      const restored = await fc.check(
        convergence(),
        propertyConfig({ numRuns: 1, examples: [concurrentRows] }),
      );
      expect(restored.failed).toBe(false);
    },
    propertyTestTimeout(180_000),
  );
});
