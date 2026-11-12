import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document, Paragraph } from "../../../model/document";
import { serializeDocumentToDocx } from "../../../serialize/docx";
import { storyParagraphs } from "../../blocks";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
  type EditorIntentMode,
} from "../../editorIntent";
import { equalForStaleness } from "../../equality";
import {
  type IdLease,
  introducedIds,
  leasedIdCensus,
  trimIdLease,
} from "../../idLease";
import { paragraphLength } from "../../offsets";
import { DOCUMENT_OP_SCHEMA_VERSION, OP_STORIES } from "../../types";
import { createClient } from "../client";
import {
  BATCH_REJECTION_REASONS,
  type BatchRejectionReason,
  type DocumentBatch,
  type SequencedBatch,
} from "../envelope";
import { type BatchSubmission, createSequencer } from "../sequencer";

const NUM_RUNS = 40;
const DATE = "2026-01-01T00:00:00Z";

/** Each peer owns one paragraph of the base: concurrent edits never touch one block. */
const baseParagraph = (peer: number): Paragraph => ({
  type: "paragraph",
  paraId: (0x1000_0000 + peer).toString(16).toUpperCase(),
  content: [
    { type: "run", content: [{ type: "text", text: `Peer ${peer} ` }] },
    {
      type: "insertion",
      info: { id: 10 + peer, author: "base", date: DATE },
      content: [{ type: "run", content: [{ type: "text", text: "tracked" }] }],
    },
    {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 100 + peer, tag: "clause" },
      content: [{ type: "run", content: [{ type: "text", text: "clause" }] }],
    },
    { type: "run", content: [{ type: "text", text: " tail" }] },
  ],
});

const baseDocument = (peers: number): Document => ({
  package: {
    document: {
      content: Array.from({ length: peers }, (_, peer) => baseParagraph(peer)),
    },
  },
});

/** Edits that mint ids: new paragraphs, tracked text and content-control cuts. */
const EDITS = [
  "split",
  "suggestSplit",
  "suggestInsert",
  "suggestDelete",
  "insertParagraph",
  "suggestParagraph",
] as const;
type Edit = (typeof EDITS)[number];

type Step =
  | { type: "edit"; peer: number; edit: Edit; seed: number }
  | { type: "submit"; peer: number }
  | { type: "deliver"; seed: number };

/** How peers mint: from their own lease, or the lowest free ids, ignoring it. */
type Minting = "leased" | "sequential";

type Peer = {
  actor: string;
  client: ReturnType<typeof createClient>;
  lease: IdLease | undefined;
  /** Paragraphs this peer created or was given; no other peer edits them. */
  owned: string[];
  allocate: ReturnType<typeof createEditorIntentIdAllocator>;
};

type Delivery =
  | { peer: number; type: "result"; result: BatchSubmission }
  | { peer: number; type: "broadcast"; batch: SequencedBatch };

const paragraphOf = (document: Document, blockId: string) =>
  storyParagraphs(document.package.document).find(
    ({ paragraph }) => paragraph.paraId === blockId,
  )?.paragraph;

const compileEdit = (
  peer: Peer,
  edit: Edit,
  seed: number,
  serial: number,
): DocumentBatch | undefined => {
  const opId = `edit-${serial}`;
  const document = peer.client.document;
  const blockId = peer.owned.at(seed % peer.owned.length);
  if (blockId === undefined) return undefined;
  const paragraph = paragraphOf(document, blockId);
  if (paragraph === undefined) return undefined;
  const length = paragraphLength(paragraph);
  const offset = Math.floor(seed / 7) % (length + 1);
  const at = { story: OP_STORIES.MAIN, blockId, offset };
  const suggesting = edit.startsWith("suggest");
  let request: Parameters<Peer["allocate"]>[1];
  switch (edit) {
    case "split":
    case "suggestSplit":
      request = { type: "splitParagraph", at };
      break;
    case "suggestInsert":
      request = { type: "replaceText", from: at, to: at, text: "x" };
      break;
    case "suggestDelete":
      if (offset >= length) return undefined;
      request = {
        type: "replaceText",
        from: at,
        to: { ...at, offset: offset + 1 },
        text: "",
      };
      break;
    case "insertParagraph":
    case "suggestParagraph":
      request = {
        type: "insertParagraphs",
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId },
        paragraphs: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text: opId }] }],
          },
        ],
      };
      break;
    default: {
      const exhaustive: never = edit;
      return exhaustive;
    }
  }
  const allocation = peer.allocate(document, request, peer.lease);
  const intent: EditorIntent =
    request.type === "splitParagraph"
      ? { ...request, newBlockId: allocation.newBlockId }
      : request;
  const leased = peer.lease === undefined ? {} : { idLease: peer.lease };
  const mode: EditorIntentMode = suggesting
    ? {
        type: "suggesting",
        revision: {
          id: allocation.revisionId,
          author: peer.actor,
          // Distinct per batch, so no new record merges into an earlier one.
          date: new Date(Date.UTC(2026, 0, 1) + serial * 1000).toISOString(),
        },
        newIds: allocation.newIds,
        ...leased,
      }
    : { type: "editing", newIds: allocation.newIds, ...leased };
  const compiled = compileEditorIntent(document, { intent, mode });
  if (compiled.isErr()) return undefined;
  return {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId,
    actor: peer.actor,
    baseRev: peer.client.headRev,
    ops: compiled.value.ops,
  };
};

type Outcome = {
  enqueued: string[];
  sequenced: string[];
  /** Why the authority refused each refused batch. */
  rejected: BatchRejectionReason[];
  dropped: string[];
  /** Sequenced batches that introduced ids. */
  mintsSequenced: number;
  /** Batches minting ids while another peer's minting batch was unsequenced. */
  concurrentMints: number;
  documents: Document[];
  authority: Document;
};

const simulate = (
  peers: number,
  steps: readonly Step[],
  minting: Minting,
): Outcome => {
  const base = baseDocument(peers);
  const sequencer = createSequencer(base);
  const team: Peer[] = Array.from({ length: peers }, (_, index) => {
    const actor = `actor-${index}`;
    const granted = sequencer.grantIdLease(actor);
    if (granted.isErr()) throw granted.error;
    return {
      actor,
      client: createClient(base),
      lease: minting === "leased" ? granted.value : undefined,
      owned: [(0x1000_0000 + index).toString(16).toUpperCase()],
      allocate: createEditorIntentIdAllocator(),
    };
  });
  const enqueued: string[] = [];
  const rejected: BatchRejectionReason[] = [];
  let mintsSequenced = 0;
  const delivery: Delivery[] = [];
  const unsequencedMints = new Map<string, number>();
  let concurrentMints = 0;
  let next = 10;

  const submit = (index: number): void => {
    const peer = team.at(index);
    const batch = peer?.client.nextSubmission();
    if (batch === undefined) return;
    const result = sequencer.submit(batch);
    delivery.push({ peer: index, type: "result", result });
    if (result.type === "reject") {
      rejected.push(result.reason.reason);
      unsequencedMints.delete(batch.opId);
      return;
    }
    if (unsequencedMints.delete(batch.opId)) mintsSequenced += 1;
    const broadcast = sequencer.broadcasts.at(-1);
    if (broadcast === undefined) return;
    for (let target = 0; target < peers; target += 1)
      delivery.push({ peer: target, type: "broadcast", batch: broadcast });
  };

  const deliver = (seed: number): void => {
    const event = delivery.splice(seed % delivery.length, 1).at(0);
    if (event === undefined) return;
    const peer = team.at(event.peer);
    if (peer === undefined) return;
    if (event.type === "broadcast")
      expect(peer.client.receiveBroadcast(event.batch).isOk()).toBe(true);
    else if (event.result.type === "ack") peer.client.receiveAck(event.result);
    else peer.client.receiveReject(event.result);
  };

  for (const step of steps) {
    if (step.type === "deliver") {
      if (delivery.length > 0) deliver(step.seed);
      continue;
    }
    const peer = team.at(step.peer % peers);
    if (peer === undefined) continue;
    if (step.type === "submit") {
      submit(step.peer % peers);
      continue;
    }
    const before = peer.client.document;
    const batch = compileEdit(peer, step.edit, step.seed, next++);
    if (batch === undefined) continue;
    const opId = batch.opId;
    const queued = peer.client.enqueue(batch);
    if (queued.isErr()) continue;
    enqueued.push(opId);
    const minted = introducedIds(
      leasedIdCensus(before),
      leasedIdCensus(peer.client.document),
    );
    if (minted.length > 0) {
      if (
        [...unsequencedMints.values()].some(
          (owner) => owner !== step.peer % peers,
        )
      )
        concurrentMints += 1;
      unsequencedMints.set(opId, step.peer % peers);
    }
    // A peer mints upward through its lease, never an id twice.
    if (peer.lease !== undefined) peer.lease = trimIdLease(peer.lease, minted);
    for (const { kind, id } of minted)
      if (kind === "paragraph")
        peer.owned.push(id.toString(16).padStart(8, "0").toUpperCase());
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
    enqueued,
    sequenced: sequencer.broadcasts.map(({ opId }) => opId),
    rejected,
    mintsSequenced,
    dropped: team.flatMap(({ client }) =>
      client.notices.map(({ opId }) => opId),
    ),
    concurrentMints,
    documents: team.map(({ client }) => client.document),
    authority: sequencer.document,
  };
};

const stepArbitrary = (peers: number): fc.Arbitrary<Step> =>
  fc.oneof(
    {
      weight: 4,
      arbitrary: fc.record({
        type: fc.constant("edit" as const),
        peer: fc.nat({ max: peers - 1 }),
        edit: fc.constantFrom(...EDITS),
        seed: fc.nat({ max: 100_000 }),
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

const scenarioArbitrary = fc
  .integer({ min: 2, max: 4 })
  .chain((peers) =>
    fc.tuple(
      fc.constant(peers),
      fc.array(stepArbitrary(peers), { minLength: 8, maxLength: 40 }),
    ),
  );

const bytes = async (document: Document): Promise<Uint8Array> =>
  new Uint8Array(await serializeDocumentToDocx(document));

/**
 * Every peer ends on the authority's bytes, every edit a peer made is either
 * sequenced or reported back to it, and none is refused over an id. Refusals
 * of operation pairs the transform does not support yet stay allowed.
 */
const ID_REFUSALS: ReadonlySet<BatchRejectionReason> = new Set([
  BATCH_REJECTION_REASONS.OUTSIDE_ID_LEASE,
  // A repeated id fails the apply contract.
  BATCH_REJECTION_REASONS.INVALID_OPERATION,
]);
const convergence = (
  minting: Minting,
  counts = { concurrent: 0, mintsSequenced: 0 },
) =>
  fc.asyncProperty(scenarioArbitrary, async ([peers, steps]) => {
    const outcome = simulate(peers, steps, minting);
    counts.concurrent += outcome.concurrentMints;
    counts.mintsSequenced += outcome.mintsSequenced;
    expect(
      outcome.rejected.filter((reason) => ID_REFUSALS.has(reason)),
    ).toEqual([]);
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

describe("id leases", () => {
  test(
    "concurrent peers minting ids converge on the authority's bytes without losing an edit",
    async () => {
      await assertProperty(convergence("leased"), {
        numRuns: NUM_RUNS,
        id: "id lease convergence",
      });
      // On a fixed sample, the schedules do race: peers mint while another's
      // minting edit is unsequenced, and minting edits reach the journal.
      const counts = { concurrent: 0, mintsSequenced: 0 };
      await checkCoverageSample(convergence("leased", counts), NUM_RUNS);
      expect(counts.concurrent).toBeGreaterThan(0);
      expect(counts.mintsSequenced).toBeGreaterThan(NUM_RUNS);
    },
    propertyTestTimeout(60_000),
  );

  test(
    "peers ignoring their leases lose edits, and the property says so",
    async () => {
      const result = await fc.check(
        convergence("sequential"),
        // The break is a counterexample on purpose: report nothing.
        propertyConfig({ numRuns: NUM_RUNS, reporter: () => undefined }),
      );
      expect(result.failed).toBe(true);
    },
    propertyTestTimeout(60_000),
  );
});
