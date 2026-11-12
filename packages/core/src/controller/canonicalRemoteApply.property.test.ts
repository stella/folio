import { panic, type Result } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  BATCH_REJECTION_REASONS,
  type ClientDocumentChange,
  createClient,
  createSequencer,
  DOCUMENT_OP_SCHEMA_VERSION,
  type SequencedBatch,
} from "@stll/docx-core/ops";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { parseDocx } from "../docx/parser";
import {
  type BatchMeta,
  caretOf,
  documentText,
  type EditingPeer,
  expectIdenticalSaves,
  expectSequencedKeystrokes,
  homeState,
  LOCAL_KINDS,
  localEdit,
  type LocalKind,
  MAX_PEERS,
  paragraphs,
  peerSeed,
} from "./__tests__/canonicalPeerEdits";
import {
  type CanonicalCommit,
  type CanonicalSessionError,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(120_000));

/*
 * N peers, each a canonical session over the same saved document, edit
 * concurrently through the real sequencer and optimistic client; every
 * remote batch enters a session through `prepareRemote`. The edits are
 * `localEdit`'s (./__tests__/canonicalPeerEdits.ts), which also says what
 * they leave out.
 */

type Command =
  | { type: "local"; peer: number; kind: LocalKind; seed: number }
  | { type: "deliver"; seed: number }
  | { type: "submit"; peer: number };

const commandArbitrary: fc.Arbitrary<Command> = fc.oneof(
  {
    weight: 5,
    arbitrary: fc.record({
      type: fc.constant("local" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
      kind: fc.constantFrom(...LOCAL_KINDS),
      seed: fc.nat(),
    }),
  },
  {
    weight: 4,
    arbitrary: fc.record({
      type: fc.constant("deliver" as const),
      seed: fc.nat(),
    }),
  },
  {
    weight: 3,
    arbitrary: fc.record({
      type: fc.constant("submit" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
    }),
  },
);

/** How a peer applies a remote change: the session's entry, or a deliberate break. */
type RemoteStrategy = "session" | "journaledLocal" | "editorPosition";

type Client = ReturnType<typeof createClient>;
type Sequencer = ReturnType<typeof createSequencer>;
type Submission = ReturnType<Sequencer["submit"]>;

type Peer = EditingPeer & {
  actor: string;
  client: Client;
  /** A pending batch was dropped: the caret's continuity is not promised. */
  replaced: boolean;
  ops: number;
};

type Delivery =
  | { peer: number; type: "result"; result: Submission }
  | { peer: number; type: "broadcast"; batch: SequencedBatch };

type Coverage = {
  remoteApplied: number;
  undoAfterRemote: number;
  droppedSteps: number;
  rejected: number;
};

const simulate = async (
  peersCount: number,
  commands: readonly Command[],
  strategy: RemoteStrategy,
  coverage: Coverage,
): Promise<void> => {
  const bytes = await peerSeed();
  const peers: Peer[] = [];
  const parse = () =>
    parseDocx(bytes, { preloadFonts: false, detectVariables: false });
  // The room starts from the document a session reads from the saved file.
  const room: Sequencer = createSequencer(
    createCanonicalSession(await parse()).unwrap().document,
  );
  for (let index = 0; index < peersCount; index += 1) {
    const actor = `actor-${index}`;
    // Every actor mints ids only from the blocks the room leases it.
    const session = createCanonicalSession(await parse(), undefined, {
      idLease: room.grantIdLease(actor).unwrap(),
    }).unwrap();
    const home = paragraphs(session.document).at(index)?.paraId;
    if (home === undefined) return panic("The seed has a paragraph per peer.");
    const peer: Peer = {
      index,
      actor,
      home,
      session,
      client: createClient(session.document),
      state: homeState(session, home),
      history: [],
      replaced: false,
      typed: 0,
      ops: 0,
    };
    // The host mirror drops what the session drops, from the bottom.
    session.onHistoryDrop(({ undo }) => {
      coverage.droppedSteps += undo;
      peer.history.splice(0, undo);
    });
    peers.push(peer);
  }
  const batches = new Map<string, BatchMeta>();
  const deliveries: Delivery[] = [];

  /** Enqueue first: a commit the room cannot take is never published. */
  const commitLocal = (
    peer: Peer,
    prepared: Result<CanonicalCommit, CanonicalSessionError>,
    meta: BatchMeta,
  ): CanonicalCommit | undefined => {
    if (prepared.isErr()) return undefined;
    const commit = prepared.value;
    const opId = `${peer.actor}-${peer.ops++}`;
    peer.client
      .enqueue({
        schema: DOCUMENT_OP_SCHEMA_VERSION,
        opId,
        actor: peer.actor,
        baseRev: peer.client.headRev,
        ops: commit.ops,
      })
      .unwrap();
    peer.state = publishCanonicalProjection({
      state: peer.state,
      commit,
      session: peer.session,
    }).unwrap().state;
    batches.set(opId, meta);
    return commit;
  };

  const applyRemote = (peer: Peer, change: ClientDocumentChange) => {
    if (change.type === "replaced") peer.replaced = true;
    coverage.remoteApplied += 1;
    if (strategy === "journaledLocal") {
      const prepared = peer.session.prepareOps(
        peer.state,
        change.batch.ops,
        caretOf(peer),
      );
      if (prepared.isErr() && prepared.error.reason === "noChange") return;
      // The host never learns that the session took the remote batch as
      // its own step.
      peer.state = publishCanonicalProjection({
        state: peer.state,
        commit: prepared.unwrap(),
        session: peer.session,
      }).unwrap().state;
      return;
    }
    for (const step of peer.history) step.rebased = true;
    const before = peer.state.selection;
    const prepared = peer.session.prepareRemote(peer.state, change);
    if (prepared.isErr() && prepared.error.reason === "noChange") return;
    const commit = prepared.unwrap();
    peer.state = publishCanonicalProjection({
      state: peer.state,
      commit,
      session: peer.session,
    }).unwrap().state;
    if (strategy === "editorPosition")
      peer.state = peer.state.apply(
        peer.state.tr.setSelection(
          before.map(peer.state.doc, commit.transaction.mapping),
        ),
      );
  };

  const local = (peer: Peer, kind: LocalKind, choice: number) =>
    localEdit({
      peer,
      peers,
      kind,
      choice,
      commit: commitLocal,
      onUndoAfterRemote: () => {
        coverage.undoAfterRemote += 1;
      },
    });

  const submit = (peer: Peer) => {
    const batch = peer.client.nextSubmission();
    if (batch === undefined) return;
    const result = room.submit(batch);
    deliveries.push({ peer: peer.index, type: "result", result });
    if (result.type === "reject") {
      // Leased ids never collide, so no batch leaves its actor's lease.
      expect(result.reason.reason).not.toBe(
        BATCH_REJECTION_REASONS.OUTSIDE_ID_LEASE,
      );
      coverage.rejected += 1;
      return;
    }
    const broadcast = room.broadcasts.at(-1);
    if (broadcast === undefined) return;
    for (const target of peers)
      deliveries.push({
        peer: target.index,
        type: "broadcast",
        batch: broadcast,
      });
  };

  const deliver = (index: number) => {
    const event = deliveries.splice(index % deliveries.length, 1).at(0);
    if (event === undefined) return;
    const peer = peers.at(event.peer);
    if (peer === undefined) return;
    if (event.type === "broadcast") {
      for (const change of peer.client.receiveBroadcast(event.batch).unwrap())
        applyRemote(peer, change);
    } else if (event.result.type === "ack")
      peer.client.receiveAck(event.result);
    else
      for (const change of peer.client.receiveReject(event.result))
        applyRemote(peer, change);
  };

  for (const command of commands) {
    switch (command.type) {
      case "local": {
        const peer = peers.at(command.peer % peersCount);
        if (peer !== undefined) local(peer, command.kind, command.seed);
        break;
      }
      case "submit": {
        const peer = peers.at(command.peer % peersCount);
        if (peer !== undefined) submit(peer);
        break;
      }
      case "deliver":
        if (deliveries.length > 0) deliver(command.seed);
        break;
      default: {
        const exhaustive: never = command;
        return exhaustive;
      }
    }
  }
  for (let steps = 0; ; steps += 1) {
    expect(steps).toBeLessThan(10_000);
    for (const peer of peers) submit(peer);
    if (deliveries.length === 0) break;
    deliver(0);
  }

  const finalText = documentText(room.document);
  for (const peer of peers) {
    expect(peer.client.pending).toHaveLength(0);
    expect(peer.client.headRev).toBe(room.headRev);
    expect(documentText(peer.session.document)).toBe(finalText);
  }
  await expectIdenticalSaves(peers.map(({ session }) => session));
  // The caret followed each peer's own typing across every remote edit.
  expectSequencedKeystrokes({
    finalText,
    journal: room.broadcasts,
    batches,
    peers: peersCount,
    contiguous: (index) => peers.at(index)?.replaced === false,
  });
};

const law = (strategy: RemoteStrategy, coverage: Coverage) =>
  fc.asyncProperty(
    fc.integer({ min: 2, max: MAX_PEERS }),
    fc.array(commandArbitrary, { minLength: 10, maxLength: 60 }),
    async (peersCount, commands) =>
      simulate(peersCount, commands, strategy, coverage),
  );

/**
 * Peer 0 types, peer 1 types and submits, peer 0 receives that batch (the
 * second delivery: the submitter's result comes first) and undoes.
 */
const UNDO_AFTER_REMOTE: readonly Command[] = [
  { type: "local", peer: 0, kind: "type", seed: 0 },
  { type: "local", peer: 1, kind: "type", seed: 0 },
  { type: "submit", peer: 1 },
  { type: "deliver", seed: 1 },
  { type: "local", peer: 0, kind: "undo", seed: 0 },
];

/**
 * Peer 0 types, peer 1 puts a digit at the start of peer 0's paragraph,
 * peer 0 receives it and types on: the caret must still follow its text.
 */
const TYPE_AFTER_REMOTE_EDIT: readonly Command[] = [
  { type: "local", peer: 0, kind: "type", seed: 0 },
  { type: "local", peer: 1, kind: "prefix", seed: 0 },
  { type: "submit", peer: 1 },
  { type: "deliver", seed: 1 },
  { type: "local", peer: 0, kind: "type", seed: 0 },
];
const EXAMPLES: [number, Command[]][] = [
  [2, [...UNDO_AFTER_REMOTE]],
  [2, [...TYPE_AFTER_REMOTE_EDIT]],
];

const emptyCoverage = (): Coverage => ({
  remoteApplied: 0,
  undoAfterRemote: 0,
  droppedSteps: 0,
  rejected: 0,
});

test("peers applying remote batches through the session converge and keep their own history", async () => {
  await assertProperty(law("session", emptyCoverage()), {
    numRuns: 40,
    examples: EXAMPLES,
    id: "canonical remote apply convergence",
  });
  // On a fixed sample: remote batches arrive and undo follows them.
  const coverage = emptyCoverage();
  await checkCoverageSample(law("session", coverage), 15, {
    map: ([peersCount, commands]) => [
      peersCount,
      [...UNDO_AFTER_REMOTE, ...commands],
    ],
  });
  expect(coverage.remoteApplied).toBeGreaterThan(0);
  expect(coverage.undoAfterRemote).toBeGreaterThan(0);
});

test("the property catches a remote batch journaled as a local undo step", async () => {
  const result = await fc.check(
    law("journaledLocal", emptyCoverage()),
    propertyConfig({ numRuns: 100, endOnFailure: true, examples: EXAMPLES }),
  );
  expect(result.failed).toBe(true);
});

test("the property catches a selection mapped by editor position", async () => {
  const result = await fc.check(
    law("editorPosition", emptyCoverage()),
    propertyConfig({ numRuns: 100, endOnFailure: true, examples: EXAMPLES }),
  );
  expect(result.failed).toBe(true);
});
