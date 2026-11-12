import { panic, Result } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { parseDocx } from "../docx/parser";
import {
  type BatchMeta,
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
  type CanonicalCollaboration,
  type CanonicalCollaborationEvent,
  type CanonicalCollaborationFault,
  connectCanonicalRoom,
} from "./canonicalCollaboration";
import {
  type CanonicalMemoryConnection,
  canonicalSnapshotId,
  createCanonicalMemoryRoom,
} from "./canonicalRoomTransport";
import {
  type CanonicalCommit,
  type CanonicalSessionError,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(120_000));

/*
 * N peers (3 or 4), each a canonical session joined to one in-memory room
 * through the collaboration controller, edit concurrently with `localEdit`'s
 * edits (./__tests__/canonicalPeerEdits.ts says what they leave out). Every
 * hop between a peer and the room (a submission reaching the authority, its
 * answer, a sequenced batch reaching a subscriber) waits in one network
 * queue, so deliveries are delayed and reordered; a hop may also run twice,
 * and a peer may lose its connection (its hops in flight are dropped) and get
 * it back. The last peer may join late, replaying the journal from rev 0.
 * Small lease blocks make the peers renew their id leases.
 *
 * Checked once the network is drained: every peer's text and paragraph ids
 * equal the authority's, every peer saves the same bytes, every sequenced
 * keystroke (less the undone ones) survives, every local edit was either
 * sequenced or reported to the host as refused, and no peer stopped
 * following the room.
 */

type Command =
  | { type: "local"; peer: number; kind: LocalKind; seed: number }
  | { type: "deliver"; seed: number }
  | { type: "repeat"; seed: number }
  | { type: "disconnect"; peer: number }
  | { type: "reconnect"; peer: number }
  | { type: "join" };

const commandArbitrary: fc.Arbitrary<Command> = fc.oneof(
  {
    weight: 6,
    arbitrary: fc.record({
      type: fc.constant("local" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
      kind: fc.constantFrom(...LOCAL_KINDS),
      seed: fc.nat(),
    }),
  },
  {
    weight: 6,
    arbitrary: fc.record({
      type: fc.constant("deliver" as const),
      seed: fc.nat(),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("repeat" as const),
      seed: fc.nat(),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("disconnect" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
    }),
  },
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("reconnect" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
    }),
  },
  { weight: 1, arbitrary: fc.constant({ type: "join" as const }) },
);

type Peer = EditingPeer & {
  actor: string;
  connection: CanonicalMemoryConnection;
  collaboration: CanonicalCollaboration;
  refused: Set<string>;
  events: CanonicalCollaborationEvent[];
  /** A pending batch was dropped: the caret's continuity is not promised. */
  replaced: boolean;
};

type Coverage = {
  remoteApplied: number;
  undoAfterRemote: number;
  refused: number;
  resubmitted: number;
  lateJoins: number;
  renewals: number;
};

/** Let every promise continuation run: the controller's answers and leases. */
const settle = () => new Promise<void>((resolve) => setImmediate(resolve));

const PARAGRAPH_BLOCK = 8;
const BLOCK_SIZES = {
  paragraph: PARAGRAPH_BLOCK,
  revision: 64,
  control: 16,
  comment: 16,
} as const;

const simulate = async (
  peersCount: number,
  lateJoin: boolean,
  commands: readonly Command[],
  fault: CanonicalCollaborationFault | undefined,
  coverage: Coverage,
): Promise<void> => {
  const bytes = await peerSeed();
  const parse = () =>
    parseDocx(bytes, { preloadFonts: false, detectVariables: false });
  const snapshot = await canonicalSnapshotId(bytes);
  const hops: (() => void)[] = [];
  // The room starts from the document a session reads from the saved file.
  const room = createCanonicalMemoryRoom({
    document: createCanonicalSession(await parse()).unwrap().document,
    snapshot,
    sequencer: {
      idLeases: { blockSizes: BLOCK_SIZES, grants: [], minted: [] },
    },
    schedule: (hop) => hops.push(hop),
  });
  const peers: Peer[] = [];
  const batches = new Map<string, BatchMeta>();
  const localOpIds = new Set<string>();
  let nextMeta: BatchMeta | undefined;

  const join = async (index: number): Promise<void> => {
    const actor = `actor-${index}`;
    const connection = room.connect();
    const document = await parse();
    const refused = new Set<string>();
    const events: CanonicalCollaborationEvent[] = [];
    let ops = 0;
    // The peer once joined; until then, the state the editor shows.
    const self: { peer?: Peer; state?: EditingPeer["state"] } = {};
    const joined = await connectCanonicalRoom(
      {
        transport: connection,
        actor,
        snapshot,
        renewBelow: PARAGRAPH_BLOCK / 2,
        onEvent: (event) => {
          events.push(event);
          if (event.type === "localEditRefused") {
            refused.add(event.opId);
            coverage.refused += 1;
            if (self.peer !== undefined) self.peer.replaced = true;
          }
        },
        open: (idLease) =>
          createCanonicalSession(document, undefined, {
            idLease,
            newOpId: () => `${actor}-${ops++}`,
          }).map((session) => {
            const home = paragraphs(session.document).at(index)?.paraId;
            if (home === undefined)
              return panic("The seed has a paragraph per peer.");
            self.state = homeState(session, home);
            return {
              session,
              editor: {
                state: () =>
                  self.peer?.state ?? self.state ?? panic("No state."),
                publish: (commit: CanonicalCommit) => {
                  const peer = self.peer ?? panic("No peer.");
                  const published = publishCanonicalProjection({
                    state: peer.state,
                    commit,
                    session,
                  });
                  if (published.isErr()) return published;
                  peer.state = published.value.state;
                  if (commit.origin === "remote") coverage.remoteApplied += 1;
                  return Result.ok(undefined);
                },
              },
            };
          }),
      },
      fault,
    );
    const collaboration = joined.unwrap();
    const { session } = collaboration;
    const home = paragraphs(session.document).at(index)?.paraId;
    if (home === undefined || self.state === undefined)
      return panic("The peer opened its session.");
    const peer: Peer = {
      index,
      actor,
      home,
      session,
      state: self.state,
      history: [],
      typed: 0,
      connection,
      collaboration,
      refused,
      events,
      replaced: false,
    };
    self.peer = peer;
    const joinedPeer = peer;
    session.onCommitted((committed) => {
      if (committed.origin === "remote") {
        for (const step of joinedPeer.history) step.rebased = true;
        return;
      }
      if (nextMeta === undefined) return panic("A commit nobody made.");
      batches.set(committed.opId, nextMeta);
      localOpIds.add(committed.opId);
    });
    // The host mirror drops what the session drops, from the bottom. A
    // dropped step may have been a pending batch's: the caret's continuity
    // is no longer promised.
    session.onHistoryDrop(({ undo }) => {
      joinedPeer.history.splice(0, undo);
      joinedPeer.replaced = true;
    });
    peers[index] = peer;
  };

  const commitLocal = (
    peer: Peer,
    prepared: Result<CanonicalCommit, CanonicalSessionError>,
    meta: BatchMeta,
  ): CanonicalCommit | undefined => {
    if (prepared.isErr()) return undefined;
    nextMeta = meta;
    // The controller enqueues the commit from the session's feed.
    peer.state = publishCanonicalProjection({
      state: peer.state,
      commit: prepared.value,
      session: peer.session,
    }).unwrap().state;
    nextMeta = undefined;
    return prepared.value;
  };

  const joinedAtStart = lateJoin ? peersCount - 1 : peersCount;
  for (let index = 0; index < joinedAtStart; index += 1) await join(index);
  await settle();

  for (const command of commands) {
    switch (command.type) {
      case "local": {
        const peer = peers.at(command.peer % peersCount);
        if (peer === undefined) break;
        localEdit({
          peer,
          peers: peers.filter((candidate) => candidate !== undefined),
          kind: command.kind,
          choice: command.seed,
          commit: commitLocal,
          onUndoAfterRemote: () => {
            coverage.undoAfterRemote += 1;
          },
        });
        break;
      }
      case "deliver":
      case "repeat": {
        if (hops.length === 0) break;
        const index = command.seed % hops.length;
        const hop =
          command.type === "deliver" ? hops.splice(index, 1)[0] : hops[index];
        hop?.();
        break;
      }
      case "disconnect":
        peers.at(command.peer % peersCount)?.connection.disconnect();
        break;
      case "reconnect": {
        const peer = peers.at(command.peer % peersCount);
        if (peer === undefined || peer.connection.connected) break;
        if (peer.collaboration.pending.length > 0) coverage.resubmitted += 1;
        peer.connection.reconnect();
        break;
      }
      case "join":
        if (peers.length < peersCount) {
          coverage.lateJoins += 1;
          await join(peersCount - 1);
        }
        break;
      default: {
        const exhaustive: never = command;
        return exhaustive;
      }
    }
    await settle();
  }
  if (peers.length < peersCount) await join(peersCount - 1);
  for (const peer of peers) peer.connection.reconnect();
  await settle();
  for (let steps = 0; hops.length > 0; steps += 1) {
    expect(steps).toBeLessThan(20_000);
    hops.shift()?.();
    await settle();
  }

  const finalText = documentText(room.document);
  const finalIds = paragraphs(room.document).map(({ paraId }) => paraId);
  for (const peer of peers) {
    expect(peer.events.filter(({ type }) => type === "stopped")).toEqual([]);
    expect(peer.collaboration.pending).toHaveLength(0);
    expect(peer.collaboration.headRev).toBe(room.headRev);
    expect(documentText(peer.session.document)).toBe(finalText);
    expect(
      paragraphs(peer.session.document).map(({ paraId }) => paraId),
    ).toEqual(finalIds);
  }
  coverage.renewals +=
    room.idLeases.grants.filter(({ block }) => block.kind === "paragraph")
      .length - peers.length;
  await expectIdenticalSaves(peers.map(({ session }) => session));
  expectSequencedKeystrokes({
    finalText,
    journal: room.journal,
    batches,
    peers: peersCount,
    contiguous: (index) => peers.at(index)?.replaced === false,
  });
  // No silent loss: a local edit the room did not sequence was reported.
  const sequenced = new Set(room.journal.map(({ opId }) => opId));
  const reported = new Set<string>();
  for (const { refused } of peers)
    for (const opId of refused) reported.add(opId);
  for (const opId of localOpIds)
    expect(sequenced.has(opId) || reported.has(opId)).toBe(true);
};

const law = (
  fault: CanonicalCollaborationFault | undefined,
  coverage: Coverage,
) =>
  fc.asyncProperty(
    fc.integer({ min: 3, max: MAX_PEERS }),
    fc.boolean(),
    fc.array(commandArbitrary, { minLength: 10, maxLength: 50 }),
    async (peersCount, lateJoin, commands) =>
      simulate(peersCount, lateJoin, commands, fault, coverage),
  );

/**
 * Peer 0 puts a digit at the start of peer 1's paragraph, peer 1 types at
 * its caret; peer 1's batch reaches the room first and its broadcast reaches
 * peer 0 while peer 0's digit is still pending, so the typed character must
 * land one place further on.
 */
const REMOTE_OVER_PENDING: readonly Command[] = [
  { type: "local", peer: 0, kind: "prefix", seed: 0 },
  { type: "local", peer: 1, kind: "type", seed: 0 },
  { type: "deliver", seed: 1 },
  { type: "deliver", seed: 1 },
];

/**
 * Peer 0 types; the room sequences it, then peer 0 loses the connection
 * before the answer and the broadcast reach it, and reconnects.
 */
const ANSWER_LOST: readonly Command[] = [
  { type: "local", peer: 0, kind: "type", seed: 0 },
  { type: "deliver", seed: 0 },
  { type: "disconnect", peer: 0 },
  { type: "reconnect", peer: 0 },
];

/** A late peer joins after two sequenced edits and replays them from rev 0. */
const LATE_JOIN: readonly Command[] = [
  { type: "local", peer: 0, kind: "type", seed: 0 },
  { type: "local", peer: 1, kind: "splitEnd", seed: 0 },
  { type: "deliver", seed: 0 },
  { type: "deliver", seed: 0 },
  { type: "join" },
  { type: "local", peer: 3, kind: "type", seed: 0 },
];

const EXAMPLES: [number, boolean, Command[]][] = [
  [3, false, [...REMOTE_OVER_PENDING]],
  [3, false, [...ANSWER_LOST]],
  [4, true, [...LATE_JOIN]],
];

const emptyCoverage = (): Coverage => ({
  remoteApplied: 0,
  undoAfterRemote: 0,
  refused: 0,
  resubmitted: 0,
  lateJoins: 0,
  renewals: 0,
});

test("peers joined to a room through the controller converge on the authority's document", async () => {
  await assertProperty(law(undefined, emptyCoverage()), {
    numRuns: 25,
    examples: EXAMPLES,
    id: "canonical collaboration convergence",
  });
  // On a fixed sample: remote edits arrive, undo follows them, a peer
  // resubmits after reconnecting, a late peer joins and leases are renewed.
  const coverage = emptyCoverage();
  await checkCoverageSample(law(undefined, coverage), 12, {
    map: ([peersCount, lateJoin, commands], index) => {
      const example = EXAMPLES[index % EXAMPLES.length];
      return example === undefined
        ? [peersCount, lateJoin, commands]
        : [example[0], example[1], [...example[2], ...commands]];
    },
  });
  expect(coverage.remoteApplied).toBeGreaterThan(0);
  expect(coverage.undoAfterRemote).toBeGreaterThan(0);
  expect(coverage.resubmitted).toBeGreaterThan(0);
  expect(coverage.lateJoins).toBeGreaterThan(0);
  expect(coverage.renewals).toBeGreaterThan(0);
});

test("the property catches remote batches applied without the pending rebase", async () => {
  const result = await fc.check(
    law("unrebasedRemote", emptyCoverage()),
    propertyConfig({ numRuns: 100, endOnFailure: true, examples: EXAMPLES }),
  );
  expect(result.failed).toBe(true);
});

test("the property catches a resubmission under a fresh operation id", async () => {
  const result = await fc.check(
    law("freshOpIdOnResubmit", emptyCoverage()),
    propertyConfig({ numRuns: 100, endOnFailure: true, examples: EXAMPLES }),
  );
  expect(result.failed).toBe(true);
});
