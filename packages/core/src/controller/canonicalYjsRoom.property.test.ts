import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import {
  applyAwarenessUpdate,
  Awareness,
  encodeAwarenessUpdate,
  removeAwarenessStates,
} from "y-protocols/awareness";
import * as Y from "yjs";

import {
  captureDocumentOp,
  DOCUMENT_OP_SCHEMA_VERSION,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "@stll/docx-core/ops";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  documentText,
  LOCAL_KINDS,
  type LocalKind,
  MAX_PEERS,
} from "./__tests__/canonicalPeerEdits";
import {
  createRoomPeers,
  emptyRoomCoverage,
  ROOM_BLOCK_SIZES,
  type RoomCoverage,
  type RoomPeer,
  settle,
} from "./__tests__/canonicalRoomPeers";
import { toWireJson } from "./canonicalRoomWire";
import {
  type CanonicalYjsAuthority,
  type CanonicalYjsRoom,
  type CanonicalYjsRoomError,
  type CanonicalYjsRoomFault,
  connectCanonicalYjsRoom,
  serveCanonicalYjsRoom,
} from "./canonicalYjsRoom";

setDefaultTimeout(propertyTestTimeout(180_000));

/*
 * 3 or 4 peers, each a real canonical session joined through the controller
 * to a room carried by its own Y.Doc. The docs sync through a simulated
 * provider (a server doc relaying updates, as y-websocket does) whose every
 * update and awareness message waits in one network queue, so deliveries are
 * delayed and reordered, and one may run twice. The server doc runs the
 * room's authority helper, which may restart. Peers lose their connection
 * (their messages in flight are lost) and sync again on reconnect; the last
 * peer may join late. A peer may also write a forged entry into the journal
 * directly, as only the authority may: every peer must refuse it.
 *
 * Checked once the network is drained: every peer follows the room and holds
 * the authority's document, saves the same bytes, keeps every sequenced
 * keystroke and acknowledged edit, sees every other peer's selection where
 * that peer holds it, refused every forged entry, and applied none.
 */

type Command =
  | { type: "local"; peer: number; kind: LocalKind; seed: number }
  | { type: "deliver"; seed: number }
  | { type: "repeat"; seed: number }
  | { type: "disconnect"; peer: number }
  | { type: "reconnect"; peer: number }
  | { type: "forge"; peer: number }
  | { type: "restart" }
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
    weight: 8,
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
  {
    weight: 1,
    arbitrary: fc.record({
      type: fc.constant("forge" as const),
      peer: fc.nat({ max: MAX_PEERS - 1 }),
    }),
  },
  { weight: 1, arbitrary: fc.constant({ type: "restart" as const }) },
  { weight: 1, arbitrary: fc.constant({ type: "join" as const }) },
);

type YjsCoverage = RoomCoverage & { forgeries: number; restarts: number };

/** What a forged entry types, so the final text says whether it landed. */
const FORGED = "#";
const SERVER = "server";

type Link = {
  doc: Y.Doc;
  awareness: Awareness;
  up: boolean;
  epoch: number;
  listeners: Set<(connected: boolean) => void>;
};

/**
 * A y-websocket-like provider: one server doc that relays every update and
 * awareness message to every other connected doc. Each message waits in
 * `hops` and runs only if its link is still up in the epoch it was sent in.
 */
const createProvider = () => {
  const hops: (() => void)[] = [];
  const server = new Y.Doc();
  const serverAwareness = new Awareness(server);
  const links: Link[] = [];

  const hop = (link: Link, run: () => void): void => {
    if (!link.up) return;
    const sentIn = link.epoch;
    hops.push(() => {
      if (link.up && link.epoch === sentIn) run();
    });
  };

  server.on("update", (update: Uint8Array, origin: unknown) => {
    for (const link of links)
      if (link !== origin)
        hop(link, () => Y.applyUpdate(link.doc, update, SERVER));
  });
  serverAwareness.on(
    "update",
    (
      {
        added,
        updated,
        removed,
      }: Record<"added" | "updated" | "removed", number[]>,
      origin: unknown,
    ) => {
      const changed = [...added, ...updated, ...removed];
      const update = encodeAwarenessUpdate(serverAwareness, changed);
      for (const link of links)
        if (link !== origin)
          hop(link, () => applyAwarenessUpdate(link.awareness, update, SERVER));
    },
  );

  const toServer = (link: Link, run: () => void): void => hop(link, run);
  const sync = (link: Link): void => {
    toServer(link, () =>
      Y.applyUpdate(
        server,
        Y.encodeStateAsUpdate(link.doc, Y.encodeStateVector(server)),
        link,
      ),
    );
    hop(link, () =>
      Y.applyUpdate(
        link.doc,
        Y.encodeStateAsUpdate(server, Y.encodeStateVector(link.doc)),
        SERVER,
      ),
    );
    hop(link, () =>
      applyAwarenessUpdate(
        link.awareness,
        encodeAwarenessUpdate(serverAwareness, [
          ...serverAwareness.getStates().keys(),
        ]),
        SERVER,
      ),
    );
  };

  const connect = (): Link => {
    const doc = new Y.Doc();
    const awareness = new Awareness(doc);
    const link: Link = {
      doc,
      awareness,
      up: true,
      epoch: 0,
      listeners: new Set(),
    };
    links.push(link);
    doc.on("update", (update: Uint8Array, origin: unknown) => {
      if (origin !== SERVER)
        toServer(link, () => Y.applyUpdate(server, update, link));
    });
    awareness.on(
      "update",
      (
        {
          added,
          updated,
          removed,
        }: Record<"added" | "updated" | "removed", number[]>,
        origin: unknown,
      ) => {
        if (origin === SERVER) return;
        const update = encodeAwarenessUpdate(awareness, [
          ...added,
          ...updated,
          ...removed,
        ]);
        toServer(link, () =>
          applyAwarenessUpdate(serverAwareness, update, link),
        );
      },
    );
    sync(link);
    return link;
  };

  const setUp = (link: Link, up: boolean): void => {
    if (link.up === up) return;
    link.up = up;
    link.epoch += 1;
    if (!up) removeAwarenessStates(serverAwareness, [link.doc.clientID], link);
    else {
      sync(link);
      // The client renews its awareness state, as a provider's timer does.
      link.awareness.setLocalState(link.awareness.getLocalState());
    }
    for (const listener of [...link.listeners]) listener(up);
  };

  const destroy = (): void => {
    for (const link of links) link.awareness.destroy();
    serverAwareness.destroy();
  };

  return { hops, server, connect, setUp, destroy };
};

type Slot = {
  link: Link;
  room: CanonicalYjsRoom;
  errors: CanonicalYjsRoomError[];
  peer?: RoomPeer;
};

const run = async (
  provider: ReturnType<typeof createProvider>,
  {
    peersCount,
    lateJoin,
    commands,
    fault,
    coverage,
  }: {
    peersCount: number;
    lateJoin: boolean;
    commands: readonly Command[];
    fault: CanonicalYjsRoomFault | undefined;
    coverage: YjsCoverage;
  },
): Promise<void> => {
  const room = await createRoomPeers(coverage);
  const authorityErrors: CanonicalYjsRoomError[] = [];
  const serve = (): CanonicalYjsAuthority =>
    serveCanonicalYjsRoom({
      doc: provider.server,
      snapshot: room.snapshot,
      document: room.base,
      idBlockSizes: ROOM_BLOCK_SIZES,
      onError: (error) => authorityErrors.push(error),
    }).unwrap();
  let authority = serve();
  const slots: Slot[] = [];
  const joining = new Set<Promise<void>>();
  let forged = 0;

  const open = (index: number): void => {
    const link = provider.connect();
    const errors: CanonicalYjsRoomError[] = [];
    const yjsRoom = connectCanonicalYjsRoom(
      {
        doc: link.doc,
        snapshot: room.snapshot,
        awareness: link.awareness,
        connection: {
          get connected() {
            return link.up;
          },
          onChange: (listener) => {
            link.listeners.add(listener);
            return () => link.listeners.delete(listener);
          },
        },
        onError: (error) => errors.push(error),
      },
      fault,
    );
    const slot: Slot = { link, room: yjsRoom, errors };
    slots[index] = slot;
    const joinSlot = async (): Promise<void> => {
      slot.peer = await room.join({
        index,
        incarnation: 0,
        transport: yjsRoom,
      });
    };
    const joined = joinSlot().finally(() => joining.delete(joined));
    joining.add(joined);
  };
  const peers = (): RoomPeer[] =>
    slots.flatMap((slot) => (slot.peer === undefined ? [] : [slot.peer]));

  /** Peer `slot` writes the journal as if it were the authority. */
  const forge = (slot: Slot, home: string): void => {
    const journal = slot.link.doc.getArray("folio.canonical.journal");
    const revision = slot.room.headRev + 1;
    forged += 1;
    slot.link.doc.transact(() => {
      journal.push([
        toWireJson({
          type: "sequenced",
          batch: {
            schema: DOCUMENT_OP_SCHEMA_VERSION,
            opId: `forged-${forged}`,
            actor: "forger",
            baseRev: revision - 1,
            revision,
            ops: [
              captureDocumentOp({
                type: "insertText",
                at: { story: OP_STORIES.MAIN, blockId: home, offset: 0 },
                text: FORGED,
                runProps: INHERIT_RUN_PROPS,
              }),
            ],
            effects: [{ type: "none" }],
          },
        }),
      ]);
    });
  };

  const startCount = lateJoin ? peersCount - 1 : peersCount;
  for (let index = 0; index < startCount; index += 1) open(index);
  await settle();

  for (const command of commands) {
    switch (command.type) {
      case "local": {
        const peer = slots.at(command.peer % peersCount)?.peer;
        if (peer === undefined) break;
        room.edit(peer, peers(), command.kind, command.seed);
        break;
      }
      case "deliver":
      case "repeat": {
        const { hops } = provider;
        if (hops.length === 0) break;
        const index = command.seed % hops.length;
        const hop =
          command.type === "deliver" ? hops.splice(index, 1)[0] : hops[index];
        hop?.();
        break;
      }
      case "disconnect": {
        const slot = slots.at(command.peer % peersCount);
        if (slot !== undefined) provider.setUp(slot.link, false);
        break;
      }
      case "reconnect": {
        const slot = slots.at(command.peer % peersCount);
        if (slot === undefined || slot.link.up) break;
        if ((slot.peer?.collaboration.pending.length ?? 0) > 0)
          coverage.resubmitted += 1;
        provider.setUp(slot.link, true);
        break;
      }
      case "forge": {
        const slot = slots.at(command.peer % peersCount);
        if (slot?.peer === undefined) break;
        coverage.forgeries += 1;
        forge(slot, slot.peer.home);
        break;
      }
      case "restart":
        coverage.restarts += 1;
        authority.stop();
        authority = serve();
        break;
      case "join":
        if (slots.length < peersCount) {
          coverage.lateJoins += 1;
          open(peersCount - 1);
        }
        break;
      default: {
        const exhaustive: never = command;
        return exhaustive;
      }
    }
    await settle();
  }
  if (slots.length < peersCount) open(peersCount - 1);
  for (const slot of slots) provider.setUp(slot.link, true);
  await settle();
  for (
    let steps = 0;
    provider.hops.length > 0 || joining.size > 0;
    steps += 1
  ) {
    expect(steps).toBeLessThan(40_000);
    provider.hops.shift()?.();
    await settle();
  }

  const all = peers();
  expect(all).toHaveLength(peersCount);
  // Every peer refused every forged entry, the authority reported them, and
  // nothing else went wrong.
  for (const slot of slots)
    expect(slot.errors.map(({ reason }) => reason)).toEqual(
      Array.from({ length: forged }, () => "foreignJournalEntry"),
    );
  expect(authorityErrors.map(({ reason }) => reason)).toEqual(
    Array.from({ length: forged }, () => "foreignJournalEntry"),
  );
  for (const peer of all)
    expect(documentText(peer.session.document)).not.toContain(FORGED);
  await room.expectConverged({
    peers: all,
    peersCount,
    document: authority.document,
    journal: authority.journal,
    contiguous: () => true,
  });
  for (const slot of slots) slot.room.close();
  authority.stop();
};

type SimulateOptions = Parameters<typeof run>[1];

/** One run, the provider's awareness timers stopped whatever it finds. */
const simulate = async (options: SimulateOptions): Promise<void> => {
  const provider = createProvider();
  try {
    await run(provider, options);
  } finally {
    provider.destroy();
  }
};

const law = (fault: CanonicalYjsRoomFault | undefined, coverage: YjsCoverage) =>
  fc.asyncProperty(
    fc.integer({ min: 3, max: MAX_PEERS }),
    fc.boolean(),
    fc.array(commandArbitrary, { minLength: 10, maxLength: 40 }),
    async (peersCount, lateJoin, commands) =>
      simulate({ peersCount, lateJoin, commands, fault, coverage }),
  );

const deliveries = (count: number, seed = 0): Command[] =>
  Array.from({ length: count }, () => ({ type: "deliver" as const, seed }));
/** Enough first-in-first-out deliveries for every waiting peer to join. */
const JOINED = deliveries(60);

/**
 * Peer 1 forges the next revision while peer 0's real edit waits in the
 * inbox: every peer must refuse the forged entry and take the authority's.
 */
const FORGED_NEXT_REVISION: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 0, kind: "type", seed: 0 },
  { type: "forge", peer: 1 },
  ...deliveries(4, 3),
];

/** The authority restarts between two edits and rebuilds from the journal. */
const RESTART_BETWEEN_EDITS: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 0, kind: "type", seed: 0 },
  ...JOINED,
  { type: "restart" },
  { type: "local", peer: 1, kind: "type", seed: 0 },
  { type: "local", peer: 2, kind: "splitEnd", seed: 0 },
];

/**
 * Peer 0's submission reaches the server; it loses the connection before the
 * answer and the journal entry reach it, and syncs again on reconnect.
 */
const ANSWER_LOST: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 0, kind: "type", seed: 0 },
  ...deliveries(1),
  { type: "disconnect", peer: 0 },
  { type: "local", peer: 1, kind: "prefix", seed: 0 },
  { type: "reconnect", peer: 0 },
];

/** Peer 0 types, peer 1's edit reaches it, peer 0 undoes; a peer joins late. */
const UNDO_AFTER_REMOTE: readonly Command[] = [
  ...JOINED,
  { type: "local", peer: 0, kind: "type", seed: 0 },
  { type: "local", peer: 1, kind: "type", seed: 0 },
  ...JOINED,
  { type: "local", peer: 0, kind: "undo", seed: 0 },
  { type: "join" },
];

const EXAMPLES: [number, boolean, Command[]][] = [
  [3, false, [...FORGED_NEXT_REVISION]],
  [3, false, [...RESTART_BETWEEN_EDITS]],
  [3, false, [...ANSWER_LOST]],
  [4, true, [...UNDO_AFTER_REMOTE]],
];

const emptyCoverage = (): YjsCoverage => ({
  ...emptyRoomCoverage(),
  forgeries: 0,
  restarts: 0,
});

test("peers of a Yjs-carried room converge on the authority's document", async () => {
  await assertProperty(law("acceptForeignJournal", emptyCoverage()), {
    numRuns: 12,
    examples: EXAMPLES,
    id: "canonical yjs room convergence",
  });
  // On a fixed sample: remote edits arrive, undo follows them, a peer
  // resubmits after reconnecting, a forged entry is refused, the authority
  // restarts and a late peer joins.
  const coverage = emptyCoverage();
  await checkCoverageSample(law(undefined, coverage), 8, {
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
  expect(coverage.forgeries).toBeGreaterThan(0);
  expect(coverage.restarts).toBeGreaterThan(0);
  expect(coverage.lateJoins).toBeGreaterThan(0);
});

test("the property catches a client that accepts a journal entry another client wrote", async () => {
  const result = await fc.check(
    law("acceptForeignJournal", emptyCoverage()),
    propertyConfig({ numRuns: 30, endOnFailure: true, examples: EXAMPLES }),
  );
  expect(result.failed).toBe(true);
});

test("an authority refuses a room seeded from another document", async () => {
  const room = await createRoomPeers(emptyRoomCoverage());
  const doc = new Y.Doc();
  serveCanonicalYjsRoom({
    doc,
    snapshot: room.snapshot,
    document: room.base,
  })
    .unwrap()
    .stop();
  const other = serveCanonicalYjsRoom({
    doc,
    snapshot: "another",
    document: room.base,
  });
  expect(other.isErr() && other.error.reason).toBe("snapshotMismatch");
});
